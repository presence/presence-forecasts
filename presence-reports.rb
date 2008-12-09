#!/usr/bin/env ruby

#Require the appropriate libraries
require 'lib/libs.rb'
require 'lib/logger.rb'
require 'activesupport'

#Load our configuration file
$config = YAML.load_file("config.yml")

if $config["days_to_add"] != nil
  process_time = Time.now + $config["days_to_add"].days
else
  process_time = Time.now
end

#First lets set the name of the quarter we are in
case process_time.at_beginning_of_quarter.month
  when 1
    this_quarter = "Q1"
  when 4
    this_quarter = "Q2"
  when 7
    this_quarter = "Q3"
  when 10
    this_quarter = "Q4"
end

#Build an output array containing the hashes of each record returned by the SOAP web service of Sugar CRM
def collect_output results
  output = []
  for entry in results.entry_list
    item = {}
    for name_value in entry.name_value_list
      item[name_value.name]=name_value.value
    end
    output << item
  end
  return output
end

#Rate source initialization using http://www.xe.com
begin
  provider = Currency::Exchange::Rate::Source::Xe.new
  deriver = Currency::Exchange::Rate::Deriver.new(:source => provider)
  cache = Currency::Exchange::Rate::Source::TimedCache.new(:source => deriver)
  Currency::Exchange::Rate::Source.default = cache
rescue => err
  puts err
  exit
end
      
#Build our credentials hash to be passed to the SOAP factory and converted to XML to pass to Sugar CRM
credentials = { "user_name" => $config["username"], "password" => Digest::MD5.hexdigest($config["password"]) }
begin
  #Connect to the Sugar CRM WSDL and build our methods in Ruby
  ws_proxy = SOAP::WSDLDriverFactory.new($config["wsdl_url"]).create_rpc_driver
  ws_proxy.streamhandler.client.receive_timeout = 3600
  
  #This may be toggled on to log XML requests/responses for debugging
  #ws_proxy.wiredump_file_base = "soap"
  
  #Login to Sugar CRM
  session = ws_proxy.login(credentials, nil)
rescue => err
  puts err
  exit
end

#Check to see we got logged in properly
if session.error.number.to_i != 0
  puts session.error.description + " (" + session.error.number + ")"
  puts "Exiting"
  exit
else
  puts "Successfully logged in"
end

#Build our query for leads
module_name = "Opportunities"
query = "opportunities.sales_stage != 'Closed Lost'" #{}" && opportunities.date_closed <= #{process_time.at_end_of_quarter}" # gets all the acounts, you can also use SQL like "accounts.name like '%company%'"
order_by = "opportunities.amount" # in default order. you can also use SQL like "accounts.name"
offset = 0 # I guess this is like the SQL offset
select_fields = [] #could be ['name','industry']
max_results = "10000" # if set to 0 or "", this doesn't return all the results, like you'd expect, set to 100 as do not expect more, and times out with too many
deleted = 0 # whether you want to retrieve deleted records, too, we don't want to

#Query the SOAP WS of Sugar CRM for the Leads that we are interested in
begin
  results = ws_proxy.get_entry_list(session['id'], module_name, query, order_by, offset, select_fields, max_results, deleted)
rescue => err
  puts err
  exit
rescue Timeout::Error => err
  puts err
  exit
end

#Organize the results into a nice array of hashes to be output into our reports
leads = collect_output(results)

#Hash to track the totals of the forecast
totals = { "grand" => Currency::Money('0', :USD),
           "closed" => Currency::Money('0', :USD),
           "commit" => Currency::Money('0', :USD),
           "upside" => Currency::Money('0', :USD),
           "swing" => Currency::Money('0', :USD),
           "undetermined" => Currency::Money('0', :USD) }
#Now lets qualify which leads we should be using in the forecast
forecast_report_table = Table(%w[Account Opportunity Owner Probability NextStep ExpectedClose Amount])
#Create totals table
totals_report_table = Table(%w[Currency Closed Commit Upside Swing Undetermined GrandTotal])

leads.each do |lead|
  #First, we make sure they are in the current quarter
  if lead["date_closed"].to_date >= process_time.at_beginning_of_quarter.to_date && lead["date_closed"].to_date <= process_time.at_end_of_quarter.to_date
    #Then we make sure it is in the US team
    if lead["assigned_user_name"] == 'pknotek' || lead["assigned_user_name"] == 'kodonnell'
      forecast_report_table << [ lead["account_name"],
                                 lead["name"],
                                 lead["assigned_user_name"],
                                 lead["probability"] + "%",
                                 lead["next_step"],
                                 lead["date_closed"].to_s,
                                 Currency::Money(lead["amount"].to_s, :USD).format ]
      totals["grand"] += lead["amount"].to_i
      $config["probabilities"].each do | probability |
        if lead["probability"].to_i >= probability["start"].to_i && lead["probability"].to_i <= probability["end"].to_i
          totals["#{probability['name']}"] += lead["amount"].to_i
        end
      end
    end
  end
end

#Sort the forecast table
#forecast_report_table.sort_rows_by!(["Probability", "Amount", "ExpectedClose"], :order => :descending)
forecast_report_table.sort_rows_by!("Probability", :order => :descending)

#USD value
totals_report_table << [ "USD",
                         totals["closed"].format,
                         totals["commit"].format,
                         totals["upside"].format,
                         totals["swing"].format,
                         totals["undetermined"].format,
                         totals["grand"].format ]
#Euro value
totals_report_table << [ "EUR",
                         totals["closed"].convert(:EUR).format,
                         totals["commit"].convert(:EUR).format,
                         totals["upside"].convert(:EUR).format,
                         totals["swing"].convert(:EUR).format,
                         totals["undetermined"].convert(:EUR).format,
                         totals["grand"].convert(:EUR).format ]
                               
pdf_filename = "tmp/Forecast_" + process_time.to_s.gsub(" ", "-") + ".pdf"
TableRenderer.render_pdf( :file => pdf_filename,
                          :report_title => "NA & APAC Sales Forecast",
                          :sales_quarter => this_quarter + "FY" + process_time.year.to_s,
                          :salesperson => "USA Team",
                          :report_type => "sales_forecast",
                          :data => [ forecast_report_table, totals_report_table ] )
                          
#Generate CSV file
csv_filename = "tmp/Forecast_" + process_time.to_s.gsub(" ", "-") + ".csv"
File.open(csv_filename, "w") do |outfile|
  outfile.puts forecast_report_table.to_csv
end

#send_email pdf_filename, csv_filename
#File.delete(pdf_filename)
#File.delete(csv_filename)
           
puts "Completed"