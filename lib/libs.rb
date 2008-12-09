require 'rubygems'
require 'yaml'
require 'timeout'
require 'active_record'
require 'activesupport'
require 'ruport'
require 'currency'
require 'currency/exchange/rate/deriver'
require 'currency/exchange/rate/source/xe'
require 'currency/exchange/rate/source/timed_cache'
require 'net/smtp'
require 'mailfactory'
require 'ruby-units'
require 'soap/wsdlDriver'
require 'digest/md5'

module StandardPDFReport
 
  def build_standard_report_header
    pdf_writer.select_font("Times-Roman")
 
    options.text_format = { :font_size => 14, :justification => :right }
 
    add_text "<b>#{options.report_title}</b>"
    if options.sales_quarter != nil
      add_text "<b>#{options.sales_quarter}</b>"
    end
    add_text "Generated @ #{Time.now.strftime('%H:%M on %Y-%m-%d')}"
    add_text "<i>Company Confidential</i>"
    add_text "Forecast Definitions: http://wiki.presenceco.com/index.php/Salesops"
 
    center_image_in_box "img/logo.jpg", 
                        :x => left_boundary,
                        :y => top_boundary - 75,
                        :width => 100,
                        :height => 100
 
    move_cursor_to top_boundary - 80
 
    pad_bottom(20) { hr }
 
    options.text_format[:justification] = :left
    options.text_format[:font_size] = 12
  end
 
  def finalize_standard_report
    render_pdf
    pdf_writer.save_as(options.file)
  end
 
end
 
Ruport::Formatter::Template.create(:default) do |format|
 
    format.page = {
      :layout => :landscape
    }
 
    format.grouping = {
      :style => :separated
    }
 
    format.text = {
      :font_size => 12,
      :justification => :left
    }
 
    format.table = {
      :font_size => 10,
      :heading_font_size => 10,
      :maximum_width => 720,
      :width => 720
    }
 
    format.column = {
      :alignment => :left
    }
 
    format.heading = {
      :alignment => :left,
      :bold => false
    }
 
end
  
class DocumentRenderer < Ruport::Controller
 
  stage :standard_report_header
  finalize :standard_report
 
end
 
class TableRenderer < Ruport::Controller
 
  stage :standard_report_header, :table_body
  finalize :standard_report
 
end
 
class FormatterForPDF < Ruport::Formatter::PDF
  #page_layout :landscape
 
  renders :pdf, :for => [DocumentRenderer, TableRenderer]
  
  include StandardPDFReport
  
  def build_table_body
    if options.report_type == "sales_forecast"
      
      #Build sales forecast report
      pad_bottom(5) do
        add_text "<b>Forecast for " + options.salesperson + "</b>", :font_size => 14
      end
      draw_table(data[0])
      add_text " "
      pad_bottom(5) do
        add_text "<b>Totals</b>", :font_size => 14
      end
      
      draw_table(data[1])
      
    end
  end
end

def send_email pdf_file, csv_file
  #Create a new email object and build the headers
  emailMessage = MailFactory.new
  emailMessage.replyto=($config["email_from_address"])
  emailMessage.from=($config["email_from_address"])
  emailMessage.subject=("AUTO-GENERATED - Presence Sales Forecast Report")
  emailMessage.text=("Attached files for your perusal. Please remember the 'csv' file may be opened with Excel.")
  emailMessage.attach(pdf_file)
  
  if csv_file != nil
    emailMessage.attach(csv_file)
  end

  #Send the email via the SMTP server
  if $config["email_smtp_username"] == nil
    begin
      smtp = Net::SMTP.start($config["email_smtp_gateway"], $config["email_smtp_port"], $config["email_smtp_from_domain"])
    rescue => err
      puts "Error opening SMTP connection - " + err
    end
  else
    begin
      smtp = Net::SMTP.start($config["email_smtp_gateway"], $config["email_smtp_port"], $config["email_smtp_from_domain"],
                             $config["email_smtp_username"], $config["email_smtp_password"], $config["email_smtp_type"])
    rescue => err
      puts "Error opening SMTP connection - " + err
      exit
    end
  end
  begin
    smtp.send_message(emailMessage.to_s, $config["email_from_address"], $config["email_to_address"])
  rescue => err
    puts "Error sending email message - " + err
    exit
  end
  smtp.finish
  puts "Emailed report to: " + $config["email_to_address"]
end