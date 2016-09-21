require 'mysql2'
require 'csv'
require 'mail'
require 'spreadsheet'

# Variables
@start_datetime = (Date.today - 1).to_s + " 22:00:00"
@end_datetime = Time.new.utc.strftime("%Y-%m-%d %H:%M:%S")
@end_phdatetime = Time.new.strftime("%Y%b%d_%H%M").to_s

# Load configuration
CONFIG = YAML.load_file(File.join(__dir__, 'config.yml'))

# MySQL connection
def connect_to_db
  db_server = CONFIG["mysql"]["host"]
  puts "Connecting to server #{db_server}..."
  @con = Mysql2::Client.new(
    :host => CONFIG["mysql"]["host"],
    :database => CONFIG["mysql"]["database"],
    :username => CONFIG["mysql"]["username"],
    :password => CONFIG["mysql"]["password"])
  puts 'Connection successful!'
end

# MySQL query
def db_query
  puts 'Executing query...'
  @results = @con.query("SELECT
    'Agent Name',
    'Qualified / Appointments',
    'AM / VM /Fax',
    'Not Interested',
    'Disconnected / Wrong Number / Duplicate',
    'No Answer',
    'Call Back',
    'Hang Up',
    'Too Far',
    'Disqualified'
  UNION SELECT CONCAT(user.first_name,' ',user.last_name) as 'Agent Name' ,
    SUM(lead.call_dispo='Qualified / Appointments') as 'Qualified / Appointments',
    SUM(lead.call_dispo='AM / VM /Fax') as 'AM / VM / Fax',
    SUM(lead.call_dispo='Not Interested') as 'Not Interested',
    SUM(lead.call_dispo='Disconnected / Wrong Number / Duplicate') as 'Disconnected / Wrong Number / Duplicate',
    SUM(lead.call_dispo='No Answer') as 'No Answer',
    SUM(lead.call_dispo='Call Back') as 'Call Back',
    SUM(lead.call_dispo='Hang Up') as 'Hang Up',
    SUM(lead.call_dispo='Too Far') as 'Too Far',
    SUM(lead.call_dispo='Disqualified') as 'Disqualified'
  FROM note LEFT JOIN lead ON note.parent_id = lead.id
  LEFT JOIN user ON user.id = lead.assigned_user_id
  WHERE note.modified_at BETWEEN CAST('#{@start_datetime}' AS DATETIME) AND CAST('#{@end_datetime}' AS DATETIME)
  AND lead.call_dispo<>''
  AND note.data NOT LIKE '%\"became\":{\"callDispo\":\"None\"}%'
  AND note.data NOT LIKE '%\"assignedUserName\":%'
  AND note.data NOT LIKE '{}'
  GROUP BY user.id")
  puts 'Query completed.'
end

# Save all rows into a CSV file
def parse_to_csv
  puts 'Parsing results to CSV...'
  CSV.open("#{@end_phdatetime}_crm_lead_dispo_stat.csv", "wb", converters: :numeric) do |csv|
    @results.each(:as => :array) do |row|
      csv << row
    end
  end

  # Get total count leads modified
  # total_leads = Array.new
  # CSV.foreach("#{@end_phdatetime}_crm_lead_dispo_stat.csv", "wb", converters: :numeric)

  puts "Parse completed. File #{@end_phdatetime}_crm_lead_dispo_stat.csv is now available."
end


# Convert to XLS
def convert_to_xls
  puts "Converting CSV to XLS..."
  book = Spreadsheet::Workbook.new
  sheet1 = book.create_worksheet

  header_format = Spreadsheet::Format.new(
    :color => :blue,
    :weight => :bold,
    :horizontal_align => :center,
    :bottom => :double
  )

  sheet1.row(0).default_format = header_format

  CSV.open("#{@end_phdatetime}_crm_lead_dispo_stat.csv", 'r') do |csv|
    csv.each_with_index do |row, i|
      sheet1.row(i).replace(row)
    end
  end



  book.write("#{@end_phdatetime}_crm_lead_dispo_stat.xls")
  puts "Done with the conversion."
end

# Prepare to attach CSV file and send to email
def send_email_with_attachment
  puts 'Preparing email...'
  end_phdatetime = @end_phdatetime

  options = { :address              => CONFIG["o365_smtp"]["address"],
              :port                 => CONFIG["o365_smtp"]["port"],
              :domain               => CONFIG["o365_smtp"]["domain"],
              :user_name            => CONFIG["o365_smtp"]["username"],
              :password             => CONFIG["o365_smtp"]["password"],
              :authentication       => :login,
              :enable_starttls_auto => true  }

  mail = Mail.new do
    delivery_method :smtp, options
    from     CONFIG["o365_smtp"]["from"]
    to       CONFIG["o365_smtp"]["to"]
    subject  "#{end_phdatetime} CRM Lead Dispo Stats"
    body     "Attached is the CRM lead disposition status as of #{end_phdatetime}."
    add_file :filename => "#{end_phdatetime}_crm_lead_dispo_stat.xls", :content => File.read("#{end_phdatetime}_crm_lead_dispo_stat.xls")
  end

  mail.deliver
  puts 'You got mail.'
end

# Cleaning up
def cleanup
  puts 'Cleaning up...'
  File.delete("#{@end_phdatetime}_crm_lead_dispo_stat.csv")
  puts "#{@end_phdatetime}_crm_lead_dispo_stat.csv has been deleted."
end

connect_to_db
db_query
parse_to_csv
convert_to_xls
send_email_with_attachment
cleanup
