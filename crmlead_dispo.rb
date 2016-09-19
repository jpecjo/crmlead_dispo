require 'mysql2'
require 'csv'
require 'mail'

# Variables
@start_datetime = (Date.today - 1).to_s + " 22:00:00"
@end_datetime = Time.new.utc.strftime("%Y-%m-%d %H:%M:%S")
@end_phdatetime = Time.new.strftime("%Y-%m-%d_%H%M")

# Load configuration
CONFIG = YAML.load_file("./config.yml")

# MySQL connection
def connect_to_db
  puts 'Connecting to server CONFIG["mysql"]["host"]...'
  @con = Mysql2::Client.new(
    :host => CONFIG["mysql"]["host"],
    :database => CONFIG["mysql"]["database"],
    :username => CONFIG["mysql"]["username"],
    :password => CONFIG["mysql"]["password"])
  puts 'Connectiong successful!'
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
  FROM lead LEFT JOIN user ON lead.assigned_user_id=user.id
  WHERE lead.modified_at BETWEEN CAST('#{@start_datetime}' AS DATETIME) AND CAST('#{@end_datetime}' AS DATETIME)
  GROUP BY user.id")
  puts 'Query completed.'
end

# Save all rows into a CSV file
def parse_to_csv
  puts 'Parsing results to CSV...'
  CSV.open("./#{@end_phdatetime}_crm_lead_dispo_stat.csv", "wb") do |csv|
    @results.each(:as => :array) do |row|
      csv << row
    end
  end
  puts "Parse completed. File #{@end_phdatetime}_crm_lead_dispo_stat.csv is now available."
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
    add_file :filename => "./#{end_phdatetime}_crm_lead_dispo_stat.csv", :content => File.read("./#{end_phdatetime}_crm_lead_dispo_stat.csv")
  end

  mail.deliver
  puts 'You got mail.'
end

# Cleaning up
def cleanup
  puts 'Cleaning up...'
  File.delete("./#{@end_phdatetime}_crm_lead_dispo_stat.csv")
  puts "#{@end_phdatetime}_crm_lead_dispo_stat.csv has been deleted."
end

connect_to_db
db_query
parse_to_csv
send_email_with_attachment
cleanup
