require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'
require_relative '../lib/workbook_parser'

all_files =  Dir.glob('excel_files/**/*').select{ |e| File.file? e }

csv_header = ["Supplier ID", "Supplier", "Panel/Side ID", "Environment", "Format Group", "Format Type", "Address", "Postcode", "City", "Longitude", "Latitude"]

data = []
data.push(csv_header)

all_files.each do |file_path|
  puts "Processing #{file_path}"
  wp = WorkbookParser.new(file_path)
  all_workbook_rows = wp.get_all_rows
  puts "#{all_workbook_rows.length} records found"
  data.concat(all_workbook_rows)
end

puts "DATA: #{data.length} "

csv_file_path = "csv_files/final.csv"

CSV.open(csv_file_path, "wb") do |csv|
  data.each do |row|
    csv << row
  end
end
