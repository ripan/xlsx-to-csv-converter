require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'
require 'colorize'
require_relative '../lib/workbook_parser'

all_files =  Dir.glob('excel_files/**/*').select{ |e| File.file? e }

#csv_header = ["Supplier ID", "Supplier", "Panel/Side ID", "Environment", "Format Group", "Format Type", "Address", "Postcode", "City", "Longitude", "Latitude", "FileName", "SheetName"]

data = []

all_files.each do |file_path|
  puts "\nProcessing #{file_path}"
  begin
    wp = WorkbookParser.new(file_path)
    all_workbook_rows = wp.get_all_rows
    puts "#{all_workbook_rows.length} records found"
    data.concat(all_workbook_rows)
  rescue Exception => e
  	puts "\nERROR: #{e}".red
    FileUtils.mv(file_path, 'error_files')
  end
end

puts "\nDATA: #{data.length} "

csv_file_path = "csv_files/final.csv"

CSV.open(csv_file_path, "wb") do |csv|
  csv << data[0].keys # to get CSV header columns
  data.each do |row|
    csv << row.values
  end
end
