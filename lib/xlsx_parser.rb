require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'

Spreadsheet.client_encoding = 'UTF-8'
all_files =  Dir.glob('excel_files/**/*').select{ |e| File.file? e }
total_records = 0

csv_header = ["Supplier ID", "Supplier", "Panel/Side ID", "Environment", "Format Group", "Format Type", "Address", "Postcode", "City", "Longitude", "Latitude"]

data = []
data.push(csv_header)

all_files.each do |file|
  puts "Processing #{file}"
  puts File.extname(file)


  case File.extname(file)
  when '.xlsx'
    workbook = RubyXL::Parser.parse(file)
    workbook.worksheets.each do |ws|
      rows = ws.extract_data
      rows.each do |row|
        data.push(row)
      end
    end

  when '.xls'
    book = Spreadsheet.open(file)
    book.worksheets
    sheet1 = book.worksheet(0)
    sheet1.each 1 do |row|
      total_records += 1
      data.push(row)
    end
    
  else
    puts "Wrong extention"
  end
end

puts total_records
puts data.length


csv_file_path = "csv_files/final.csv"

CSV.open(csv_file_path, "wb") do |csv|
  data.each do |row|
    csv << row
  end
end
