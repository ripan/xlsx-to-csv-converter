require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'

class WorkbookParser
  attr_accessor :file_path, :all_rows
  def initialize(file_path)
    @file_path = file_path
    @all_rows = []
  end

  def get_all_rows
    case File.extname(@file_path)
    when '.xlsx'
      parse_xlsx
    when '.xls'
      parse_xls
    else
      raise "Wrong extention"
    end
    return @all_rows
  end

  def parse_xls
    Spreadsheet.client_encoding = 'UTF-8'
    workbook = Spreadsheet.open(@file_path)
    workbook.worksheets.each_with_index do |row, index|
      sheet = workbook.worksheet(index)
      sheet.each 1 do |row|
        @all_rows.push(row)
      end
    end
  end

  def parse_xlsx
    workbook = RubyXL::Parser.parse(@file_path)
    workbook.worksheets.each do |ws|
      rows = ws.extract_data
      rows.drop(1).each do |row|
        @all_rows.push(row)
      end
    end
  end

end

# file_path = 'excel_files/Billboards_3bigcities_Clear Channel_135.xls'
# wp = WorkbookParser.new(file_path)
# puts wp.get_all_rows
