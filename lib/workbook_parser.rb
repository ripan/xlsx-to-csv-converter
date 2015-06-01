require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'

class WorkbookParser
  attr_accessor :file_path, :all_rows, :row
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
        add_row(row.to_a)
      end
    end
  end

  def parse_xlsx
    workbook = RubyXL::Parser.parse(@file_path)
    workbook.worksheets.each do |ws|
      rows = ws.extract_data
      rows.drop(1).each do |row|
        add_row(row)
      end
    end
  end

  def add_row(row)
    @row = row
    row_hash = {
      "Supplier ID" => supplier_id,
      "Supplier" => supplier,
      "Panel/Side ID" => panel_side_id,
      "Environment" => environment,
      "Format Group" => format_group,
      "Format Type" => format_type,
      "Address" => address,
      "Postcode" => postcode,
      "City" => city,
      "Longitude" => longitude,
      "Latitude" => latitude,
      "FileName" => fileName,
    }
    puts @row.inspect
    puts row_hash.inspect
    asdas
    #@all_rows.push(row)
  end

  def supplier_id
    @row[0]
  end

  def supplier
    @row[1]
  end

  def panel_side_id
    case fileName
    when 'JCDecaux_all.xlsx'
      @row[2] = "#{@row[2]} / #{@row[3]}"
      @row.delete_at(3)
    when 6
      puts "It's 6"
    else
    end

    @row[2]
  end

  def environment
    @row[3]
  end

  def format_group
    @row[4]
  end

  def format_type
    @row[5]
  end

  def address
    @row[6]
  end

  def postcode
    @row[7]
  end

  def city
    @row[8]
  end

  def longitude
    @row[9]
  end

  def latitude
    @row[10]
  end

  def fileName
    File.basename(@file_path)
  end
end

# file_path = 'excel_files/Billboards_3bigcities_Clear Channel_135.xls'
# wp = WorkbookParser.new(file_path)
# puts wp.get_all_rows
