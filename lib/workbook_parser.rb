require 'rubygems'
require 'rubyXL'
require 'spreadsheet'
require 'csv'

class WorkbookParser
  attr_accessor :file_path, :all_rows, :row, :sheet_name
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
      @sheet_name = sheet.name
      sheet.each 1 do |row|
        add_row(row.to_a)
      end
    end
  end

  def parse_xlsx
    workbook = RubyXL::Parser.parse(@file_path)
    workbook.worksheets.each do |ws|
      @sheet_name = ws.sheet_name
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
      "SheetName" => sheet_name,
    }
    @all_rows.push(row_hash)
  end

  def supplier_id
    @row[0]
  end

  def supplier
    @row[1]
  end

  def panel_side_id
    panelsideid = @row[2]
    case fileName
    when 'JCDecaux_all.xlsx'
      panelsideid = "#{@row[2]} / #{@row[3]}"
      @row.delete_at(3)
    end
    panelsideid
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
    addr = @row[6]
    case fileName
    when 'CS Digital_all.xlsx'
      addr = @row[7]
    end
    addr
  end

  def postcode
    pc = @row[7]
    case fileName
    when 'CS Digital_all.xlsx'
      pc = @row[8]
    end
    pc
  end

  def city
    city = @row[8]
    case fileName
    when 'CS Digital_all.xlsx'
      city = @row[6]
    end
    city
  end

  def longitude
    lng = @row[9]
    case fileName
    when 'CS Digital_all.xlsx'
      lng = @row[11]
    end
    lng.to_s.split('.').join('.').gsub(' ','').to_f
  end

  def latitude
    lat = @row[10]
    case fileName
    when 'CS Digital_all.xlsx'
      lat = @row[10]
    end
    lat.to_s.split('.').join('.').gsub(' ','').to_f
  end

  def fileName
    File.basename(@file_path)
  end
end

# file_path = 'excel_files/Billboards_3bigcities_Clear Channel_135.xls'
# wp = WorkbookParser.new(file_path)
# puts wp.get_all_rows
