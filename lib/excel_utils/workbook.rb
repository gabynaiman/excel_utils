module ExcelUtils
  class Workbook

    attr_reader :filename, :normalize_column_names
    
    def initialize(filename, normalize_column_names: false, extension: nil)
      @filename = filename
      @normalize_column_names = normalize_column_names
      @spreadsheet = Roo::Spreadsheet.open filename, extension: extension
    end

    def sheets
      @sheets ||= spreadsheet.sheets.map do |name| 
        Sheet.new name, spreadsheet, normalize_column_names: normalize_column_names
      end
    end

    def [](sheet_name)
      sheets.detect { |sheet| sheet.name == sheet_name }
    end

    def to_h
      sheets.each_with_object({}) do |sheet, hash|
        hash[sheet.name] = sheet.to_a
      end
    end

    private

    attr_reader :spreadsheet

  end
end