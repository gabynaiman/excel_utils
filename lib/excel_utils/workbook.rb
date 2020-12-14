module ExcelUtils
  class Workbook

    attr_reader :filename, :normalize_column_names, :file_extension

    def initialize(filename, normalize_column_names: false, extension: nil)
      @filename = filename
      @normalize_column_names = normalize_column_names
      @file_extension = get_file_extension filename, extension
      @spreadsheet = Spreadsheet.build filename, extension: file_extension,
                                                 normalize_column_names: normalize_column_names
    end

    def sheets
      spreadsheet.sheets
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

    def get_file_extension(filename, extension)
      ext = extension.nil? ? File.extname(filename[1..-1]) : extension
      ext.tr '.', ''
    end

  end
end