module ExcelUtils
  class Workbook

    attr_reader :filename, :normalize_column_names, :file_ext

    def initialize(filename, normalize_column_names: false, extension: nil)
      @filename = filename
      @normalize_column_names = normalize_column_names
      @file_ext = file_extension filename, extension

      if file_ext == 'csv'
        @spreadsheet = NesquikCSV.open filename
      else
        @spreadsheet = Roo::Spreadsheet.open filename, extension: file_ext
      end
    end

    def sheets
      @sheets ||= begin
        if file_ext == 'csv'
          iterator = WorksheetIterators.iterator_for 'default', spreadsheet, file_ext, normalize_column_names

          unique_sheet = Sheet.new 'default', iterator
          [unique_sheet]
        else
          spreadsheet.sheets.map do |name|
            sheet = spreadsheet.sheet name
            iterator = WorksheetIterators.iterator_for name, sheet, file_ext, normalize_column_names

            Sheet.new name, iterator
          end
        end
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

    def file_extension(filename, extension)
      ext = extension.nil? ? File.extname(filename[1..-1]) : extension
      ext.tr '.', ''
    end

  end
end