module ExcelUtils
  module Workbooks
    class Excel

      attr_reader :filename, :normalize_column_names

      def initialize(filename, normalize_column_names: false, extension: nil)
        @filename = filename
        @normalize_column_names = normalize_column_names
        @spreadsheet = Roo::Spreadsheet.open filename, extension: extension
      end

      def sheets
        @sheets ||= spreadsheet.sheets.map do |name|
          sheet_class.new name: name,
                          normalize_column_names: normalize_column_names,
                          spreadsheet: spreadsheet
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

      def sheet_class
        @sheet_class ||= spreadsheet.respond_to?(:each_row_streaming) ? Sheets::ExcelStream : Sheets::Excel
      end

    end
  end
end