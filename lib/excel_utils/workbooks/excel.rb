module ExcelUtils
  module Workbooks
    class Excel

      attr_reader :filename, :normalize_column_names

      def initialize(filename, normalize_column_names: false, extension: nil)
        @spreadsheet = Roo::Spreadsheet.open filename, extension: extension
        @normalize_column_names = normalize_column_names
      end

      def sheets
        @sheets ||= begin
          streaming = spreadsheet.respond_to? :each_row_streaming
          spreadsheet.sheets.map do |name|
            if streaming
              Sheets::ExcelStream.new name, @spreadsheet, normalize_column_names
            else
              Sheets::Excel.new name, @spreadsheet, normalize_column_names
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

    end

  end
end