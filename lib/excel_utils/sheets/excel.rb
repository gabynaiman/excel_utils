module ExcelUtils
  module Sheets
    class Excel < Base

      def initialize(name, spreadsheet, normalize_column_names)
        super name, normalize_column_names
        @spreadsheet = spreadsheet
      end

      private

      attr_reader :spreadsheet, :normalize_column_names

      def each_row
        if sheet.first_row
          first = true
          sheet.each do |row|
            yield row unless first
            first = false
          end
        else
          []
        end
      end

      def get_column_names
        if sheet.first_row
          sheet.row sheet.first_row
        else
          []
        end
      end

      def sheet
        spreadsheet.sheet name
      end

    end

  end
end
