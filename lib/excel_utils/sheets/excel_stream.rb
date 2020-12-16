module ExcelUtils
  module Sheets
    class ExcelStream < Base

      def initialize(spreadsheet:, **options)
        super(**options)
        @spreadsheet = spreadsheet
      end

      private

      attr_reader :spreadsheet

      def first_row
        row = sheet.each_row_streaming(pad_cells: true, max_rows: 0).first || []
        normalize_row row
      end

      def each_row
        sheet.each_row_streaming(pad_cells: true, offset: 1) do |row|
          yield normalize_row(row)
        end
      end

      def normalize_row(row)
        row.map { |cell| cell ? cell.value : cell }
      end

      def sheet
        spreadsheet.sheet name
      end

    end
  end
end