module ExcelUtils
  module Sheets
    class Excel < Base

      def initialize(spreadsheet:, **options)
        super(**options)
        @spreadsheet = spreadsheet
      end

      private

      attr_reader :spreadsheet

      def first_row
        with_sheet do |sheet|
          sheet.first_row ? sheet.row(sheet.first_row) : []
        end
      end

      def each_row
        with_sheet do |sheet|
          (sheet.first_row + 1).upto(sheet.last_row) do |i|
            yield sheet.row(i)
          end
        end
      end

      def with_sheet
        yield spreadsheet.sheet(name)
      end

    end

  end
end