module ExcelUtils
  module WorksheetIterators
    class ExcelIterator

      include Normalizer

      def initialize(name, spreadsheet, normalize_column_names)
        @name = name
        @spreadsheet = spreadsheet
        @normalize_column_names = normalize_column_names
      end

      private

      attr_reader :name, :spreadsheet, :normalize_column_names

      def sheet
        spreadsheet.sheet name
      end

    end
  end
end
