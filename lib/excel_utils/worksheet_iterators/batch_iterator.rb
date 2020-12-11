module ExcelUtils
  module WorksheetIterators
    class BatchIterator

      include Normalizer

      def initialize(name, spreadsheet, normalize_column_names)
        @name = name
        @spreadsheet = spreadsheet
        @normalize_column_names = normalize_column_names
      end

      def each(&block)
        rows.each(&block)
      end

      def column_names
        @column_names ||= begin
          if sheet.first_row
            first_row = sheet.row sheet.first_row
            normalize_column_names ? normalize_columns(first_row) : first_row
          else
            []
          end
        end
      end

      private

      attr_reader :spreadsheet, :name, :normalize_column_names

      def sheet
        spreadsheet.sheet name
      end

      def rows
        @rows ||= begin
          if sheet.first_row
            sheet.to_a[1..-1].map do |row|
              Hash[column_names.zip(row)]
            end
          else
            []
          end
        end
      end

    end
  end
end
