module ExcelUtils
  module WorksheetIterators
    class ExcelBatchIterator < ExcelIterator

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