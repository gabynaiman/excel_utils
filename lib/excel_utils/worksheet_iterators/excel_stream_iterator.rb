module ExcelUtils
  module WorksheetIterators
    class ExcelStreamIterator < ExcelIterator

      def each(&block)
        sheet.each_row_streaming(pad_cells: true, offset: 1) do |row|
          cells = row.map { |cell| cell ? cell.value : cell }
          block.call column_names.zip(cells).to_h
        end
      end

      def column_names
        @column_names ||= begin
          columns = []
          sheet.each_row_streaming(pad_cells: true, max_rows: 0) do |row|
            cells = row.map { |cell| cell ? cell.value : cell }
            columns = normalize_column_names ? normalize_columns(cells) : cells
          end
          columns
        end
      end

    end
  end
end