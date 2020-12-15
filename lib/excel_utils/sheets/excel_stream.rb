module ExcelUtils
  module Sheets
    class ExcelStream < Excel

      def each_row
        sheet.each_row_streaming(pad_cells: true, offset: 1) do |row|
          cells = row.map { |cell| cell ? cell.value : cell }
          yield cells
        end
      end

      def get_column_names
        columns = []
        sheet.each_row_streaming(pad_cells: true, max_rows: 0) do |row|
          cells = row.map { |cell| cell ? cell.value : cell }
          columns = cells
        end
        columns
      end
    end

  end
end
