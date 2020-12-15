module ExcelUtils
  module Sheets
    class CSV < Base

      def initialize(filename, name, normalize_column_names)
        super name, normalize_column_names
        @filename = filename
      end

      private

      def each_row
        first = true
        NesquikCSV.foreach(@filename) do |row|
          yield row unless first
          first = false
        end
      end

      def get_column_names
        columns = []
        NesquikCSV.foreach(@filename) do |header|
          columns = header
          break
        end
        columns
      end
    end

  end
end
