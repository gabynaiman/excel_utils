module ExcelUtils
  module WorksheetIterators
    class CsvIterator

      include Normalizer

      def initialize(name, data, normalize_column_names)
        @data = data
        @normalize_column_names = normalize_column_names
      end

      def each(&block)
        data.each(&block)
      end

      def column_names
        []
      end

    end
  end
end
