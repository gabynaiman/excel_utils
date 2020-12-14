module ExcelUtils
  module WorksheetIterators
    class CsvIterator

      include Normalizer

      attr_reader :column_names

      def initialize(reader, normalize_column_names)
        @normalize_column_names = normalize_column_names
        @column_names = get_columns reader.gets
        @rows = reader.read.map { |r| Hash[column_names.zip(r)] }
      end

      def each(&block)
        @rows.each(&block)
      end

      private

      def get_columns(columns)
        @normalize_column_names ? normalize_columns(columns) : columns
      end

    end
  end
end
