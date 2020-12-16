module ExcelUtils
  module Sheets
    class Base

      include Enumerable

      attr_reader :name, :normalize_column_names

      def initialize(name:, normalize_column_names: false)
        @name = name
        @normalize_column_names = normalize_column_names
      end

      def column_names
        @column_names ||= normalize_column_names ? normalize_columns(first_row) : first_row
      end

      def each
        if column_names.any?
          each_row do |row|
            break if empty_row? row
            yield Hash[column_names.zip(row)]
          end
        end
      end

      private

      def normalize_columns(names)
        names.map do |name|
          Inflecto.underscore(name.strip.gsub(' ', '_')).to_sym
        end
      end

      def empty_row?(row)
        row.all? { |cell| cell.to_s.strip.empty? }
      end

    end
  end
end