module ExcelUtils
  module Sheets
    class Base

      include Enumerable

      attr_reader :name

      def initialize(name, normalize_column_names)
        @name = name
        @normalize_column_names = normalize_column_names
      end

      def each(&block)
        each_row do |row|
          block.call column_names.zip(row).to_h
        end
      end

      def column_names
        @column_names ||= begin
          columns = get_column_names
          @normalize_column_names ? normalize_columns(columns) : columns
        end
      end

      private

      def normalize_columns(names)
        names.map do |name|
          Inflecto.underscore(name.strip.gsub(' ', '_')).to_sym
        end
      end

    end

  end
end
