module ExcelUtils
  module Workbooks
    class CSV

      DEFAULT_SHEET_NAME = 'default'.freeze

      def initialize(filename, normalize_column_names: false)
        @iterator = Sheets::CSV.new filename, DEFAULT_SHEET_NAME, normalize_column_names
      end

      def sheets
        [@iterator]
      end

      def [](sheet_name)
        sheet_name == DEFAULT_SHEET_NAME ? @iterator : nil
      end

      def to_h
        {DEFAULT_SHEET_NAME => @iterator.to_a}
      end

    end

  end
end