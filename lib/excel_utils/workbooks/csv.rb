module ExcelUtils
  module Workbooks
    class CSV

      SHEET_NAME = 'default'.freeze

      attr_reader :filename, :normalize_column_names

      def initialize(filename, normalize_column_names: false)
        @filename = filename
        @normalize_column_names = normalize_column_names

        @sheet = Sheets::CSV.new name: SHEET_NAME,
                                 normalize_column_names: normalize_column_names,
                                 filename: filename
      end

      def sheets
        [sheet]
      end

      def [](sheet_name)
        sheet_name == SHEET_NAME ? sheet : nil
      end

      def to_h
        {SHEET_NAME => sheet.to_a}
      end

      private

      attr_reader :sheet

    end
  end
end