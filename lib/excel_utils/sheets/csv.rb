module ExcelUtils
  module Sheets
    class CSV < Base

      def initialize(filename:, **options)
        super(**options)
        @filename = filename
      end

      private

      attr_reader :filename

      def first_row
        NesquikCSV.open(filename) { |csv| csv.readline } || []
      end

      def each_row
        first = true
        NesquikCSV.foreach(filename) do |row|
          yield row unless first
          first = false
        end
      end

    end
  end
end