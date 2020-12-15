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

    class Excel < Base

      def initialize(name, spreadsheet, normalize_column_names)
        super name, normalize_column_names
        @spreadsheet = spreadsheet
      end

      private

      attr_reader :spreadsheet, :normalize_column_names

      def each_row
        first = true
        spreadsheet.each do |row|
          yield row unless first
          first = false
        end
      end

      def get_column_names
        if sheet.first_row
          sheet.row sheet.first_row
        else
          []
        end
      end

      def sheet
        spreadsheet.sheet name
      end

    end

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
