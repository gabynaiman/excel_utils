module ExcelUtils
  module Sheets

    class Base

      include Enumerable

      attr_reader :name

      def each(&block)
        raise 'Not implemented'
      end

      def normalize_columns(names)
        names.map do |name|
          Inflecto.underscore(name.strip.gsub(' ', '_')).to_sym
        end
      end

    end

    class Excel < Base

      def initialize(name, spreadsheet, normalize_column_names)
        @name = name
        @spreadsheet = spreadsheet
        @normalize_column_names = normalize_column_names
      end

      def each(&block)
        rows.each(&block)
      end

      def column_names
        @column_names ||= begin
          if sheet.first_row
            first_row = sheet.row sheet.first_row
            normalize_column_names ? normalize_columns(first_row) : first_row
          else
            []
          end
        end
      end

      private

      attr_reader :spreadsheet, :normalize_column_names

      def sheet
        spreadsheet.sheet name
      end

      def rows
        @rows ||= begin
          if sheet.first_row
            sheet.to_a[1..-1].map do |row|
              Hash[column_names.zip(row)]
            end
          else
            []
          end
        end
      end

    end

    class ExcelStream < Excel

      def each(&block)
        sheet.each_row_streaming(pad_cells: true, offset: 1) do |row|
          cells = row.map { |cell| cell ? cell.value : cell }
          block.call column_names.zip(cells).to_h
        end
      end

      def column_names
        @column_names ||= begin
          columns = []
          sheet.each_row_streaming(pad_cells: true, max_rows: 0) do |row|
            cells = row.map { |cell| cell ? cell.value : cell }
            columns = normalize_column_names ? normalize_columns(cells) : cells
          end
          columns
        end
      end
    end

    class CSV < Base

      attr_reader :column_names

      def initialize(filename, normalize_column_names)
        reader = NesquikCSV.open filename
        @normalize_column_names = normalize_column_names
        @column_names = get_columns reader.gets
        @rows = reader.read.map { |r| Hash[column_names.zip(r)] }
        reader.close
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
