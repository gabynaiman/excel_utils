module ExcelUtils
  module Spreadsheet

    class ExcelSpreadsheet

      def initialize(filename, options)
        @reader = Roo::Spreadsheet.open filename, **options
        @options = options
      end

      def sheets
        @sheets ||= begin
          @reader.sheets.map do |name|
            iterator = WorksheetIterators.iterator_for name, @reader, @options.fetch(:extension),
                                                                      @options.fetch(:normalize_column_names)
            Sheet.new name, iterator
          end
        end
      end
    end

    class CsvSpreadsheet

      def initialize(filename, options)
        @reader = NesquikCSV.open filename
        @normalize_column_names = options.fetch :normalize_column_names
      end

      def sheets
        @sheets ||= begin
          iterator = WorksheetIterators::CsvIterator.new @reader, @normalize_column_names
          sheet = Sheet.new 'default', iterator
          [sheet]
        end
      end
    end

    EXTENSIONS = {
      'csv' => CsvSpreadsheet,
      'xls' => ExcelSpreadsheet,
      'xlsx' => ExcelSpreadsheet
    }

    def self.build(filename, options)
      EXTENSIONS[options.fetch(:extension)].new filename, options
    end

  end

end
