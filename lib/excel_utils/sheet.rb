module ExcelUtils
  class Sheet

    include Enumerable

    attr_reader :name, :normalize_column_names

    def initialize(name, spreadsheet, normalize_column_names: false, iterator_strategy: 'batch')
      @name = name
      @spreadsheet = spreadsheet
      @normalize_column_names = normalize_column_names
      @iterator_strategy = iterator_strategy
    end

    def column_names
      @column_names ||= iterator.column_names
    end

    def each(&block)
      iterator.each(&block)
    end

    private

    attr_reader :spreadsheet, :iterator_strategy

    def sheet
      spreadsheet.sheet name
    end

    def iterator
      @iterator ||= begin
        if iterator_strategy == 'stream'
          raise 'Invalid extension. Cannot use stream strategy with this file' unless sheet.respond_to? :each_row_streaming
          WorksheetIterators::StreamIterator.new sheet, @normalize_column_names
        else
          WorksheetIterators::BatchIterator.new sheet, @normalize_column_names
        end
      end
    end

  end
end