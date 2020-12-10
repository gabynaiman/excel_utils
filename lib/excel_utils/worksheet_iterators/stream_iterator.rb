module WorksheetIterators
  class StreamIterator

    include Normalizer

    attr_reader :column_names

    def initialize(sheet, normalize_column_names)
      @sheet = sheet
      @normalize_column_names = normalize_column_names
      @column_names = nil
    end

    def each(&block)
      column_names = nil
      sheet.each_row_streaming(pad_cells: true) do |row|
        cells = row.map { |cell| cell ? cell.value : cell }
        break if cells.all? { |cell| cell.to_s.strip.empty? }

        if column_names.nil?
          column_names = normalize_column_names ? normalize_columns(cells) : cells
        else
          block.call column_names.zip(cells).to_h
        end
      end
    end

    private

    attr_reader :sheet, :normalize_column_names

  end
end
