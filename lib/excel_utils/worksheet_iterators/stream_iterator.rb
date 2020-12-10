module WorksheetIterators
  class StreamIterator < Base

    def initialize(sheet, normalize_column_names, offset:1)
      @sheet = sheet
      @normalize_column_names = normalize_column_names
      @offset = offset
    end

    def each(&block)
      sheet.each_row_streaming(pad_cells: true, offset: offset) do |row|
        cells = row.map { |cell| cell ? cell.value : cell }
        break if cells.all? { |cell| cell.to_s.strip.empty? }
        block.call column_names.zip(cells).to_h
      end
    end

    private

    attr_reader :offset

  end
end
