module WorksheetIterators
  class BatchIterator < Base

    def initialize(sheet, normalize_column_names)
      @sheet = sheet
      @normalize_column_names = normalize_column_names
    end

    def each(&block)
      rows.each(&block)
    end

    private

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
end
