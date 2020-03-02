module ExcelUtils
  class Sheet

    include Enumerable
    
    attr_reader :name, :normalize_column_names
    
    def initialize(name, spreadsheet, normalize_column_names: false)
      @name = name
      @spreadsheet = spreadsheet
      @normalize_column_names = normalize_column_names
    end

    def column_names
      @column_names ||= begin
        if sheet.first_row
          first_row = sheet.row sheet.first_row
          normalize_column_names ? first_row.map { |n| Inflecto.underscore(n.strip.gsub(' ', '_')).to_sym } : first_row
        else
          []
        end
      end
    end

    def each(&block)
      rows.each(&block)
    end

    private

    attr_reader :spreadsheet

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
end