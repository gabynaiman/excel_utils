module WorksheetIterators
  class Base

    def initialize(sheet, normalize_column_names)
      @sheet = sheet
      @normalize_column_names = normalize_column_names
    end

    def column_names
      @column_names ||= begin
        if sheet.first_row
          first_row = sheet.row sheet.first_row
          normalize_column_names ? first_row.map { |name| normalize_column name } : first_row
        else
          []
        end
      end
    end

    def normalize_column(name)
      Inflecto.underscore(name.strip.gsub(' ', '_')).to_sym
    end

    private

    attr_reader :sheet, :normalize_column_names

  end
end
