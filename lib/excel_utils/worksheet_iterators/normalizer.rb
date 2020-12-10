module WorksheetIterators
  module Normalizer

    def normalize_columns(names)
      names.map do |name|
        Inflecto.underscore(name.strip.gsub(' ', '_')).to_sym
      end
    end

  end
end
