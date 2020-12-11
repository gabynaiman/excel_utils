module ExcelUtils
  module WorksheetIterators

    EXTENSION_MAPPINGS = {
      'xls' => ExcelBatchIterator,
      'xlsx' => ExcelStreamIterator,
      'csv' => CsvIterator
    }

  	def self.iterator_for(name, sheet, extension, normalize_column_names)
      raise "Invalid extension #{extension}" unless EXTENSION_MAPPINGS.key? extension
      EXTENSION_MAPPINGS[extension].new name, sheet, normalize_column_names
    end

  end
end
