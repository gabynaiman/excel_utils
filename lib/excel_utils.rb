require 'date'
require 'time'
require 'roo'
require 'roo-xls'
require 'write_xlsx'
require 'inflecto'

require_relative 'excel_utils/version'
require_relative 'excel_utils/workbook'
require_relative 'excel_utils/sheet'
require_relative 'excel_utils/writer'

module ExcelUtils
  
  def self.read(filename, **options)
    Workbook.new filename, **options
  end

  def self.write(filename, data)
    Writer.write filename, data
  end

end
