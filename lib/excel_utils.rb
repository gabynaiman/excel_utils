require 'date'
require 'time'
require 'roo'
require 'roo-xls'
require 'write_xlsx'
require 'inflecto'
require 'nesquikcsv'

require_relative 'excel_utils/version'
require_relative 'excel_utils/workbooks/excel'
require_relative 'excel_utils/workbooks/csv'
require_relative 'excel_utils/sheets/base'
require_relative 'excel_utils/sheets/excel'
require_relative 'excel_utils/sheets/excel_stream'
require_relative 'excel_utils/sheets/csv'
require_relative 'excel_utils/writer'

module ExcelUtils

  def self.read(filename, **options)
    extension = options.fetch(:extension, File.extname(filename)[1..-1])
    if extension == 'csv'
      Workbooks::CSV.new filename, **options
    else
      Workbooks::Excel.new filename, **options
    end
  end

  def self.write(filename, data)
    Writer.write filename, data
  end

end
