module ExcelUtils
  class Writer

    TIME_FORMAT = '%Y-%m-%dT%H:%M:%S'

    EXCEL_FORMATS = {
      date: 'yyyy-mm-dd',
      date_time: 'yyyy-mm-dd hh:mm:ss'
    }

    class << self

      def write(filename, data)
        workbook = WriteXLSX.new filename, strings_to_urls: false

        formats = add_formats workbook

        data.each do |sheet_name, sheet_data|
          add_sheet workbook, sheet_name, sheet_data, formats
        end

        workbook.close
      end

      private

      def add_formats(workbook)
        EXCEL_FORMATS.each_with_object({}) do |(type, value), formats|
          formats[type] = workbook.add_format num_format: value
        end
      end

      def add_sheet(workbook, sheet_name, sheet_data, formats)
        sheet = workbook.add_worksheet sheet_name

        if sheet_data.any?
          header = sheet_data.flat_map(&:keys).uniq
          sheet.write_row 0, 0, header.map(&:to_s)

          sheet_data.each_with_index do |row, r|
            row_index = r + 1
            header.each_with_index do |column, col_index|
              if row[column]
                if row[column].is_a? String
                  sheet.write_string row_index, col_index, row[column]

                elsif row[column].respond_to? :to_time
                  time = row[column].to_time
                  type = date?(time) ? :date : :date_time
                  sheet.write_date_time row_index, col_index, time.to_time.strftime(TIME_FORMAT), formats[type]

                else
                  sheet.write row_index, col_index, row[column]
                end
              end
            end
          end
        end
      end

      def date?(time)
        time.hour + time.min + time.sec == 0
      end

    end

  end
end