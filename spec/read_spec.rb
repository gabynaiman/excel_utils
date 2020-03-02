require 'minitest_helper'

describe ExcelUtils, 'Read' do

  def expected_rows(sheet)
    rows_by_sheet[sheet.name].map { |r| columns_by_sheet[sheet.name].zip(r).to_h }
  end

  let :rows_by_sheet do
    {
      'Sheet1' => [
        [1.0, 'some text'],
        [2.0, 1.35],
        [3.0, Date.parse('2019-08-17')],
        [4.0, nil]
      ],
      'Sheet2' => [
        [123.0, 'Text 1'],
        [456.0, 'Text 2']
      ],
      'Sheet3' => []
    }
  end

  ['xls', 'xlsx'].each do |extension|

    describe extension do
  
      let(:filename) { File.expand_path "../sample.#{extension}", __FILE__ }

      describe 'Original column names' do

        let(:workbook) { ExcelUtils.read filename }
        
        let :columns_by_sheet do
          {
            'Sheet1' => ['Column A', 'Column B'],
            'Sheet2' => ['ID', 'Value'],
            'Sheet3' => []
          }
        end

        it 'to_h' do
          workbook.to_h.must_equal 'Sheet1' => expected_rows(workbook['Sheet1']),
                                   'Sheet2' => expected_rows(workbook['Sheet2']),
                                   'Sheet3' => []
        end

        ['Sheet1', 'Sheet2', 'Sheet3'].each do |sheet_name|

          describe sheet_name do

            let(:sheet) { workbook[sheet_name] }

            it 'Column names' do
              sheet.column_names.must_equal columns_by_sheet[sheet_name]
            end

            it 'Rows' do
              sheet.to_a.must_equal expected_rows(sheet)
            end

          end

        end

      end

      describe 'Normalized column names' do

        let(:workbook) { ExcelUtils.read filename, normalize_column_names: true }
        
        let :columns_by_sheet do
          {
            'Sheet1' => [:column_a, :column_b],
            'Sheet2' => [:id, :value],
            'Sheet3' => []
          }
        end

        ['Sheet1', 'Sheet2', 'Sheet3'].each do |sheet_name|

          describe sheet_name do

            let(:sheet) { workbook[sheet_name] }

            it 'Column names' do
              sheet.column_names.must_equal columns_by_sheet[sheet_name]
            end

            it 'Rows' do
              sheet.to_a.must_equal expected_rows(sheet)
            end

          end

        end

      end

    end

  end

end