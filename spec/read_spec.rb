require 'minitest_helper'

describe ExcelUtils, 'Read' do

  Dir.glob(File.join(RESOURCES_PATH, 'basic.*')).each do |filename|

    [true, false].each do |normalize_column_names|

      describe File.basename(filename), "Workbook (normalize_column_names: #{normalize_column_names})" do

        let(:workbook) { ExcelUtils.read filename, normalize_column_names: normalize_column_names }

        let(:sheet) { workbook.sheets.first }

        let(:csv?) { File.extname(filename) == '.csv' }

        let(:expected_sheet_name) { csv? ? 'default' : 'Sheet1' }

        let(:column_a) { normalize_column_names ? :column_a : 'Column A' }

        let(:column_b) { normalize_column_names ? :column_b : 'Column B' }

        let(:expected_columns) { [column_a, column_b] }

        let :expected_rows do
          if csv?
            [
              {column_a => '1', column_b => 'some text'},
              {column_a => '2', column_b => '1,35'},
              {column_a => '3', column_b => '17/08/2019'},
              {column_a => '4', column_b => nil}
            ]
          else
            [
              {column_a => 1, column_b => 'some text'},
              {column_a => 2, column_b => 1.35},
              {column_a => 3, column_b => Date.parse('2019-08-17')},
              {column_a => 4, column_b => nil}
            ]
          end
        end

        it 'filename' do
          workbook.filename.must_equal filename
        end

        it 'normalize_column_names' do
          workbook.normalize_column_names.must_equal normalize_column_names
        end

        it 'sheets' do
          workbook.sheets.count.must_equal 1
          workbook[workbook.sheets.first.name].must_equal sheet
        end

        it 'to_h' do
          workbook.to_h.must_equal workbook.sheets.first.name => sheet.to_a
        end

        describe 'Sheet' do

          it 'name' do
            sheet.name.must_equal expected_sheet_name
          end

          it 'normalize_column_names' do
            sheet.normalize_column_names.must_equal workbook.normalize_column_names
          end

          it 'column_names' do
            sheet.column_names.must_equal expected_columns
          end

          it 'count' do
            sheet.count.must_equal 4
          end

          it 'to_a' do
            sheet.to_a.must_equal expected_rows
          end

        end

      end

    end

  end

  it 'empty.csv' do
    workbook = ExcelUtils.read resource_path('empty.csv')

    workbook.sheets.map(&:name).must_equal ['default']

    workbook['default'].column_names.must_equal []
    workbook['default'].to_a.must_equal []

    workbook.to_h.must_equal 'default' => []
  end

  it 'only_headers.csv' do
    workbook = ExcelUtils.read resource_path('only_headers.csv')

    workbook.sheets.map(&:name).must_equal ['default']

    workbook['default'].column_names.must_equal ['ID', 'Value']
    workbook['default'].to_a.must_equal []

    workbook.to_h.must_equal 'default' => []
  end

  it 'custom_extension.tmp' do
    workbook = ExcelUtils.read resource_path('custom_extension.tmp'), extension: 'xlsx'

    expected_rows = [
      {'ID' => 1, 'Value' => 'Text 1'},
      {'ID' => 2, 'Value' => 'Text 2'}
    ]

    workbook.sheets.map(&:name).must_equal ['Sheet1']

    workbook['Sheet1'].column_names.must_equal ['ID', 'Value']
    workbook['Sheet1'].to_a.must_equal expected_rows

    workbook.to_h.must_equal 'Sheet1' => expected_rows
  end

  it 'multiple.xlsx' do
    workbook = ExcelUtils.read resource_path('multiple.xlsx'), normalize_column_names: true

    workbook.sheets.map(&:name).must_equal ['Sheet1', 'Sheet2', 'Sheet3']

    expected_rows_sheet_1 = [
      {id: 1, value: 'Text 1'},
      {id: 2, value: 'Text 2'}
    ]

    workbook['Sheet1'].column_names.must_equal [:id, :value]
    workbook['Sheet1'].to_a.must_equal expected_rows_sheet_1

    workbook['Sheet2'].column_names.must_equal [:id, :value]
    workbook['Sheet2'].to_a.must_equal []

    workbook['Sheet3'].column_names.must_equal []
    workbook['Sheet3'].to_a.must_equal []

    workbook.to_h.must_equal 'Sheet1' => expected_rows_sheet_1,
                             'Sheet2' => [],
                             'Sheet3' => []
  end

end