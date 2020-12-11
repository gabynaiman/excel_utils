require 'minitest_helper'

describe ExcelUtils, 'Read' do

  def expected_rows(sheet)
    rows_by_sheet[sheet.name].map { |r| Hash[columns_by_sheet[sheet.name].zip(r)] }
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

  [['xls', 'batch'], ['xlsx', 'batch'], ['xlsx', 'stream']].each do |extension, strategy|

    describe "#{extension} #{strategy}" do
  
      let(:filename) { File.expand_path "../sample.#{extension}", __FILE__ }

      describe 'Original column names' do

        let(:workbook) { ExcelUtils.read filename, iterator_strategy: strategy }
        
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

  it 'Force extension' do
    filename = File.expand_path "../sample.tmp", __FILE__
    workbook = ExcelUtils.read filename, extension: 'xlsx'
    workbook.sheets.count.must_equal 3
  end

  describe 'Worksheet Iterators' do

    let(:filename) { File.expand_path "../large.xlsx", __FILE__ }

    def print_memory(prefix)
      pid, memory_usage = `ps ax -o pid,rss | grep -E "^[[:space:]]*#{$$}"`.strip.split.map(&:to_i)
      puts "[#{prefix}] MEMORY USAGE: #{memory_usage} kilobytes"
    end

    it 'Batch Iterator' do
      print_memory 'batch-BEFORE'
      workbook = ExcelUtils.read filename, iterator_strategy: 'batch'
      count = 0
      workbook['Sheet1'].each do |row|
        count += 1
      end
      print_memory 'batch-AFTER'
      count.must_equal 20000
    end

    it 'Stream Iterator' do
      print_memory 'stream-BEFORE'
      workbook = ExcelUtils.read filename, iterator_strategy: 'stream'
      count = 0
      workbook['Sheet1'].each do |row|
        count += 1
      end
      print_memory 'stream-AFTER'
      count.must_equal 20000
    end

  end

end