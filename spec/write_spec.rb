require 'minitest_helper'

describe ExcelUtils do

  let(:tmp_path) { File.expand_path '../../tmp', __FILE__ }

  let(:filename) { File.join tmp_path, "#{SecureRandom.uuid}.xlsx" }

  before do
    FileUtils.rm_rf tmp_path
    FileUtils.mkpath tmp_path
  end

  it 'Write' do
    data = {
      'Sheet 1' => [
        {'column_a' => 1.5, 'column_b' => 'text 1', 'column_c' => DateTime.parse('2019-05-25T16:30:00')},
        {'column_b' => 'text 2', 'column_c' => Date.parse('2019-07-09'), 'column_a' => 2},
      ],
      'Sheet 2' => [
        {'Column A' => 'Text A', 'Column B' => 'Text B'}
      ],
      'Sheet 3' => []
    }

    ExcelUtils.write filename, data

    workbook = ExcelUtils.read filename

    workbook.to_h.must_equal data
  end

end