require 'minitest_helper'

describe ExcelUtils do

  let(:tmp_path) { File.expand_path '../../tmp', __FILE__ }

  let(:filename) { File.join tmp_path, "#{SecureRandom.uuid}.xlsx" }

  before do
    FileUtils.rm_rf tmp_path
    FileUtils.mkpath tmp_path
  end

  it 'Write' do
    long_url = 'https://external.xx.fbcdn.net/safe_image.php?d=AQCY1f77h1RtuFfa&w=720&h=720&url=https%3A%2F%2Fwww.mercadolibre.com%2Fjms%2Fmla%2Flgz%2Fbackground%2Fsession%2Farmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%3Fbackground%3Darmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%26message%3DeyJqc190eXBlIjoianNfY29va2llIiwidmFsdWUiOiJ4In0&cfs=1&_nc_hash=AQB9Sf1HyWCwCVhF'
    long_text = 'AQCY1f77h1RtuFfa&w=720&h=720&url=https%3A%2F%2Fwww.mercadolibre.com%2Fjms%2Fmla%2Flgz%2Fbackground%2Fsession%2Farmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%3Fbackground%3Darmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%26message%3DeyJqc190eXBlIjoianNfY29va2llIiwidmFsdWUiOiJ4In0&cfs=1&_nc_hash=AQB9Sf1HyWCwCVhF'

    data = {
      'Sheet 1' => [
        {'column_a' => 1.5, 'column_b' => 'text 1', 'column_c' => DateTime.parse('2019-05-25T16:30:00')},
        {'column_b' => 'text 2', 'column_c' => Date.parse('2019-07-09'), 'column_a' => 2},
      ],
      'Sheet 2' => [
        {'Column A' => long_url, 'Column B' => long_text}
      ],
      'Sheet 3' => []
    }

    ExcelUtils.write filename, data

    workbook = ExcelUtils.read filename

    workbook.to_h.must_equal data
  end

end