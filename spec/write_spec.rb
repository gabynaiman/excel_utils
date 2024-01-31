require 'minitest_helper'

describe ExcelUtils do

  def assert_wrote(data, expected=data)
    filename = tmp_path "#{SecureRandom.uuid}.xlsx"

    ExcelUtils.write filename, data

    workbook = ExcelUtils.read filename

    expected = {'Sheet1' => expected} if expected.is_a?(Array)

    workbook.to_h.must_equal expected
  end

  describe 'Single sheet' do

    it 'Basic data' do
      rows = [
        {'column_a' => '1', 'column_b' => 'text 1'},
        {'column_a' => '2.5', 'column_b' => nil, 'column_c' => 'text 2'},
      ]

      assert_wrote rows, rows.map { |r| {'column_c' => nil}.merge r }
    end

    it 'Symbolized column names' do
      rows = [
        {column_a: 'text 1'},
        {column_a: 'text 2'},
      ]

      assert_wrote rows, rows.map { |r| {'column_a' => r[:column_a]} }
    end

    it 'Long values' do
      long_url = 'https://external.xx.fbcdn.net/safe_image.php?d=AQCY1f77h1RtuFfa&w=720&h=720&url=https%3A%2F%2Fwww.mercadolibre.com%2Fjms%2Fmla%2Flgz%2Fbackground%2Fsession%2Farmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%3Fbackground%3Darmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%26message%3DeyJqc190eXBlIjoianNfY29va2llIiwidmFsdWUiOiJ4In0&cfs=1&_nc_hash=AQB9Sf1HyWCwCVhF'
      long_text = 'AQCY1f77h1RtuFfa&w=720&h=720&url=https%3A%2F%2Fwww.mercadolibre.com%2Fjms%2Fmla%2Flgz%2Fbackground%2Fsession%2Farmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%3Fbackground%3Darmor.c27f26204365785956aaf9b30c289d0d4b0a201b4b01dfc3d83162b82608973be9a82ac85c2097c3f23b522681adca38b143f5e1a10fda251b430051a1592bf19caba204c07179120c48d47a1ceb5a45.774f4b8036d4b909101bab92870d262a%26message%3DeyJqc190eXBlIjoianNfY29va2llIiwidmFsdWUiOiJ4In0&cfs=1&_nc_hash=AQB9Sf1HyWCwCVhF'

      rows = [
        {'column_a' => long_url},
        {'column_a' => long_text},
      ]

      assert_wrote rows
    end

    it 'Numbers' do
      rows = [
        {'column_a' => 1},
        {'column_a' => 2.5},
      ]

      assert_wrote rows
    end

    it 'Date and DateTime' do
      rows = [
        {'column_a' => DateTime.parse('2019-05-25T16:30:00')},
        {'column_a' => Date.parse('2019-07-09')},
      ]

      assert_wrote rows
    end

    it 'Objects' do
      rows = [
        {'column_a' => ['text', 1]},
        {'column_a' => {key: 123}},
        {'column_a' => Set.new(['text', 1])}
      ]

      assert_wrote rows, rows.map { |r| {'column_a' => r['column_a'].to_s} }
    end

    it 'Booleans' do
      rows = [
        {'column_a' => true, 'column_b' => 'text 1'},
        {'column_a' => false, 'column_b' => 'text 2'}
      ]

      assert_wrote rows, [{'column_a' => 'true', 'column_b' => 'text 1'}, {'column_a' => 'false', 'column_b' => 'text 2'}]
    end

  end

  it 'Multiple sheets' do
    data = {
      'strings' => [
        {'column_a' => 'text 1', 'column_b' => 'text 2'},
        {'column_a' => 'text 3', 'column_b' => 'text 4'}
      ],
      'numbers' => [
        {'x' => 1, 'y' => 2},
        {'x' => 3, 'y' => 4},
        {'x' => 5, 'y' => 6}
      ],
      'empty' => [],
    }

    assert_wrote data
  end

end
