# ExcelUtils

[![Gem Version](https://badge.fury.io/rb/excel_utils.svg)](https://rubygems.org/gems/excel_utils)
[![Build Status](https://travis-ci.org/gabynaiman/excel_utils.svg?branch=master)](https://travis-ci.org/gabynaiman/excel_utils)
[![Coverage Status](https://coveralls.io/repos/github/gabynaiman/excel_utils/badge.svg?branch=master)](https://coveralls.io/github/gabynaiman/excel_utils?branch=master)
[![Code Climate](https://codeclimate.com/github/gabynaiman/excel_utils.svg)](https://codeclimate.com/github/gabynaiman/excel_utils)

Excel utils for easy read and write

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'excel_utils'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install excel_utils

## Usage

### Write
```ruby
data = {
  'Sheet1' => [
    {column_a: 1.0, column_b: 'some text'},
    {column_a: 2.0, column_b: 1.35},
    {column_a: 3.0, column_b: Date.parse('2019-08-17')},
    {column_a: 4.0, column_b: nil}
  ],
  'Sheet2' => [
    {'Column 1' => 123.0, 'Column 2' => 'Text 1'},
    {'Column 1' => 456.0, 'Column 2' => 'Text 2'}
  ],
  'Sheet3' => []
}

ExcelUtils.write '/path/to/file.xlsx', data
```

### Read
```ruby
workbook = ExcelUtils.read '/path/to/file.xlsx'
workbook.to_h => # {'Sheet1' => [{'Column A' => 1, ...}, ...], ...}
sheet = workbook['Sheet1']
sheet.name => # 'Sheet1'
sheet.column_names => # ['Column A', ...]
sheet.each # Implements enumerable module (map, to_a, ...)

workbook = ExcelUtils.read '/path/to/file.xlsx', normalize_column_names: true
workbook.to_h => # {'Sheet1' => [{column_a: 1, ...}, ...], ...}
sheet = workbook.sheets.first
sheet.column_names => # [:column_a, ...]
```

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/gabynaiman/excel_utils.

## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).