require 'coverage_helper'
require 'minitest/autorun'
require 'minitest/colorin'
require 'pry-nav'

require 'excel_utils'

RESOURCES_PATH = File.expand_path '../resources', __FILE__

TMP_PATH = File.expand_path '../../tmp', __FILE__
FileUtils.mkpath TMP_PATH unless Dir.exists? TMP_PATH

class Minitest::Test

  def resource_path(relative_path)
    File.join RESOURCES_PATH, relative_path
  end

  def tmp_path(relative_path)
    File.join TMP_PATH, relative_path
  end

end

Minitest.after_run do
  Dir.glob(File.join(TMP_PATH, '*')).each do |filename|
    FileUtils.remove_file filename
  end
end