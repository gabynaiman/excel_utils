require 'coverage_helper'
require 'minitest/autorun'
require 'minitest/colorin'
require 'pry-nav'

require 'excel_utils'

RESOURCES_PATH = File.expand_path '../resources', __FILE__

class Minitest::Test

  def resource_path(relative_path)
    File.join RESOURCES_PATH, relative_path
  end

end