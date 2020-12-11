module ExcelUtils
  class Sheet

    include Enumerable

    attr_reader :name

    def initialize(name, iterator)
      @name = name
      @iterator = iterator
    end

    def column_names
      @column_names ||= iterator.column_names
    end

    def each(&block)
      iterator.each(&block)
    end

    private

    attr_reader :iterator

  end
end