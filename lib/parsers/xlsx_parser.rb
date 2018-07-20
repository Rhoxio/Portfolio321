require 'rubyXL'

class XLSXParser

  def self.parse(path)
    # Can do 
    workbook = RubyXL::Parser.parse(path)

  end

end