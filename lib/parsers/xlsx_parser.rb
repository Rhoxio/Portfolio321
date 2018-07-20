class XLSXParser

  def self.parse(path)
    # Can do 
    workbook = RubyXL::Parser.parse("path/to/Excel/file.xlsx")
  end

end