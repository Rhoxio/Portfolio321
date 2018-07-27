class XLSXParser

  def self.open(path)
    # Opens and reads an entire .xlsx workbook
    workbook = RubyXL::Parser.parse(path)
  end

  def self.new(path)
    # Opens a new blank .xlsx workbook 
    # Should handle file already exists error: success is workbook != to nil?
    workbook = RubyXL::Workbook.new(path)
  end

  def self.write(workbook, path)
  	# Writes out an entire .xlsx workbook
   	workbook.write(path)
  end

  def self.fill_clm_values(row, clm, worksheet, data)
  	# Add values to a column in the specified worksheet of a workbook
  	rho = row
  	wksheet = worksheet
    # Fill a worksheet column with data[] values starting at row
    data.each do |x| 
   	  cell_data = worksheet.add_cell(rho,clm,x)		# RubyXL: x.add_cell works, x.change_contents doesn't
   	  rho += 1
      end
    cell_data = worksheet.add_cell(rho,clm, nil)	# mark end of read column area with nil
  end

  def self.read_clm_values(row, clm, worksheet, data)
  	# Read values from a column in the specified worksheet of a workbook starting at the row specified
  	rho = row
  	wksheet = worksheet
    # Fill a worksheet column with data[] values starting at row
    i=0
    cell_data = 0
    while cell.data != nil 
   	  cell_data = worksheet.sheet_data(rho+i,clm)		
   	  if cell_data != nil
   	  	data[i] = cell_data
   	  	i += 1
   	  end
  end

   def self.set_weights(weights)
   	# set random weights until weights total 100 (percent) 
   	weights.map { |x| x=0 }					# clear weights to zeros
   	rand_max = weights.length				# set size of selection pool
   	#same weight may be set twice: do until 5 separate weights are set
   	while weights.inject(:+) < 100
   		weights[rand(rand_max) - 1] = 20	# convert random number{1..X} to random index{0..(X-1)}
   	end
  end


end