module Experiment

  # This makes it so you can call 'driver', 'exporter', or 'xlsx_parser' to read 
  # the variable or object set in initialize, but can't redefine it in this class itself.
  # Being able to set an instance variable requires you to use 'attr_accessor' instead of "attr_reader".

  attr_reader :driver, :exporter, :xlsx_parser
 
  def self.setup()

  	# This method reads the preset Experiment Setup spreadsheet to load experiment parameters, including 
  	# file and directory paths, P123 URLs, and run repeat values.  Generating data file names/path relies
  	# on the spreadsheet access functions formatting pc paths such that '\'s are replaced by '\\'s.
  	 
  	# open the experiment setup spreadsheet
  	# wkbk = RubyXL::Parser.parse(EXPERIMENT_SETUP)
    setup_path = ENV["EXPERIMENT_SETUP"]
    wkbk = RubyXL::Parser.parse(setup_path)
  	wksht = wkbk[0]

  	# run_count_index tracks the current run no. of $runs_todo and is used to create filenames for run results
  	# it must be written to Experiment Setup for each run including a nil value at experiment completion 

  	runs_todo = wksht[2][1].value				                  # read in the number of P123 runs to be performed
    run_count_index = wksht[4][1].value                   # for some reason using "||= 0" fails on this line so add next line
    run_count_index.value = run_count_index.value ||= 0   # 0 unless resuming an interrupted series of runs

  	# read in the path to store experiment data and convert '\\' to '/'
    experiment_folder = convert_pc_path(wksht[1][1].value)
    ap experiment_folder

    # extract the experiment name from the end of of the path
    idx = 0are
    chrs = experiment_folder.chars    
    chrs.each_index { |x| if chrs[x] == '/' then idx = x end }
    idx+=1
    aname = chrs[idx..-1]
    experiment_name = aname.join

    # build the Results and Todo file paths
  	results_path = experiment_folder + '/0-' + experiment_name + '_Results.xlsx'
  	todo_path = experiment_folder + '/1-' + experiment_name + '_Todo.xlsx'

  	# create the base run result file name with path: run number to be appended
  	# experiment_folder/run_data_filename-01, -02..runs_todo
  	run_data_filename = wksht[3][1].value ||= experiment_name
  	base_runfile = experiment_folder + '/' + run_data_filename + '-'

  	# also read in P123 URLs needed for rank weight access, universe selection, etc.
  	# may include backtest results URL?

  end

  def self.convert_pc_path(path_string)

    # converts a pc format path like "C:\\Users\\Scott" to "C:/Users/Scott"
    char_array = path_string.chars
    char_array.each_index { |x| if char_array[x] == "\\" then char_array[x] = "/" end }
    path_string = char_array.join
    ap path_string
    
    # does this alter parameter path_string itself or just return converted string?
  end

  def self.run_file_path(path_string)

    path = path_string + if$run_count_index < 10 then "0" else "" end + run_count_index.to_s + ".xlsx"

    # does this alter parameter path_string itself or just return converted string?
  end

end
