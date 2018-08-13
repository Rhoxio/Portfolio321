class Spreadsheet

  def initialize()

  	# This method reads the preset Experiment Setup spreadsheet to load experiment parameters, including 
  	# file and directory paths, P123 URLs, and run repeat values.  Generating data file names/path relies
  	# on the spreadsheet access functions formatting pc paths such that '\'s are replaced by '\\'s.

  	# open the experiment setup spreadsheet
    setup = experiment_setup
    setup_wkbk = RubyXL::Parser.parse(setup[:path])
  	@setup_wksht = setup_wkbk[0]
  	# run_done tracks the current run no. of runs_todo and is used to create filenames for run results
  	# written to Experiment Setup at each run completion including a nil value at experiment completion 

    # current_run is initialized to 0 or last run completed, it's incremented when next run is started
    # so it becomes either run 1 or the run# after the last previously completed run (the resume run#)
    @setup_cells = setup_file_cel1s
    @current_run = @setup_cells[:current_run].value ||= 0   # contains current run# to start run over if abend 

    # read in the number of P123 runs to be performed
  	@runs_todo = @setup_cells[:runs_todo].value	

  	# read in the path to store experiment data and convert '\\' to '/'
    @experiment_folder = convert_pc_path(@setup_cells[:experiment_folder].value)

    # extract the experiment name from the end of of the path
    idx = 0
    chrs = @experiment_folder.chars 
    chrs.each_index { |x| if chrs[x] == '/' then idx = x end }
    idx += 1
    aname = chrs[idx..-1]
    @experiment_name = aname.join

    # build the Results and Todo file paths
    @todo_path = @experiment_folder + '/1-Todo_' + @experiment_name + '.xlsx'
    @todo_wkbk = RubyXL::Workbook.new
    @todo_wksht = @todo_wkbk[0]

    # prepare Analysis filepath   
    @analysis_path = @experiment_folder + '/0-Analysis_' + @experiment_name + '.xlsx'

    # open Analysis Template to hold Analysis data later saved to Analysis file
    @analysis_template = convert_pc_path(@setup_cells[:analysis_template].value)
    @analysis_wkbk = RubyXL::Parser.parse(@analysis_template)
    @analysis_wksht = @analysis_wkbk[0]

  	# create the base run result file name with path to the experiment's data folder
  	# path will be updated as each run completes: results_filename-01, -02..runs_todo
  	@results_filename = @setup_cells[:results_filename].value # ||= @experiment_name
  	@base_results_path = @experiment_folder + '/' + @results_filename + '-'

    @results_wkbk = RubyXL::Workbook.new
    @results_wksht = @results_wkbk[0]

  	# read in P123 URLs needed to access rank weight access, universe selection
  	# may include backtest results URL?
    @p123_rank_url = @setup_cells[:P123_rank_URL].value
    @p123_screens_url = @setup_cells[:P123_screens_URL].value

    # ****  FAKE A RUN to test results file writing  ****
    log_run_done()
    write_run_results(mock_test_table())
binding.pry

  end


  def incr_run_todo()

   # update the run counter, write results file and add performance results to Analysis
    @current_run += 1
  end
  

  def if_runs_todo()
    # this method checks for remaining runs and can catch bogus run conditions due to bad Setup values

    if @runs_todo > 0 && @runs_todo >= @current_run then return true else return false end
  
  end


  def log_run_done()

    # write results file and add performance results to Analysis file
    # Setup always contains last run completed because run# increments at run start
    @setup_cells[:current_run] = @current_run
    
  end

  def convert_pc_path(path_string)

    # converts a pc format path like "C:\\Users\\Scott" to "C:/Users/Scott"
    char_array = path_string.chars
    char_array.each_index { |x| if char_array[x] == "\\" then char_array[x] = "/" end }
    path_string = char_array.join
    
  end

  def write_run_results(results_table)

    results_file = run_results_path(@base_results_path)
#    @results_wkbk.@source_file_path = results_file       # update internal spreadsheet variable to match latest filename?
    results_table.each_index { |row|
      results_table[row].each_index { |clm| @results_wksht.add_cell(row, clm, results_table[row][clm])
      }
    }
   @results_wkbk.write(results_file)

  end


  def run_results_path(path_string)

    # add a leading '0' for run numbers < 10
    path = path_string + if @current_run < 10 then "0" else "" end + @current_run.to_s + ".xlsx"

  end

  def mock_test_table()

    #generate a small table for testing xlsx array writing
    table = Array.new (4) { Array.new(3, 1) }
    b=0
    table.each { |x| table[b] = [b,b+2,b*3]; b+=1 } 
   # return table

  end

  # ****************
  #  PRIVATE METHODS
  # ****************

  # private methods not needed so much now but want placeholder at class grows
  private


  def experiment_setup
    {path: ENV["EXPERIMENT_SETUP"]}
  end

  def setup_file_cel1s
    { current_run:        @setup_wksht[1][1] ,
      runs_todo:          @setup_wksht[3][1] ,
      experiment_folder:  @setup_wksht[4][1] ,
      analysis_template:  @setup_wksht[5][1] ,
      results_filename:   @setup_wksht[6][1] ,
      P123_rank_URL:      @setup_wksht[8][1] ,
      P123_screens_URL:   @setup_wksht[9][1] }
  end


end