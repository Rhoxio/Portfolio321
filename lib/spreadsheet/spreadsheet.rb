class Spreadsheet
################################################################################################
# this class manages all spreadsheet activities for an experiment, a series of runs or backtests:
#   - getting experiment Setup values entered by user
#   - creating experiment Todo file of node weight and unverse selections for each run
#   - creating experiment Analysis file from template containing results analysis algorithms
#   - managing all state variable that guide/track progress through experiment's runs
#   - gathering results for each run and logging them into files 
#   - transferring key results to Analysis for collation and analysis processing
################################################################################################

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
  #  idx = 0
  # chrs = @experiment_folder.chars 
  #  chrs.each_index { |x| if chrs[x] == '/' then idx = x end }
  #  idx += 1
  #  aname = chrs[idx..-1]
  #  @experiment_name = aname.join
    @experiment_name = @experiment_folder.split('/').last

    # build the Results and Todo file paths
    @todo_filepath = @experiment_folder + '/1-Todo_' + @experiment_name + '.xlsx'
    @todo_wkbk = if experiment_file_exists?(@todo_filepath) 
      then RubyXL::Parser.parse(@todo_filepath)
      else RubyXL::Workbook.new
      end
    @todo_wksht = @todo_wkbk[0]

    # prepare Analysis filepath   
    @analysis_filepath = @experiment_folder + '/0-Analysis_' + @experiment_name + '.xlsx'

    # open Analysis Template to hold Analysis data later saved to Analysis file
    @analysis_template = convert_pc_path(@setup_cells[:analysis_template].value)
    @analysis_wkbk = if experiment_file_exists?(@analysis_filepath) 
      then RubyXL::Parser.parse(@analysis_filepath)
      else RubyXL::Workbook.new(@analysis_template)
      end
    @analysis_wksht = @analysis_wkbk[0]

    # create the base run result file name with path to the experiment's data folder
    # path will be updated as each run completes: results_filename-01, -02..runs_todo
    @results_filename = @setup_cells[:results_filename].value
    if @results_filename == "experiment" then @results_filename = @experiment_name end
    @base_results_filepath = @experiment_folder + '/' + @results_filename + '-'
   	
    # read in which type of P123 run results report to download
    @results_report = @setup_cells[:results_report].value

    # this is a scratchpad workbook that needn't be checked for already existing
    @results_wkbk = RubyXL::Workbook.new
    @results_wksht = @results_wkbk[0]

    # read in P123 URLs needed to access rank weight access, universe selection
    # may include backtest results URL?
#    @p123_rank_url = @setup_cells[:P123_rank_URL].value
#    @p123_screens_url = @setup_cells[:P123_screens_URL].value


  	# ****  FAKE A RUN to test methods  ****
#    log_run_done()
#    write_run_results(mock_test_table)
#    results = nab_results_file(@results_report)
#    write_run_results(results)
  end

  def get_p123_urls()
  # write results file and add performance results to Analysis file

   # Setup always contains last run completed because run# increments at run start
   p123_urls = Array.new
   p123_urls[0] = @setup_cells[:P123_rank_URL].value 
   p123_urls[1] = @setup_cells[:P123_screens_URL].value
   return p123_urls
  end

  def nab_results_file(from="chart")	
  # nabs a .csv run results file in mid-download as Open/Save dialog box waits for action choice
  
    # create the search filepath for any '.xls.part' files in the Users */Temp directory
    file_path = convert_pc_path(@setup_cells[:nab_filename].value) + "/*.xls.part"

     # get a list of files in directory with suitable extension
    rezults = [""]
    cnt = 0
    shazam = Dir.glob(file_path)
	   # for now assume only one candidate file in directory - consider failsafes later
    File.open(shazam[0]).each { |record| rezults[cnt] = record; cnt+=1 }

    # rezults[] is now an array of tab delimited values with a "\n" last character
    # remove "\n" and split string at "\t"
    rezults.each_index { |num| rezults[num] = rezults[num].gsub("\n", "") }
    rezults.each_index { |num| rezults[num] = rezults[num].split("\t") }

    # convert all results except date to numbers - date in 1st column of chart download, 2nd column of table
    skip = if from != "chart"  then 1 else 0 end
    rezults.each_index { |line| rezults[line].each_index { |cnt|  
      if line > 0 then
        if cnt != skip then rezults[line][cnt] = rezults[line][cnt].to_f end  
      end
    }
   }
  return (rezults)
  end

  def log_run_done()
   # write results file and add performance results to Analysis file

   # Setup always contains last run completed because run# increments at run start
   @setup_cells[:current_run] = @current_run
  end

  def setup_todo(node_names, universe_names)
  # set up the Todo file with rank node weights and universe selections for each run

    # instantiate for later use whether resuming or not
    if experiment_file_exists?(@todo_filepath) then return end            # *********

    # put node names in column 0 of Todo for human reference
    fill_clm_values(0, 0, @todo_wksht, node_names)

    # generate a column of randomly selected node weights in Todo for each run of the experiment
    # each run's weights are in column[@curr_run] 
    node_weights = Array.new(node_names.length, 0)                # initialize weight set to 0
    cnt = 1
    while cnt <= @runs_todo
      node_weights = set_node_weights(node_weights.length)        # get a set of node weights
      fill_clm_values(0, cnt, @todo_wksht, node_weights, true)		# load set into Todo
      cnt += 1
    end
  
    # add a column past the weights columns in Todo that lists universes for each run
    # universes listed by name for human readability
    universe_list = Array.new(@runs_todo, "")
    universe_list.each_index { |idx| universe_list[idx] = universe_names.sample }
#    cnt = 0
#    while cnt < @runs_todo 
#      universe_list[cnt] = universe_names.sample    
#      cnt += 1
#    end
    fill_clm_values(0, @runs_todo + 2, @todo_wksht, universe_list, true)    # load set into Todo

    # write node weights and universes into Todo spreadsheet
    @todo_wkbk.write(@todo_filepath)
  end

  def fill_clm_values(row, clm, worksheet, indata, buffer = false)
	# Add values to a column in the specified worksheet of a workbook

    # Fill a worksheet column with data values starting at given row
    rho = row
    indata.each { |x| 
       cell_data = worksheet.add_cell(rho,clm,x)		# RubyXL: x.add_cell works, x.change_contents doesn't
       rho += 1
    }
    if buffer != 'nil' then cell_data = worksheet.add_cell(rho,clm, nil)	end # buffer column of entries with nil cell
  end

  def set_node_weights(node_cnt)
  # fill a random set of node weights until weights total 100 (percent): fixed 20%/node for now

    weights = Array.new(node_cnt, 0)  # clear weights to zeros

   	#same weight may randowmly by picked twice: do until 5 separate weights are set
   	while weights.inject(:+) < 100
   		weights[rand(node_cnt) - 1] = 20	# convert random number{1..X} to random index{0..(X-1)}
    end
    return weights
  end

  def experiment_file_exists?(filepath="")
  # if resuming, detects a file created when the experiment was set up initially 

    file_name = filepath.split('/').last    # extra the file name from filepath

    # remove file name from filepath to get directory path
    dirctry = filepath
    dirctry = dirctry.sub(filepath.split('/').last, "")

    # get the list of files in the directory and see if filename is on it  
    file_list = Dir.entries(dirctry)
    exists = false
    file_list.each { |fname| if fname == file_name then exists = true end }
    return exists
  end

  def if_runs_todo()
  # checks for remaining runs and can catch bogus run conditions due to bad Setup values

    if @runs_todo > 0 && @runs_todo >= @current_run then return true else return false end
  end

  def get_next_run_todo(node_weights)
  # return the node weights and universe for the upcoming run

    # weights get returned through the calling parameter array
    node_weights.each_index { |idx| node_weights[idx] = @todo_wksht[idx][@current_run].value }
    # return with this run's universe name gotten from Todo
    universe_name = @todo_wksht[@current_run - 1][@runs_todo + 2].value
    return universe_name
  end
  
  def incr_run_todo()
  # update the run counter to upcoming run number

    @current_run += 1
  end
  
  def convert_pc_path(path_string)
	# converts a pc format path like "C:\\Users\\Scott" to "C:/Users/Scott"

    char_array = path_string.chars
    char_array.each_index { |x| if char_array[x] == "\\" then char_array[x] = "/" end }
    path_string = char_array.join
  end

  def write_run_results(results_table)
    # writes run results to Excel file

    results_filepath = run_results_filepath(@base_results_filepath)
    results_table.each_index { |row|
      results_table[row].each_index { |clm| @results_wksht.add_cell(row, clm, results_table[row][clm])
      }
    }
    @results_wkbk.write(results_filepath)
  end

  def run_results_filepath(path_string)
  # complete the filepath for the current run results file by appending the run number
 
    # add a leading '0' for run numbers < 10
    path = path_string + if @current_run < 10 then "0" else "" end + @current_run.to_s + ".xlsx"
  end

  def mock_test_table()
  #generate a small table for testing xlsx array writing

    table = Array.new (4) { Array.new(3, 1) }
    b=0
    table.each { |x| table[b] = [b,b+2,b*3]; b+=1 } 
  end

  # ****************
  #  PRIVATE METHODS
  # ****************
  # private methods not needed so much now but want placeholder as class grows
  private


  def experiment_setup
   {path: ENV["EXPERIMENT_SETUP_FILEPATH"]}
  end

  def setup_file_cel1s
    { current_run:        @setup_wksht[1][1] ,
      runs_todo:          @setup_wksht[3][1] ,
      experiment_folder:  @setup_wksht[4][1] ,
      analysis_template:  @setup_wksht[5][1] ,
      nab_filename:       @setup_wksht[6][1] ,
      results_report:     @setup_wksht[7][1] ,
      results_filename:   @setup_wksht[8][1] ,
      P123_rank_URL:      @setup_wksht[10][1] ,
      P123_screens_URL:   @setup_wksht[11][1]
    }
  end
end