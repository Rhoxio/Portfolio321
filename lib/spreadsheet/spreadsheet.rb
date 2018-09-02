class Spreadsheet
#################################################################################################
# This class manages all spreadsheet activities for an experiment, a series of runs or backtests:
#   - getting experiment Setup values entered by user
#   - creating experiment Todo file listing node weight and unverse selections for each run
#   - creating experiment Analysis file from formula template file containing analysis algorithms
#   - managing all state variables that guide/track progress through the experiment's runs
#   - gathering results for each run and logging them files 
#   - transferring key results to Analysis for analysis processing  
# The term 'filepath' is used here to mean a path that includes a whole or base filename
# A 'base' filename is something like 'Run - ' to which a number is appended (e.g., run # '07')
# See nab_results_file() to read how downloading P123 backtest (run) results is managed
# A separate utility 'Clear Excel.rb' is provided to clear the slate when the program bombs out 
# with an unforseen or hard to manage error. It handles Excel instance issues by brute force.  
#################################################################################################

  def initialize()

  # This method reads the preset Experiment Setup spreadsheet to load experiment parameters, including 
  # file and directory paths, P123 URLs, and run repeat values.  Generating data file names/path relies
  # on the spreadsheet access functions formatting pc paths such that '\'s are replaced by '\\'s.

    # create a specific Excel instance for experiment to avoid affecting other user's open spreadsheets
    # set all experiment files to be visible so it's easy to close them manually if program (errors out)
    # default to calculation manual mode in all cases even though this is only an Analysis issue
    @p123_excel = RobustExcelOle::Excel.create(:reuse => true)
    @p123_excel.for_this_instance(:visible => true, :calculation => :manual)

    # open the experiment setup workbook and spreadsheet
    setup = experiment_setup()          # setup <= Setup filepath .env declaration
    @setup_filepath = setup[:path]
    wkbk_sht_re_open("setup", @setup_filepath)

    # current_run is set to 0 or last run completed (resuming), it's incremented at the next run start
    # so it becomes either run 1 or the run# after the last previously completed run (the resume run#)
    @current_run = @setup_cells[:current_run].value ||= 0 

    # read in the number of P123 runs to do and use it to set the col index for universe names
    @runs_todo = @setup_cells[:runs_todo].value
    @universe_todo_col = @runs_todo + 2 

    # read in the path to store experiment data and convert any '\\' to '/'
    @experiment_folder = convert_pc_path(@setup_cells[:experiment_folder].value)

    # extract default experiment name from the end of the experiment data path
    # use experiment name to build Todo, Analysis, and base Results filepaths
    @experiment_name = @experiment_folder.split('/').last 
    @todo_filepath = @experiment_folder + '/1-Todo_' + @experiment_name + '.xlsx'
    @analysis_filepath = @experiment_folder + '/0-Analysis_' + @experiment_name + '.xlsx'
    @results_filename = @setup_cells[:results_filename].value

    # check that Setup specifies an alternate experiment name for run files, else use default
    if @results_filename == "experiment" then @results_filename = @experiment_name end
    @base_results_filepath = @experiment_folder + '/' + @results_filename + '-'
    
    resuming_experiment?()    # IMPORTANT: must build Analysis filepath before this call

    wkbk_sht_re_open("todo", @todo_filepath)  # open new workbook  

    if @resuming 
      then  # resuming experiment: open existing Analysis with data collected so far 
        wkbk_sht_re_open("analysis", @analysis_filepath)
      else # starting new experiment: open Template file with formulae and save as Analysis
        analysis_template_filepath = convert_pc_path(@setup_cells[:analysis_template].value)
        wkbk_sht_re_open("analysis", analysis_template_filepath) 
        @analysis_wkbk.save_as(@analysis_filepath)
      end

    # read whether to download a P123 chart or table run results report
    @results_report = @setup_cells[:results_report].value
    if @results_report != "table" then @results_report = "chart" end 

    # create the search filepath for picking up interim P123 interim download files
    @pickup_filepath = convert_pc_path(@setup_cells[:nab_file_path].value) + "/*.xls"

    close_wkbks()
  end   # of initialize()

            ####################################### API Methods #########################################
            #############################################################################################

  def get_p123_urls()
  # experiment API method: return the p123 rank system and screen Urls listed in Setup

    # Setup always contains last run completed because run# increments at run start
    wkbk_sht_re_open("setup")
    p123_urls = Array.new
    p123_urls[0] = @setup_cells[:P123_rank_URL].value 
    p123_urls[1] = @setup_cells[:P123_screens_URL].value
    close_wkbks()
    return p123_urls
  end

  def results_type()
  # Experiment API method: reports if Setup specified a p123 chart or table results report
    return @results_report
  end

  def if_runs_todo()
  # experiment API method: checks if all runs done; may find Setup errors with improper run conditions
  # @runs_todo is incremented after this check so test is '>' not '>='

    if @runs_todo > 0 && @runs_todo > @current_run then return true else return false end
  end
 
  def incr_run_todo()
  # experiment API method: update the run counter to upcoming run number

    @current_run += 1
    puts Time.now.strftime("%r - Run " + @current_run.to_s)
  end
 
  def setup_todo(node_names, universe_names)
  # experiment API method
  # fill in the Todo file with rank node weights and universe selections for each run

    # just return if resuming because Todo was set up during previous experiment start
    if @resuming then return else end

    wkbk_sht_re_open("todo")

    # put node names in column 1 of Todo for human reference
    fill_Excel_clm_values(1, 1, @todo_wksht, node_names)

    # for each run of the experiment write a column of randomly selected node weights into Todo
    # each run's weights are in put in column[@current_run+1] because node names take up col 1
    node_weights = Array.new(node_names.length, 0)                # initialize weight set to 0
    cnt = 1
    while cnt <= @runs_todo
      node_weights = set_node_weights(node_weights.length)        # get a set of node weights
      fill_Excel_clm_values(1, cnt+1, @todo_wksht, node_weights)  # load set into Todo
      cnt += 1
    end
  
    # add a column that lists universes for each run right of the last weights column in Todo
    # universes listed by name for human readability
    universe_list = Array.new(@runs_todo, "")
    universe_list.each_index { |idx| universe_list[idx] = universe_names.sample }
    fill_Excel_clm_values(1, @universe_todo_col, @todo_wksht, universe_list)    # load set into Todo

    # save Todo with the new node weights and universes for every run in the experiment
    @todo_wkbk.save()
    close_wkbks()
  end   # of setup_todo

  def check_run_results_dir()
  # experiment API method called just prior to run start: helps to correctly id the upcomeing interim
  # results download file by logging the number of .xls files in the P123/Google download directory

    @prerun_xls_file_cnt = Dir[@pickup_filepath].length
  end

  def record_run_results()
  # experiment API method
  # write results file and add performance results to Analysis file
  # an interim tab-delimited version of the results file is download is placed on the
  # hard drive before [view|save|cancel] download is selected

    # capture latest run results from the interim download file
    # save to permanent file and transfer selected key results to Analysis
    data = nab_results_file()
    log_analysis_results()    # analyze new Analysis data and save results
  end

  def log_run_done()
  # experiment API file:  logs number of run just completed to Setup to cue what   
  # number run to resume with if a program failure occurs during the next run

    wkbk_sht_re_open("setup")
    @setup_cells[:current_run].value = @current_run
    @setup_wkbk.save()
    close_wkbks()
  end

            ###################################### Setup Methods ########################################
            #############################################################################################

  def set_node_weights(node_cnt)
  # fill a random set of node weights until weights total 100 (percent): fixed 20%/node for now

    weights = Array.new(node_cnt, 0)  # clear weights to zeros

    #same weight may randowmly by picked twice: do until 5 separate weights are set
    while weights.inject(:+) < 100
      weights[rand(node_cnt) - 1] = 20  # convert random number{1..X} to random index{0..(X-1)}
    end
    return weights
  end

  def resuming_experiment?()
  # sets @resuming according to whether or not an Analysis file already exists: should exist only when the
  # experiment was initiated previously - or when mistakenly starting a new experiment with the same name

    file_name = @analysis_filepath.split('/').last    # extract the file name from filepath

    # remove file name from filepath to get directory path
    dirctry = @analysis_filepath
    dirctry = dirctry.sub(@analysis_filepath.split('/').last, "")

    # get the list of files in the directory and see if filename is on it  
    file_list = Dir.entries(dirctry)
    @resuming = false
    file_list.each { |fname| if fname == file_name then @resuming = true end }
  end

            ####################################### Run Methods #########################################
            #############################################################################################

  def get_next_run_todo(node_weights)
  # experimnet API method:  read Todo for the upcoming run's node weights and universe
  # weights are returned through the calling parameter array

    # Todo col 1 contains node names so data col index is current_run + 1
    # idx counts from 0 so must be +1 to match Todo counting rows from 1
    wkbk_sht_re_open("todo")
    node_weights.each_index { |idx| node_weights[idx] = @todo_wksht[idx+1, @current_run+1].value }

    universe_name = @todo_wksht[@current_run, @universe_todo_col].value
    close_wkbks()
    return universe_name
  end

  def nab_results_file() 
  # When a 'save results' key is hit P123 downloads an interim tab-delimited file of results data with 
  # a .xls extension.  P123 then waits for a confirmation to convert download to a true .xls (save) or 
  # to remove it (cancel). This interim file is copied and saved off as a .xlsx, which avoids dealing 
  # with any .xls files or with navigating the P123/Google 'save/cancel download' menu system.

    # continue collecting list of download directory files with .xls extension 
    # until a new one shows up or a timeout occurs (no. of sleep waits exceeds 5 sec)
    file_list = []
    tries = 0
    while file_list.length <= @prerun_xls_file_cnt && tries < 25    
      file_list = Dir[@pickup_filepath]
      sleep(0.2)
      tries =+ 1
    end

    # housekeeping should have deleted all but the latest .xls file but only get the newest file anyway
    @download_filepath = newest_file(file_list)

    # open interim download file and save as results "Run-XX.xlsx" file
    wkbk_sht_re_open("download", @download_filepath)
    @download_wkbk.save_as(run_results_filepath(@base_results_filepath), :if_exists => :overwrite, :if_obstructed => :save) 

    # open run results worksheet to copy and save essential results in Analysis
    wkbk_sht_re_open("analysis")
    copy_and_paste_Excel_clm(@download_wksht.col_range(2), @analysis_wksht.col_range(2))  
    @analysis_wkbk.save()
    close_wkbks() 

    @download_wkbk.close()        # close the saved run results file 'Run-XX'
    File.delete(@download_filepath)   # delete the interim results file from the download directory
  end   # of nab_results()

  def log_analysis_results()
  # analyze most recent run data and copy the results to file's run accumulation area

    wkbk_sht_re_open("analysis")
    @analysis_wkbk.excel.Calculate()  # trigger Analysis calculations

    # transfer the row of values calculated by Analysis to the results accumulation area
    @calcs = Array.new(108, 0)    # size of @calcs must be set for read_Excel_row_section
    read_Excel_row_section(2, 19, @analysis_wksht, @calcs)
    write_Excel_row_section(10+@current_run, 19, @analysis_wksht, @calcs)

    @analysis_wkbk.save()
    close_wkbks()
  end

  def log_experiment_done() 
  # experiment API method:  end of experiment housekeeping

    # zero run completed count in Setup so next experiment doesn't try toresume
    wkbk_sht_re_open("setup")
    @setup_cells[:current_run].value = 0
    @setup_wkbk.save()

   # close experiment workbooks and Excel instance
    close_wkbks()       
    @p123_excel.close() 
  end

            ######################################## File Mgmt ##########################################
            #############################################################################################

  def convert_pc_path(path_string)
  # converts a pc format path like "C:\\Users\\Scott" to "C:/Users/Scott"

    char_array = path_string.chars
    char_array.each_index { |x| if char_array[x] == "\\" then char_array[x] = "/" end }
    path_string = char_array.join
  end

  def wkbk_sht_re_open(file_type, filepath = nil)
  # re-opens a workbook or opens a new one if a filepath is given, then opens first worksheet

    case file_type
    when "Setup", "setup"
      if filepath != nil
        then @setup_wkbk = RobustExcelOle::Workbook.open(filepath, :excel => @p123_excel)
#        then @setup_wkbk = RobustExcelOle::Workbook.unobtrusively(filepath, :rw_change_excel => {:new => @p123_excel})
#          @setup_wkbk.retained_saved()
        else @setup_wkbk.reopen()
      end
      @setup_wksht = @setup_wkbk.first_sheet
      @setup_cells = setup_file_cells()       # @setup_cells always used to access Setup parameters
    when "Todo", "todo"
      if filepath != nil  # Todo is the only workbook that doesn't start with an existing file
        then @todo_wkbk = RobustExcelOle::Workbook.open(filepath, :if_absent => :create, :excel => @p123_excel)
        else 
          @todo_wkbk.reopen()
      end
      @todo_wksht = @todo_wkbk.first_sheet
    when "Analysis", "analysis"
      if filepath != nil
        then @analysis_wkbk = RobustExcelOle::Workbook.open(filepath, :excel => @p123_excel)
        else @analysis_wkbk.reopen()
      end
      @analysis_wksht = @analysis_wkbk.first_sheet
    when "Download", "download"
      if filepath != nil
        then @download_wkbk = RobustExcelOle::Workbook.open(filepath, :excel => @p123_excel)
        else @download_wkbk.reopen()
      end
      @download_wksht = @download_wkbk.first_sheet
    else
      puts "Attempted open of non-existent workbook"
    end
  end

  def newest_file(file_list)
  # returns the most recently created file in a list of files

    # if more than one matching file in the directory list, return the most recently created file
    file = file_list[0]
    if file_list.length > 1 
    then
      idx = 1
      while idx < file_list.length  
        if File.birthtime(file_list[idx]) > File.birthtime(file) then file = file_list[idx] end
        idx +=1
      end
    end
    return file
  end

  def run_results_filepath(path_string)
  # complete the filepath for the current run results file by appending the run number
 
    # add a leading '0' for run numbers < 10
    path = path_string + if @current_run < 10 then "0" else "" end + @current_run.to_s + ".xlsx"
  end

  def close_wkbks() 
 # close workbooks not involved with downloading run results; ok to close if unopened

    @setup_wkbk.close()       # Setup
    @todo_wkbk.close()        # Todo
    @analysis_wkbk.close()    # Analysis
  end

            ####################################### Excel Utils #########################################
            #############################################################################################          

  def fill_Excel_row_values(row, start_col, worksheet, indata)
  # Excel rows & cols start at 1; accept only string entries of integers

    # Fill a worksheet row with a 1D array of data values but do not write the file
    col = start_col
    indata.each { |x|                                 
       cell_data = worksheet[row, col] = x    #   add_cell either adds if cell == nil, else replaces
       col += 1
    }
  end

  def fill_Excel_clm_values(start_row, col, worksheet, indata)
  # enter values to a column of the specified worksheet starting at cell(start_row, col)
  # Excel rows & cols start at 1; OLE utilities accept only string entries

    # Fill a worksheet column with a 1D array of data values but does not save the file
    row = start_row
    indata.each { |x|                             
       worksheet[row, col] = x           
       row += 1
    }
  end

  def copy_and_paste_Excel_clm(from_sheet_col, to_sheet_col)
  # copies an Excel column from one workbook[worksheet] to another

    # fill a worksheet column with a 1D array of data values but do not write the file
    rows = 0
    from_sheet_col.each_with_index { |c, idx| to_sheet_col[idx].value = from_sheet_col[idx].value, rows = idx } 
  end

  def read_Excel_row_section(row, col1, worksheet, row_data)
  # row_data array must be pre-sized for number of cells to read

    col = col1
    row_data.each_index { |cell| 
      row_data[cell] = worksheet[row, col]
      col += 1
    }
  end

  def write_Excel_row_section(row, col1, worksheet, row_data)
  # row_data array must be pre-sized for number of cells to read

    col = col1
    row_data.each_index { |cell|
      worksheet[row, col] = row_data[cell]
      col += 1
    }
  end

  # ****************
  #  PRIVATE METHODS
  # ****************
  # Technically it's likely better to make private any method not described above as an API method

  private

  def experiment_setup
   { path: ENV["EXPERIMENT_SETUP_FILEPATH"] }
  end

  def setup_file_cells
    # Setup cell coordinates for experiment configuration factors
    { current_run:        @setup_wksht[2,2] ,
      runs_todo:          @setup_wksht[4,2] ,
      experiment_folder:  @setup_wksht[5,2] ,
      analysis_template:  @setup_wksht[6,2] ,
      nab_file_path:      @setup_wksht[7,2] ,
      results_report:     @setup_wksht[8,2] ,
      results_filename:   @setup_wksht[9,2] ,
      P123_rank_URL:      @setup_wksht[11,2] ,
      P123_screens_URL:   @setup_wksht[12,2]
    }
  end

end