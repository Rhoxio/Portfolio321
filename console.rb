require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'rubyXL'
require 'nokogiri'
require 'pry'

Dir["./lib/**/*.rb"].each {|file| require file }

$driver = Selenium::WebDriver.for :chrome
$wait = Selenium::WebDriver::Wait.new(:timeout => 15)

# Experiment.setup()
#    wkbk = RubyXL::Parser.parse("C:/Users/Scott/GitHub_Jobs/Experiments/Experiment Setup.xlsx")
 # 	wksht = wkbk[0]
  #	runs_todo = wksht.sheet_data[2][1]				# read in the number of P123 runs to be performed
  #	run_count_index = wksht.sheet_data[4][1] #||= 0	# 0 unless resuming an interrupted series of runs
    # read in the experiment data folder path and extract the experiment name from the end of it
  	# use these pieces to construct the todo and results file names/paths

  #	experiment_folder = wksht.sheet_data[1][1]  	
  #	experiment_name = "Experiment A"   # $experiment_folder[$experiment_folder('\\')+1..-1]   # find last '\' ('\\') in the path and extract the last folder name 
  #	results_path = experiment_folder + '\0-' + experiment_name + ' Results.xlsx'
 # 	todo_path = experiment_folder + '\1-' + experiment_name + ' Todo.xlsx'

  	# create the base run result file name with path: run number to be appended
  	# experiment_folder//run_data_filename-01, -02, ..., runs_todo

 # 	run_data_filename = wksht.sheet_data[3][1] ||= experiment_name
 # 	base_runfile = experiment_folder + '\\' + run_data_filename + '-'

# This opens up an IRB prompt that run on the same line as 'pry' is evoked. 
# You can also use this to debug code within the project.
pry
