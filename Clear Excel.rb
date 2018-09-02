require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'robust_excel_ole'
require 'nokogiri'
require 'pry'
require 'open-uri'


# !!! USE ONLY AFTER SAVING AND CLOSING ANY OPEN EXCEL FILES YOU WANT PRESERVED !!!

Dir["./lib/**/*.rb"].each {|file| require file }

# utility tu sue when a run gets stuck trying to resume with a workbook already open 
# that can't be gotten control of, possibly because the Excel instance is different
RobustExcelOle::Excel.kill_all
