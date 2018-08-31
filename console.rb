require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'nokogiri'
require 'robust_excel_ole'
require 'pry'

Dir["./lib/**/*.rb"].each {|file| require file }

# launce the experiment main loop
Experiment.new()

# Pry is a debugging tool that will create a breakpoint at any line in the
# code where "binding.pry" is inserted 
pry
