require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'rubyXL'
require 'nokogiri'
require 'pry'

Dir["./lib/**/*.rb"].each {|file| require file }

$driver = Selenium::WebDriver.for :chrome
$wait = Selenium::WebDriver::Wait.new(:timeout => 15)


Experiment.new()		# launch the experiment mail loop

# Conceivably add some mop up activities here

# Pry opens up an IRB prompt that runs on the line where Pry is evoked with 'binding.pry'. 
# You can also use this to debug code within the project.

pry
