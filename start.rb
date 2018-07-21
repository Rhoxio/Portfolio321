require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'rubyXL'
require 'nokogiri'

# This is loading all of the Ruby files contained in the lib folder.
# That way, we have access to all of the classes and modules as soon as the app boots.
# Double splat is all directories, single splat is all files.
Dir["./lib/**/*.rb"].each {|file| require file }

# Global variables for the driver and wait to keep things consistent.
$driver = Selenium::WebDriver.for :chrome
$wait = Selenium::WebDriver::Wait.new(:timeout => 15)

app = Portfolio321.new({ log_in: true })
p app.driver
app.pull_and_insert_weights
