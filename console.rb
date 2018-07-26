require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'rubyXL'
require 'nokogiri'
require 'pry'

Dir["./lib/**/*.rb"].each {|file| require file }

$driver = Selenium::WebDriver.for :chrome
$wait = Selenium::WebDriver::Wait.new(:timeout => 15)

# This opens up an IRB prompt that run on the same line as 'pry' is evoked. 
# You can also use this to debug code within the project.
pry