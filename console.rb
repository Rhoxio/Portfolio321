require 'dotenv/load'
require "selenium-webdriver"
require "awesome_print"
require 'rubyXL'
require 'nokogiri'
require 'pry'

# This is loading all of the Ruby files contained in the lib folder.
# That way, we have access to all of the classes and modules as soon as the app boots.
# Double splat is all directories, single splat is all files.
Dir["./lib/**/*.rb"].each {|file| require file }

# This opens up an IRB prompt that run on the same line as 'pry' is evoked. 
# You can also use this to debug code within the project.
pry