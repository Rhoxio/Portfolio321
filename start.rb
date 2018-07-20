require 'dotenv/load'

# This is loading all of the Ruby files contained in the lib folder.
# That way, we have access to all of the classes and modules as soon as the app boots.
# Double splat is all directories, single splat is all files.
Dir["./lib/**/*.rb"].each {|file| require file }

# ExportData.test_linkage
# Portfolio321.log_in
ap XLSXParser.parse('./files/xlsx/test.xlsx')[0][0][0]