class Experiment
####################################################################################################### 	
# 	This class serves as the main loop.  It coordinates the sequencing of actions whose details are
# 	managed by Experiment (P123 interactions) and Spreadsheet (configuration and data file management).
# 	This application consists of these three classes with minor associated startup code, ENV data, etc.
# 	The architecture for gathering and saving run results data is not yet defined and could possibly 
# 	involve an atypical sidestepping of Experiment's mediation between Portfolio321 and Spreadsheet.
####################################################################################################### 	

   	def initialize(args = {})
	# instantiate other main classes, prep to execute experiment runs, launch runs, and clean up when done

		# create all paths, file names, data workbooks
		@sheet = Spreadsheet.new()

		# login to P123 so rank & universe data for run can be uploaded
		@web_app = Portfolio321.new( @sheet.get_p123_urls, { log_in: true } )
		prep_success = prep_experiment()

		if prep_success then run_experiment end
		 
		end_experiment
	end

	def prep_experiment
		if @sheet.if_runs_todo() then							# check for obvious bogus conditions from bad Setup entries
	  		@node_names = @web_app.get_node_names 		 		# get array of node names for spreadsheet use
	  		@node_weights = Array.new(@node_names.length, 0)	# instantiate for later use
 			@universe_names = @web_app.get_universe_names		# get array of universe names for spreadsheet use
		  	@sheet.setup_todo(@node_names, @universe_names)		# 
		return true 			# successful prep
		else return false 		#f ailed prep
		end
	end

	def run_experiment
	# This method contains anything dealing with P123

		if @sheet.if_runs_todo() then
			@sheet.incr_run_todo()		# set up files and counters for new run

			# get Todo data for next run and push it to P123
			@universe_name = @sheet.get_next_run_todo(@node_weights)			
			@web_app.push_node_weights(@node_weights)		# push array of node weights
			@web_app.push_universe(@universe_name)			# push selected universe name

			@web_app.execute_backtest()		# run test and hit desired results download button
			# @sheet.log_run_results 		# copy and save results from temp file downloaded, do run done housekeeping
			# @web_app.close_run			# cancel results download, clear backtest results 
			# run_completed()		# post "Run X completed." message? with timestamp? part of run done housekeeping?
		end	
		# end_experiment
	end

	def end_experiment		# consider what this does if expecting to RESUME
		# final mop up
		# LOG OUT
		# clear resume_run 
	end

end
