class Experiment

   def initialize(args = {})

   # create all paths, file names, data workbooks
   Spreadsheet.new()

   # prep_experiment
   # run_experiment
   # end_experiment

    end
=begin    

	def run_experiment
		# This method contains anything dealing with P123
		# continue = prep_experiment()
		# if continue
		if_runs_todo()			
			incr_run_todo()
			# get_next_todo()			# spreadsheet returns weights and universe for next run
			# send Todo weights & universes to P123
			# Start Backtest
			# scrape results
			# log_run_results(results)	# send results to spreadsheet for logging
			# ap run_completed()		# spreadsheet returns "Run X completed." message
		# end	
		# end_experiment
	end

	def end_experiment
		# final mop up
		# LOG OUT
		# clear resume_run

	end

	def prep_experiment
		if runs_todo()						# check no bogus condition: runs_todo > 0 && > current_run
		# login to P123
		# if experiment_first_run()			# if no todo file already setup and prep files for experiment
			# get weights and universes
				# set up experiment Todo file
			# create Analysis file from template
		# return successful prep
		# else return failed prep


	end
=end

end
