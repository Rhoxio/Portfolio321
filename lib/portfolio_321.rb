class Portfolio321
###############################################################################################
#   This class manages all communications with the P123 webite, which includes:
#   - logging in
#   - pulling the list of available node and universe choices from P123
#     - passing out node names and universe names to Experiment
#   - pushing node weight and universe selection to P123
#   - launching runs (P123 Run Backtest button)
#   - initiating and terminating run results collection from P123: both chart and table data
#   - clearing run results after collection
#   - logging off
#
#   Some legacy code is saved for reference in a comments block at the end of the file
###############################################################################################

  # attr_reader makes it so you can call 'driver', 'exporter', or 'xlsx_parser' to read 
  # the variable or object set in initialize, but can't redefine it in this class itself.
  # Being able to set an instance variable requires you to use 'attr_accessible' instead of "attr_reader".

  attr_reader :driver, :exporter, :xlsx_parser

 
  def initialize( p123_urls, args = {} )
    # Initialize is run when you create a new instance of a class.
    # See start.rb if you want an example. 

    # This block is here just in case you need to pass in a different driver.
    # default_driver is the failover method that will always use the default driver.
    # @ is an instance variable and can be evoked only in this instance of this class.

    @driver = args[:driver] ||= default_driver
    @login_info = args[:login_info] ||= default_login_info

    if args[:log_in]
      log_in()
    end

    @p123_rank_system_url = p123_urls[0]
    @p123_screens_url = p123_urls[1]

    @node_weights = pull_node_weights()
    @universes = pull_universe_options()
  end

  def log_in

    go_to(ENV["LOGIN_URL"])
   
    login_box = @driver.find_element(:id, "LoginUsername")
    login_box.send_keys(@login_info[:username])

    pw_box = @driver.find_element(:id, "LoginPassword")
    pw_box.send_keys(@login_info[:password])

    signin_btn = @driver.find_element(:id, "Login")
    signin_btn.click()
  end

  def set_p123_urls ( p123_urls)
  # set p123 page urls to those read in from Setup 

    @p123_rank_system_url = p123_urls[0]
    @p123_screens_url = p123_urls[1]
    return
  end

  def get_node_names
  # returns an array of rank node names pulled from P123

      node_names = Array.new
      @node_weights.each_with_index { |node, idx| node_names[idx] = node[:name]}  
      return node_names
  end

  def pull_node_weights
  # get node information from P123     
    
    go_to_node_weights_tab

    table = $wait.until {
      element = $driver.find_element(:id, "weights-cont-table")
    }    
    
    td = table.find_elements(:xpath, "./tbody/tr/td")
    paired_td = td.each_slice(2).to_a

    data = paired_td.map do |td|
      input = td[1].find_elements(:xpath, "./input")[0]

      { name: td[0].text, input_id: input.attribute('id'), input_value: input.attribute('value') }
    end

    # Removing the header (index 0) of the table as it isn't data we need to act upon.
    data.shift

    return data
  end

  def push_node_weights(node_weights)  # node_weights is an array of :value
  # send the Todo node weights for this run to P123

    go_to_node_weights_tab

    # value must NOT be set to 0 or all weights are filled in with an extra 0, or 10x values
    value = nil   
    idx = 0
    @node_weights.each do |node_data|
      input_element = $wait.until {@driver.find_element(:id, node_data[:input_id]) }
      @driver.execute_script("return document.getElementById('#{node_data[:input_id]}').value = '#{node_weights[idx]}';")
      input_element.send_keys(value)
      idx += 1
    end

    # new weights pushed, now click 'update' button to have P123 save changes
    @driver.execute_script("scroll(250, 0)")    # jump to top of page to get update button into browser's viewport
    update_button = $wait.until { @driver.find_elements(:tag_name, "input").find { |i| i.attribute("value") == "Update" } }
    update_button.click
  end

  def get_universe_names
  # returns an array of screener universe names pulled from P123

      universe_names = Array.new(@universes.length, "")
      @universes.each_index { |idx| universe_names[idx] = @universes[idx][:text] }  
      return universe_names
  end

  def pull_universe_options
  # get the available custom universes from P123

    go_to_screens_settings_tab

    universes_form = $wait.until { $driver.find_element(:id, "universeUid") }
    options = universes_form.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    universe_options = options.map do |option|
      { value: option.attribute("value"), text:  option.text }
    end

    # Removing the last item because it is just the "add another" option
    universe_options.pop  
    return universe_options  
  end

  def push_universe(universe_name)
  # send the Todo universe for this run to P123

    go_to_screens_settings_tab

    universe_options = $wait.until { $driver.find_element(:id, "universeUid") }
    options = universe_options.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    selected_option = options.find do |o|
      o.attribute('text') == universe_name
    end

    selected_option.click
  end
  def go_to_node_weights_tab
  # navigate to P123 rank systmens page then to the weights tab

    go_to(@p123_rank_system_url)

    weights_tab = $wait.until { $driver.find_element(:id, "rank-syst-func-tab3") }
    weights_tab.click    
  end

  def go_to_screens_settings_tab
  ## navigate to P123 Screens page then to Settings tab

    go_to(@p123_screens_url)
    settings_tab = $wait.until { $driver.find_element(:id, "scrtab_7") }
    settings_tab.click
  end

  def go_to_run_backtest_tab
  # navigate to P123 Screens page then to Backtest tab

    go_to(@p123_screens_url)     # apparently Selenium doesn't like going to the page it's already on
    tab = $wait.until { $driver.find_element(:id, "scrtab_3") }
    tab.click
  end

  def execute_backtest
  # run the backtest and download the results: assumes already on the Screens menu but not Backtest tab

    go_to_run_backtest_tab

    # all :id's - clearResults, runScreen, rerunScreen, runBacktest, reRunBacktest, runRBacktest, rerunRBacktest
    run_button = $wait.until { $driver.find_element(:id, "runBacktest") }
    run_button.click

    # need to select a wait mechanism to hold off at this point while the run is executing
    # 'sleep' will work but finding a "run completed" test condition to work with it is preferable to
    # waiting an arbitrarily long time to assure any run will have completed under any conditions
    # Also need: 
    # => button id's (missing/hidden) or other means to trigger both chart and table results file downloads
    # => some means to respond to save/cancel download choices so as to continue to next run - these are not website actions that Selenium can drive
    # => the save results process depends on the results capture process
  end

  def sate_navigation_alert
    @driver.browser.switch_to.alert.accept rescue Selenium::WebDriver::Error::NoAlertOpenError
  end


  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 
  private

  # We don't want or need these two ENV variables to be available anywhere else,
  # so we use a Prvate method to make it so only this specific instance of this class
  # can access it. Keeps the public interface clean.

  def default_login_info
    {username: ENV["LOGIN_USERNAME"], password: ENV["LOGIN_PASSWORD"]}
  end

  def go_to(url)
    # if not already on a page, go there: @driver urps trying to go to the current page
    if url != @driver.current_url then @driver.navigate.to(url) end
  end

  def default_driver
    $driver
  end

=begin  legacy code saved for reference

  # Hard coded for testing purposes. Will need an ID to run programmatically.
  def navigate_to_ranking_system(input = nil)
#    $driver.navigate.to("https://www.portfolio123.com/app/ranking-system/333916")  # make this an ENV variable?
    $driver.navigate.to(ENV["RANKING_SYSTEM_URL"])  # make this an ENV variable?
  end

  def navigate_to_settings_tab(input = nil)
#    $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/213685")  # make this an ENV variable?
    $driver.navigate.to(ENV["SCREENS_SETTINGS_URL"])  # make this an ENV variable?
  end  
  def pull_and_insert_weights

    @node_weights = pull_node_weights

    # Do some other processing...
    value = 0

    @node_weights.each do |node_data|
      input_element = @driver.find_element(:id, node_data[:input_id])

      # Setting the value to 0
      @driver.execute_script("return document.getElementById('#{node_data[:input_id]}').value = '';")

      # Setting the value in the corresponding input box.
      input_element.send_keys(value)
    end

    # Jump to top of page so the browser has the button in it's viewport... 
    @driver.execute_script("scroll(250, 0)")
    
    update_button = $wait.until { @driver.find_elements(:tag_name, "input").find { |i| i.attribute("value") == "Update" } }
    update_button.click
    
  end
=end
end
