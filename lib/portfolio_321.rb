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
#   Only two pages beside Login are accessed and they are jumped between directly, avoiding
#   intervening screen and component selection pages that must be traversed manually. This is
#   accomplished by manually setting up an experiment run and capturing the URLs which include
#   the specific P123 codes that identify the components selected for the experiment.
###############################################################################################

  # attr_reader makes it so you can call 'driver', 'exporter', or 'xlsx_parser' to read 
  # the variable or object set in initialize, but can't redefine it in this class itself.
  # Being able to set an instance variable requires you to use 'attr_accessible' instead of "attr_reader".

  attr_reader :driver, :exporter, :xlsx_parser

 
  def initialize( p123_urls, args = {} )
  # open the web driver, log in to P123, pull node weight and universe records from site

    # this block is here just in case a need arises to pass in a different driver
    # default_driver is the failover method that will always use the default driver
    open_web_driver()
    @driver = args[:driver] ||= default_driver
    @login_info = args[:login_info] ||= default_login_info

    if args[:log_in]  # log in to P123 if directed to do so
      log_in()
    end

    # save off the Urls for required P123 pages
    @p123_rank_system_url = p123_urls[0]
    @p123_screens_url = p123_urls[1]

    @node_weights = pull_node_weights()
    @universes = pull_universe_options()
  end

  def log_in
  # log in to P123 website

    go_to(ENV["LOGIN_URL"])   # go to login page

    # enter the P123 account user id and password
    login_box = @driver.find_element(:id, "LoginUsername")
    login_box.send_keys(@login_info[:username])

    pw_box = @driver.find_element(:id, "LoginPassword")
    pw_box.send_keys(@login_info[:password])

    # click the 'Login' button
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
  # experiment API: returns an array of rank node names pulled from P123

      node_names = Array.new
      @node_weights.each_with_index { |node, idx| node_names[idx] = node[:name]}  
      return node_names
  end

  def pull_node_weights
  # get node information from P123     
    
    goto_node_weights_tab

    # find the node weights table and extract the desired portion of it
    table = $wait.until {
      element = @driver.find_element(:id, "weights-cont-table")
    }    
    td = table.find_elements(:xpath, "./tbody/tr/td")

    # assign a new array with the desired element of each entry in the other array 
    paired_td = td.each_slice(2).to_a    

    # create an array of node weight hashes of the node name, storage location id, and weight value
    data = paired_td.map do |td|
      input = td[1].find_elements(:xpath, "./input")[0]

      { name: td[0].text, input_id: input.attribute('id'), input_value: input.attribute('value') }
    end

    # remove the header (index 0) of the table as it isn't data that get used
    data.shift

    return data
  end

  def push_node_weights(node_weights)  
  # experiment API: send the node weights for this run to P123
  # node_weights parameter is an array of :value

    goto_node_weights_tab
    # value must NOT be set to 0 or all weights are filled in with an extra 0, creating 10x values
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
  # experiment API: returns an array of screener universe names extracted from the custom universes pulled from P123

      universe_names = Array.new(@universes.length, "")
      @universes.each_index { |idx| universe_names[idx] = @universes[idx][:text] }  
      return universe_names
  end

  def pull_universe_options
  # experiment API: get the available custom universes from P123: custom group is at (:xpath, "./optgroup")[-1])

    goto_screens_settings_tab()
    universes_form = $wait.until { @driver.find_element(:id, "universeUid") }
    options = universes_form.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    universe_options = options.map do |option|
      { value: option.attribute("value"), text:  option.text }
    end

    # Removing the last item because it is just the "add another" option
    universe_options.pop  
    return universe_options  
  end

  def push_universe(universe_name)
  # exoeriment API: send the Todo universe for this run to P123

    goto_screens_settings_tab()

    universe_options = $wait.until { @driver.find_element(:id, "universeUid") }
    options = universe_options.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    selected_option = options.find do |o|
      o.attribute('text') == universe_name
    end
    selected_option.click
  end

  def goto_node_weights_tab()
  # navigate to P123 rank systmens page then to the weights tab

    go_to(@p123_rank_system_url)

    weights_tab = $wait.until { @driver.find_element(:id, "rank-syst-func-tab3") }
    weights_tab.click    
  end

  def goto_screens_settings_tab()
  ## navigate to P123 Screens page then to Settings tab

    go_to(@p123_screens_url)

    # dispose of the dialog box that arises when jumping from some pages to this particular page
    sate_navigation_alert()  

    settings_tab = $wait.until { @driver.find_element(:id, "scrtab_7") }
    settings_tab.click
  end

  def goto_run_backtest_tab()
  # navigate to P123 Screens page then to Backtest tab

    go_to(@p123_screens_url) 
    tab = $wait.until { @driver.find_element(:id, "scrtab_3") }
    tab.click
  end

  def execute_backtest(results_report)
  # run the backtest and download the results: assumes already on the Screens menu but not Backtest tab

    goto_run_backtest_tab()

    # all :id options - clearResults, runScreen, rerunScreen, runBacktest, reRunBacktest, runRBacktest, rerunRBacktest
    run_button = $wait.until { @driver.find_element(:id, "runBacktest") }
    run_button.click

    # wait up to 15s for run completion to post results table: Selenium::..implicit_wait(secs) if need > 15s wait 
    dl_button = $wait.until { @driver.find_element(:id, "results-table") }

    # select which button to click for downloading the specified chart or table run results
    # both buttons remain 'displayed' and 'enabled' even when outside of the page's viewport
    if results_report == "chart" then
      dl_button = $wait.until { @driver.find_element(:xpath, "//*[@id='scr-result']/div[2]/a") }   # note '' inside of ""
    else
      dl_button = $wait.until { @driver.find_element(:xpath, "//*[@id='results-table']/table/thead/tr[1]/th/div/div/a") }
    end

    # click to download chosen report, causing an interim file to be downloaded
    # the Backtest page then waits for a 'save' or 'cancel' selection to be made
    # see Spreadsheet for a description of how this interim file is captured during the wait
    dl_button.click
  end

  def terminate_backtest()
  # click 'Clear Backtest Results' to stop pending on download completion while readying Backtest page for next run

    button = $wait.until { @driver.find_element(:id, "clearResults") }
    button.click

    # leave page to cause download process to end; handle dialog box raised by this particular page change
    go_to(@p123_rank_system_url)
    sate_navigation_alert() 
    return
  end

  def end_experiment()
    @driver.quit
  end

  def sate_navigation_alert()
  # accept (click 'Leave') on the Leave/Cancel dialog box that warns of possible unsaved changes

    @driver.switch_to.alert.accept rescue Selenium::WebDriver::Error::NoAlertOpenError
  end


  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 

  private

  def default_login_info()
    {username: ENV["LOGIN_USERNAME"], password: ENV["LOGIN_PASSWORD"]}
  end

  def open_web_driver()
  # select driver for Chrome and a 15s max wait time for P123 responses

    $driver = Selenium::WebDriver.for :chrome
    $wait = Selenium::WebDriver::Wait.new(:timeout => 15)
  end

  def go_to(url)
  
    # check not already on a page before going there: 
    # @driver urps trying to go to the current page
    if url != @driver.current_url then @driver.navigate.to(url) end
  end

  def default_driver
    $driver
  end

end
