module ImportData

  # Need to set up data grabs here. Not worrying about closest paths or least
  # clicks, just getting functionality down.
  def self.get_node_weights
    navigate_to_ranking_system


    weights_tab = $wait.until { $driver.find_element(:id, "rank-syst-func-tab3") }
    weights_tab.click

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

  def self.get_universe_options
    
    navigate_to_settings_tab
    settings_tab = $wait.until { $driver.find_element(:id, "scrtab_7") }
    settings_tab.click

    universes_form = $wait.until { $driver.find_element(:id, "universeUid") }
    options = universes_form.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    universe_options = options.map do |option|
      { value: option.attribute("value"), text:  option.text }
    end

    # Removing the last one because it is just the "add another" option.
    universe_options.pop

    return universe_options
  end

  def self.backtest_results
     
    tab = $wait.until { $driver.find_element(:id, "scrtab_3") }
    tab.click
    
    $wait.until { $driver.find_element(:id, "clearResults") }.click

    backtest_button = $wait.until { $driver.find_element(:id, "runBacktest") }
    backtest_button.click

  end

  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 

  private

  # Hard coded for testing purposes. Will need an ID to run programmatically.
  def self.navigate_to_ranking_system(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/ranking-system/333916")  # make this an ENV variable?
#    $driver.navigate.to(ENV["RANKING_SYSTEM_URL"])  # make this an ENV variable?
  end

  def self.navigate_to_settings_tab(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/213685")  # make this an ENV variable?
#    $driver.navigate.to(ENV["SCREENS_SETTINGS_URL"])  # make this an ENV variable?
  end  

end