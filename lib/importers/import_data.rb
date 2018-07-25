module ImportData

  # Need to set up data grabs here. Not worrying about closest paths or least
  # clicks, just getting functionality down.
  def self.get_node_weights

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

    universe_options = $wait.until { $driver.find_element(:id, "universeUid") }
    options = universe_options.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    universe_options = options.map do |option|
      { value: option.attribute("value"), text:  option.text }
    end

    universe_options.pop
    ap universe_options
    return universe_options
  end

  def self.set_universe(universe_id)
    # Assuming you are coming from the appropriate page...
  end  


  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 

  private

  # Hard coded for testing purposes. Will need an ID to run programmatically.
  def navigate_to_ranking_system(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/ranking-system/332295")
  end

  def self.navigate_to_settings_tab(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/212386")
  end  

end