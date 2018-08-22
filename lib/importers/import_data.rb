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
    
    navigate_to_settings_tab    # navigate to P123 Screens page where the Settings tab is located
    settings_tab = $wait.until { $driver.find_element(:id, "scrtab_7") }  # now go to tab
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

    results = []
     
    tab = $wait.until { $driver.find_element(:id, "scrtab_3") }
    tab.click
    
    $wait.until { $driver.find_element(:id, "clearResults") }

    sleep 300
    
    backtest_button = $wait.until { $driver.find_element(:id, "runBacktest") }
    backtest_button.click

    # ap check_for_ratelimit    

    results_table = $wait.until { $driver.find_element(:id, "results-table") }
    base_keys = results_table.find_elements(:xpath, "./table/thead/tr")[1].find_elements(:xpath, "./th").each_with_index.map {|el, i| { el.text.gsub("\n", "") => nil } }

    base_keys.shift

    results_table.find_elements(:xpath, "./table/tbody/tr").each do |tr|
      if tr["class"].split(" ").include?("rowAlt1")
        td = tr.find_elements(:xpath, "./td").each_with_index.map {|t, i| t.text }
        td.shift
        ap td
      end
    end
    
    ap base_keys

    sleep 20000

  end

  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 

  private

  # def self.check_for_ratelimit
  #   begin
  #     ap wait = Selenium::WebDriver::Wait.new(:timeout => 15)

  #     ap error_container = wait.until { $driver.find_element(:id, "scr-error") } 
  #     ap minutes_left = error_container.text
  #   rescue
  #     false
  #   end
  # end  

  # Hard coded for testing purposes. Will need an ID to run programmatically.
  def self.navigate_to_ranking_system(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/ranking-system/333916")
  end

  def self.navigate_to_settings_tab(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/213685")
  end

  def self.navigate_to_settings_tab(input = nil)
    $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/213685")  # make this an ENV variable?
  end  

end