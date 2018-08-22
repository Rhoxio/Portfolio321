module UniverseTraverser

  def self.set_universe(universe_id)

    # Hard coding this for demo.
    # $driver.navigate.to("https://www.portfolio123.com/app/screen/summary/212386")

    # Assuming you are coming from the appropriate page...
    universe_options = $wait.until { $driver.find_element(:id, "universeUid") }
    options = universe_options.find_elements(:xpath, "./optgroup")[-1].find_elements(:xpath, "./option")

    selected_option = options.find do |o|
      o.attribute('value') == universe_id
    end

    selected_option
    selected_option.click

    run_button = $wait.until { $driver.find_element(:id, "runScreen") }
    run_button.click

  end  

end