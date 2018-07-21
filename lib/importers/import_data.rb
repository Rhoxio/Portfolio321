module ImportData

  # Need to set up data grabs here. Not worrying about closest paths or least
  # clicks, just getting functionality down.

  def self.get_node_weight_labels

    data = []

    # Hard coded for testing purposes. Will need an ID to run programmatically.
    $driver.navigate.to("https://www.portfolio123.com/app/ranking-system/332295")

    weights_tab = $wait.until { $driver.find_element(:id, "rank-syst-func-tab3") }
    weights_tab.click

    table = $wait.until {
      element = $driver.find_element(:id, "weights-cont-table")
      element if element.displayed?
    }    
    
    elements = table.find_elements(:xpath, "./tbody/tr/td")
    paired_elements = elements.each_slice(2).to_a

    paired_elements.each do |td|
      data << { name: td[0].text, input_id: td[1].find_elements(:xpath, "./input")[0].attribute('id')}
    end

    # Removing the header (index 0) of the table as it isn't pure data.
    data.shift

    return data
    
  end

end