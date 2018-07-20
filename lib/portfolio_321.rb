require "selenium-webdriver"

# Dir["./**/*.rb"].each {|file| require file }

# ExportData.test_linkage

module Portfolio321
  # Your code goes here...

  def initialize
    # Need to set up ENV to handle logging in for us automatically.
    # Need to set up global set for driver that will need to log in.
  end

  def self.log_in

    driver = Selenium::WebDriver.for :chrome
    driver.navigate.to "https://www.portfolio123.com/login.jsp?url=%2F"
    wait = Selenium::WebDriver::Wait.new(:timeout => 15)

    loginBox = driver.find_element(:id, "LoginUsername")
    loginBox.send_keys("program_roller")
    pwBox = driver.find_element(:id, "LoginPassword")
    pwBox.send_keys("x2246rwq")
    signinBtn = driver.find_element(:id, "Login")
    signinBtn.click()

    driver.navigate.to("https://www.portfolio123.com/app/opener/PTF")

    x = 1
    while x > 0
      performance_button = wait.until { driver.find_element(:link_text, "Performance") }
      performance_button.click

      general_button = wait.until { driver.find_element(:link_text, "General") }
      general_button.click

      statistics_button = wait.until { driver.find_element(:link_text, "Statistics") }
      statistics_button.click
      
      puts driver.title
    end
    

    
  end

end
