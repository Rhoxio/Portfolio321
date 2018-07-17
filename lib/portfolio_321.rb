# require "portfolio_321/version"
require "selenium-webdriver"

module Portfolio321
  # Your code goes here...

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

Portfolio321.log_in
