class Portfolio321

  # This makes it so you can call 'driver', 'exporter', or 'xlsx_parser' to read 
  # the variable or object set in initialize, but can't redefine it in this class itself.
  # Being able to set an instance variable requires you to use 'attr_accessible' instead of "attr_reader".

  attr_reader :driver, :exporter, :xlsx_parser
 
  def initialize(args = {})
    # Initialize is run when you create a new instance of a class.
    # See start.rb if you want an example. 

    # This is here just in case you need to pass in a different driver.
    # default_driver is the failover method that will always use the default driver.
    # @ is an instance variable and can be evoked only in this instance of this class.

    @driver = args[:driver] ||= default_driver
    @login_info = args[:login_info] ||= default_login_info

    @exporter = ExportData.new()
    @xlsx_parser = XLSXParser.new()

    if args[:log_in]
      log_in  
    end

  end

  def log_in

    go_to "https://www.portfolio123.com/login.jsp?url=%2F"

    login_box = @driver.find_element(:id, "LoginUsername")
    login_box.send_keys(@login_info[:username])

    pw_box = @driver.find_element(:id, "LoginPassword")
    pw_box.send_keys(@login_info[:password])

    signin_btn = @driver.find_element(:id, "Login")
    signin_btn.click()
    
  end

  def go_to(url)
    @driver.navigate.to(url)
  end

  def test_action

    go_to("https://www.portfolio123.com/app/opener/PTF")

    x = 1
    while x > 0
      performance_button = $wait.until { @driver.find_element(:link_text, "Performance") }
      performance_button.click

      general_button = $wait.until { @driver.find_element(:link_text, "General") }
      general_button.click

      statistics_button = $wait.until { @driver.find_element(:link_text, "Statistics") }
      statistics_button.click
      
      puts @driver.title
    end
  end

  def action_delegator(*args)
    # To be used to delegate which action is to be taken. Base control flow for
    # triggering other code. 
  end

  # # # # # # # # # # # # #  
  # START OF PRIVATE METHODS
  # # # # # # # # # # # # # 

  private

  # We don't want or need these two ENV variables to be available anywhere else,
  # so we use a Prvate method to make it so only this specific instance of this class
  # can access it. Keeps the public interface clean.

  def default_login_info
    {username: ENV["USERNAME"], password: ENV["PASSWORD"]}
  end

  def default_driver
    $driver
  end

end
