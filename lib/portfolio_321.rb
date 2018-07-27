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

    if args[:log_in]
      log_in
    end

  end

  def log_in
    # Go to P123 login page and, enter login values defined in .env file 
    go_to "https://www.portfolio123.com/login.jsp?url=%2F"

    login_box = @driver.find_element(:id, "LoginUsername")
    login_box.send_keys(@login_info[:username])

    pw_box = @driver.find_element(:id, "LoginPassword")
    pw_box.send_keys(@login_info[:password])

    signin_btn = @driver.find_element(:id, "Login")   #hit button to log in
    signin_btn.click()
    
  end

  def go_to(url)
    @driver.navigate.to(url)
  end

  def pull_and_insert_weights

    node_weights = ImportData.get_node_weights
    ap node_weights
    # example node_weights value: {:name=>"LTGrthMean", :input_id=>"node-weight-1", :input_value=>"0"}

    # Convert node_weights to an array of integers
    clm = []
    clm = node_weights.map { |x| x[:input_value].to_i} 
  

    path = "C:/Users/Scott/GitHub Jobs/Test Excel Write.xlsx"     # Any value to using global(s) for paths[run config, run data, etc]
    #workbook = XLSXParser.parse("C:/Users/Scott/GitHub Jobs/Test Excel Write.xlsx")
    workbook = XLSXParser.open(path)     # Test here for valid workbook and continue accordingly, errors out if workbook already open
    worksheet = workbook["Nodes"]

    XLSXParser.fill_clm_values(0, 3, worksheet, clm)
    #XLSXParser.write_workbook(workbook, "C:/Users/Scott/GitHub Jobs/Test Excel Write.xlsx")
    XLSXParser.write_workbook(workbook, path)

     clm = node_weights.map { |x| x[:name]} 
     XLSXParser.fill_clm_values(0, 0, worksheet, clm)
     XLSXParser.write(workbook, path)

    # Do some other processing...
    value = 5

    node_weights.each do |node_data|
      input_element = @driver.find_element(:id, node_data[:input_id])

      # Setting the value to 0
      @driver.execute_script("return document.getElementById('#{node_data[:input_id]}').value = '';")

      # Setting the value in the corresponding input box.
      input_element.send_keys(value)
    end

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

  def default_driver
    $driver
  end

end
