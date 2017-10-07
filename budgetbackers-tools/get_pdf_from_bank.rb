require 'selenium-webdriver'

driver = Selenium::WebDriver.for :chrome
driver.navigate.to 'https://www.discountbank.co.il'

login_button  = driver.find_element(:id, 'hpc-login-toggle')
login_button.click

driver.switch_to.frame('hpc-iframe')

forms = driver.find_elements(class: 'inputLabel')

tokens = ['066243569', '7bpRjpLNYu2aj9', '9ff2etb']

forms = forms.zip(tokens)
forms.each { |x, y| x.send_keys(y) }

driver.find_element(class: 'thebutton').click

wait = Selenium::WebDriver::Wait.new(:timeout => 10)

wait.until { driver.find_element(id: 'OSH_MAIN_WORLD-link') }

driver.find_element(id: 'OSH_MAIN_WORLD-link').click

wait.until { driver.find_element(id: 'activities_view') }

driver.find_element(id: 'activities_view').find_element(class: 'advanced-search').find_element(name: 'filter').click


driver.quit