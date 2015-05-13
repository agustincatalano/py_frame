from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.wait import WebDriverWait


class BasePage():

    def __init__(self, driver):
        self.driver = driver

    def find_element(self, locator, **kwargs):
        wait = kwargs.get('wait', 2)
        retries = kwargs.get('retries', 5)
        wd_wait = WebDriverWait(self.driver, wait)
        acum = 0
        element = None
        while acum <= retries and not element:
            if wd_wait.until(ec.visibility_of_element_located(locator)):
                element = self.driver.find_element(locator)
            acum += 1
        return element




