# -*- coding: utf-8 -*-
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait


class BasePage():
    TITLE = u''

    def __init__(self, driver):
        self.driver = driver

    def find_element(self, locator, condition=ec.visibility_of_element_located, **kwargs):
        wait = kwargs.get('wait', 2)
        retries = kwargs.get('retries', 5)
        wd_wait = WebDriverWait(self.driver, wait)
        acum = 0
        element = None
        while acum <= retries and not element:
            try:
                if wd_wait.until(condition(locator)):
                    element = self.driver.find_element(*locator)
                    acum += 1
            except TimeoutException:
                acum += 1
        return element

    def find_visible_element(self, locator, **kwargs):
        return self.find_element(locator, **kwargs)

    def find_clickable_element(self, locator, **kwargs):
        return self.find_element(locator, condition=ec.element_to_be_clickable, **kwargs)

    def select_drop_down(self, drop_down_locator, value):
        element = self.find_clickable_element(drop_down_locator)
        if element:
            my_select = Select(element)
            my_select.select_by_value(value)
        else:
            raise Exception('Unable to find drop down with locator %s' % str(drop_down_locator))

    def validate_title(self):
        assert self.driver.title == self.TITLE, 'Title does not match. Expected: %s. Obtained: %s' % (self.TITLE,
                                                                                                      self.driver.title)

    def get(self, url):
        self.driver.get(url)