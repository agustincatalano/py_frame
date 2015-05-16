# -*- coding: utf-8 -*-
from selenium.webdriver.common.by import By

from ui.base_page import BasePage


class HomePage(BasePage):

    URL = u'http://www.despegar.com.ar/'
    TITLE = u'Vuelos, hoteles, paquetes y mucho más! | Despegar.com Argentina'
    close_ad_loc = (By.CSS_SELECTOR, '.nibbler-common-overlay-close')

    def close_ad(self):
        self.find_clickable_element(self.close_ad_loc).click()
