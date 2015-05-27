# -*- coding: utf-8 -*-
from selenium.webdriver.common.by import By

from ui.base_page import BasePage


class HomePage(BasePage):

    URL = u'http://www.despegar.com.ar/'
    TITLE = u'Vuelos, hoteles, paquetes y mucho más! | Despegar.com Argentina'
    close_ad_loc = (By.CSS_SELECTOR, '.nibbler-common-overlay-close')
    country_text_box = (By.CSS_SELECTOR, '.sbox-destination')

    def close_ad(self):
        self.find_clickable_element(self.close_ad_loc).click()

    def enter_place(self, place):
        self.find_visible_element(self.country_text_box).send_keys(place)