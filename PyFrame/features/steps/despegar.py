from behave import step
from logger import logger_factory
from ui.page_object_despegar.home_page import HomePage


@step(u'I go to despegar home page')
def go_to_home_page(context):
    home_page = HomePage(context.browser.driver)
    home_page.get(home_page.URL)
    home_page.validate_title()
    context.current_page = home_page


@step(u'I close the advertisement')
def close_ad(context):
    logger_factory.get_logger().info('Closing advertisement')
    context.current_page.close_ad()


@step(u'I enter the place "{place}"')
def enter_place(context, place):
    context.current_page.enter_place(place)
