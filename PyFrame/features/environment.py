from os import environ
from behaving.web import environment as behaveenv
from configobj import ConfigObj
import os
from reports.report_excel import test_report

CONFIG = ConfigObj(os.path.join(os.getcwd(), "..", "config", "config.cfg"))


def before_all(context):
    context.default_browser = environ.get('BROWSER')
    if environ.get('GRID_URL'):
        context.remote_webdriver = True
        context.browser_args = {'url': environ.get('GRID_URL'),
                                'browser': environ.get('BROWSER')}
    context.screenshots_dir = os.path.join(os.getcwd(), CONFIG['screenshot']['dir'])
    behaveenv.before_all(context)


def after_scenario(context, scenario):
    behaveenv.after_scenario(context, scenario)


def after_all(context):
    # try:
        behaveenv.after_all(context)
        test_report.generate_execution_info(context)
