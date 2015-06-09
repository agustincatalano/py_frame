# -*- coding: utf-8 -*-

from behave import step
import os
from logger import logger_factory
from os import path
from behaving.web.steps import browser
# contex es un objeto que crea behave que se comparte entre los distintos steps (or ejemplo la config, variables etc)


@step(u'open the browser')
def step_impl(context):
    context.execute_steps(u'''
             Given a browser
        ''')


@step(u'Take a screentshot')
def take_screenshot(context):
    path = context.screenshots_dir
    if not os.path.exists(path):
        os.makedirs(path)
    logger_factory.get_logger().info('Taking a screenshot....')
    context.execute_steps(u'''
        Given I take a screenshot
        ''')

