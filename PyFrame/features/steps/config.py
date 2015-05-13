from behave import step
from selenium import webdriver
from behaving.web.steps import browser
import behaving.web
#contex es un objeto que crea behave que se comparte entre los distintos steps (or ejemplo la config, variables etc)


@step(u'open the browser')
def step_impl(context):
    context.execute_steps(u'''
             Given a browser
        ''')

