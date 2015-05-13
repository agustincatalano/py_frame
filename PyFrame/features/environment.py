from os import environ
from behaving.web import environment as behaveenv


# esto hay que intentar sacarlo de aca y ponerlo en el otro environment
def before_all(context):
    context.default_browser = environ.get('BROWSER')
    if environ.get('GRID_URL'):
        context.remote_webdriver = True
        context.browser_args = {'url': environ.get('GRID_URL'),
                                'browser': environ.get('BROWSER')}
    behaveenv.before_all(context)


def after_scenario(context, scenario):
    behaveenv.after_scenario(context, scenario)


def after_all(context):
    behaveenv.after_all(context)