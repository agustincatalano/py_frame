import argparse
import logging.config
from os import environ
import os
import traceback
from behave import __main__ as behave_script
import sys
from configobj import ConfigObj
import json
from reports.report_excel import generate_report_xlsx


ARGS = None
CONFIG = ConfigObj(os.path.join(os.getcwd(), "..", "config", "config.cfg"))
REPORTS = '..\\features'
INFO_FILE = 'results.json'


def main():
    __parse_arguments()
    __set_environment_variables()
    __set_proxy_configuration()
    __create_paths()
    __set_variables_to_behave()
    __run_test_cases()
    __generate_execution_reports()


def __parse_arguments():
    parser = argparse.ArgumentParser(
        description='Tesis')
    parser.add_argument('-t',
                        '--tags',
                        action="append",
                        help='tags para ejecutar los test con Behave',
                        required=False)
    parser.add_argument('-b',
                        '--browser',
                        default="firefox",
                        help='firefox, chrome. Default: firefox',
                        required=False)
    parser.add_argument('-g',
                        '--grid',
                        help='The remote grid url in which the tests will be executed',
                        required=False)
    parser.add_argument('-o',
                        '--output_folder',
                        default=CONFIG['results']['output_dir'],
                        help='output folder where the test results will be stored under project folder',
                        required=False)
    parser.add_argument('--http_proxy',
                        action="store_true",
                        default=False,
                        help='setup default http proxy (e.g. http://<proxy_url>:(proxy_port>)',
                        required=False)
    parser.add_argument('--https_proxy',
                        action="store_true",
                        default=False,
                        help='setup default https proxy (e.g. https://<proxy_url>:(proxy_port>)',
                        required=False)
    parser.add_argument('--no_proxy',
                        action="store_true",
                        default=False,
                        help='setup endpoints to be ignored by the proxy (e.g. "127.0.0.1,123.1.1.1")',
                        required=False)
    global ARGS
    ARGS = parser.parse_args()


def __set_environment_variables():
    __set_environment_variable('BROWSER', ARGS.browser)
    __set_environment_variable('GRID_URL', ARGS.grid)
    __set_tags_as_environment_variable(ARGS.tags)
    __set_environment_variable('OUTPUT', os.path.abspath(ARGS.output_folder))


def __set_proxy_configuration():
    __set_environment_variable('http_proxy', CONFIG['proxy_solution']['http_proxy'])
    __set_environment_variable('https_proxy', CONFIG['proxy_solution']['https_proxy'])
    __set_environment_variable('no_proxy', CONFIG['proxy_solution']['no_proxy'])


def __set_environment_variable(variable, value):
    if value:
        environ[variable] = str(value)
        print "[{}]: {}".format(variable, value)


def __set_tags_as_environment_variable(tags):
    if tags:
        for tag in tags:
            if environ.get('TAGS'):
                __set_environment_variable('TAGS', environ.get('TAGS') + ";" + tag)
            else:
                __set_environment_variable('TAGS', tag)
    else:
        environ['TAGS'] = ""


def __set_variables_to_behave():
    # tomamos los argumentos en una lista pra que los lea behave
    del sys.argv[1:]
    if environ['TAGS']: #lee los tags que estan en la variable entorno
        tags = environ['TAGS'].split(";")
        for tag in tags:
            sys.argv.append('--tags')
            sys.argv.append(tag)
    sys.argv.append('--outfile')
    sys.argv.append(os.path.join(environ['OUTPUT'], 'logs', "behave.log"))
    print sys.argv[1:]


def __run_test_cases():
    """ para correr los escenarios con behave, llamamos al main
        cuando se genera un error lo tomamos y mostramos."""
    try:
        behave_script.main()
    except Exception as e:
        logging.error("Error al ejecutar los steps: " + str(e))
        traceback.print_exc()


def __generate_execution_reports():
    """ Do reporting. """
    path_info = os.path.join(
        os.path.abspath(REPORTS), INFO_FILE)
    info_file = open(path_info, 'r')
    json_output = info_file.read()
    json_output = json.loads(json_output)
    info_file.close()
    try:
        generate_report_xlsx.generate_report(json_output)
    except Exception:
        logging.error("Test execution reports have not been generated: " +
                      str(sys.exc_info()[1]))
        traceback.print_exc()


def __create_paths():
    """ Create directory hierarchy if not already there. """
    project_paths = [environ['OUTPUT']]

    for path in project_paths:
        if not os.path.exists(path):
            os.makedirs(path)

#para leer el script, esto llama al main
if __name__ == '__main__':
    main()