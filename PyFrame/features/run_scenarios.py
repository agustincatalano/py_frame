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
REPORTS = '..\\reports'
INFO_FILE = 'test_results.json'


def main():
    __parse_arguments()
    __set_environment_variables()
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
    global ARGS
    ARGS = parser.parse_args()


def __set_environment_variables():
    __set_environment_variable('BROWSER', ARGS.browser)
    __set_environment_variable('GRID_URL', ARGS.grid)
    __set_tags_as_environment_variable(ARGS.tags)


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


#para leer el script, esto llama al main
if __name__ == '__main__':
    main()