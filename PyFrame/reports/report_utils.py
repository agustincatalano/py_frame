"""Test report utility methods for retrieving summarized
test execution information from features and scenarios"""
import re
import os
import platform
from re import match
from configobj import ConfigObj

CONFIG = ConfigObj(os.path.join(os.getcwd(), "..", "config", "config.cfg"))


def gather_steps(features):
    """retrieve a dictionary with steps used in latest test execution,
    together with the number of times they were executed"""
    steps = {}
    for feature in features:
        for scenario in feature['scenarios']:
            steps_back = [step for step in scenario['steps']]
            for step in scenario['steps'] + steps_back:
                if not step['name'] in steps:
                    steps[step['name']] = {'quantity': 0, 'total_duration': 0,
                                           'appearances': 0}
                steps[step['name']]['appearances'] += 1
                add = 1 if step['status'] != 'skipped' else 0
                steps[step['name']]['quantity'] += add
                steps[step['name']]['total_duration'] += step['duration']
    return steps


def get_summary(features, stories_v1, all_scenarios):
    """ return summary of ejecution
    :param features: features of the ejecutions
    :param stories_v1: the stories in versionOne or other webserver
    :param all_scenarios: all scenarios ejecuted
    :return: tuple the sets
    """
    stories = gather_stories(
        features, CONFIG["story_tag"]["regex"],
        CONFIG["story_tag"]["prefix"])[0]

    story_uncovered = stories_both = {}
    set_yarara = {tag for tag in sum([scenario['tags']
                                      for scenario in all_scenarios], [])}

    if stories_v1:
        stories_v1 = {
            story['Number']: {'Name': story['Name'],
                              'Description': story['Description']}
            for story in stories_v1
        }

        set_v1 = set(stories_v1.keys())
        diference = set_v1 - set_yarara
        both = set_v1 & set_yarara
        only_yarara = set_yarara - set_v1
        story_uncovered = {number: stories_v1[number] for number in diference}
        stories_both = {number: {'scenarios': stories[number],
                                 'description': stories_v1[number]['Name']}
                        for number in both}
    else:
        only_yarara = set_yarara

    return only_yarara, set_yarara, story_uncovered, stories_both


def get_tags_process(all_scenarios):
    """
    return one dict with each key contain one tuple with following format:
    in position zero have list of the dict each dict have tuple with the name
    scenario in posicion 0 and status in posicion 1
    :param all_scenario: list of the scenario
    :return: dict
    """
    tags = set(sum([scenario['tags'] for scenario in all_scenarios], []))
    tags_scenario = {}
    for tag in tags:
        if not match(CONFIG["story_tag"]["regex"], tag):
            tags_scenario[tag] = [
                (scenario['name'], scenario['status'], scenario.get('row', ''))
                for scenario in all_scenarios if tag in scenario['tags']]

    tags_scenario.update(
        {key: (value, set([status for name, status, row in value]),
               set([row for name, status, row in value]))
         for key, value in tags_scenario.items()})

    get_status = lambda x: x.get('failed', False) or x.get('passed', False) \
                           or x.get('skipped', 'skipped')

    tags_process = {key: (value[0], get_status({key: key for key in value[1]}))
                    for key, value in tags_scenario.items()}
    return tags_process


def gather_errors(scenario, retrieve_step_name=False):
    """Retrieve the error message related to a particular failing scenario"""
    error_msg = None
    error_lines = []
    error_step = None
    filename = os.path.splitext(os.path.basename(scenario['filename']))[0]
    folders_in_path = os.path.split(scenario['filename'])[0].split(os.sep)
    total_folders_in_path = len(folders_in_path)
    if total_folders_in_path > 1:
        for index in range(total_folders_in_path, 1, -1):
            filename = folders_in_path[index - 1] + "." + filename
    for line in error_lines:
        regex = r"Failing\sstep:\s[given|when|then|and|or]+ \
                    \s+(.+)\s\.\.\.\s.+"
        match_obj = re.match(regex, line, re.M | re.I)
        if match_obj:
            error_step = match_obj.group(1)
            break
        return error_msg, error_lines, error_step

    else:

        return error_msg, error_lines  # def get_stories_v1():


#     """
#     return all stories storaged in server configured in config file
#     :return: False|list
#     """
#     if CONFIG['software_manager']['manager'] != 'versionOne':
#         return False
#     config_v1 = configuration.get_config_v1()
#
#     if not config_v1 or (config_v1 and '' in config_v1['versionOne'].values()):
#         return False
#
#     os.environ.setdefault('https_proxy', config_v1['proxy']['https_proxy'])
#     v_1 = V1Meta(**config_v1['versionOne'])
#
#     try:
#         if config_v1['query']['use'] == 'where':
#             stories_v1 = v_1.Story.select('Number', 'Name', 'Description')\
#                 .where(**config_v1['where'])
#
#         elif config_v1['query']['use'] == 'filter':
#             stories_v1 = v_1.Story.select('Number', 'Name', 'Description')\
#                 .filter(config_v1['filter']['contain'])
#
#         else:
#             stories_v1 = v_1.Story.select('Number', 'Name', 'Description')
#
#         return [story for story in stories_v1]
#     except:
#         return False



def second_user_format(seconds_float):
    """formating the time from seconds"""
    if seconds_float < 0.009:
        return '00:00:00.0'
    seconds = int(seconds_float)
    hours = seconds / (60 * 60)
    minutes, seconds_ = (seconds / 60) % 60, seconds % 60
    centis = int((seconds_float - seconds) * 100)

    def _normalize_time(unidad):
        """this local function for normalized time
        """
        if int(unidad) <= 9:
            return '0' + str(unidad)
        else:
            return str(int(unidad))

    return '%s:%s:%s.%s' % (
        _normalize_time(hours),
        _normalize_time(minutes),
        _normalize_time(seconds_),
        _normalize_time(centis))


def get_status_traceability(stories):
    """
    This function return dict with status the Number with format for file config
    :param stories: all stories for traceability
    :return: dict with status for each Number
    """
    dict_status = {}
    for key, story in stories.items():
        passed = all([scenary['status'] in ('passed', 'skipped')
                      for scenary in story])
        if passed:
            passed = any([scenary['status'] == 'passed' for scenary in story])
            dict_status[key] = 'passed' if passed else 'skipped'
        else:
            dict_status[key] = 'failed'
    return dict_status


def _get_root_xml(output_path, filename):
    """
    return root xml
    :param output_path: path of the output folder
    :param filename: name file
    :return: path root
    """
    path = "TESTS-" + filename + ".xml"
    path = os.path.join(output_path, path)
    if platform.system() == 'Windows':
        from lxml import etree

        parser = etree.XMLParser(recover=True)
        path = etree.parse(path, parser=parser)
    else:
        import xml.etree.ElementTree as ET

        path = ET.parse(path)
    return path.getroot()


def normalize(string):
    """
    Normalize specified string by removing characters
    that might cause issues when used as file/folder names
    """
    return string.replace(" ", "_") \
        .replace("(", "") \
        .replace(")", "") \
        .replace("\"", "") \
        .replace("/", "")
