"""
This module generate json with representation data of the ejecution
"""
import os
import re
import json

BEHAVE_TAGS_FILE = os.path.join("behave", "behave.tags")
INFO_FILE = "results.json"
OVERALL_STATUS_FILE = "overall_status.json"


def add_step_info(step, parent_node):
    """
    craate dictionary with info the step and add in parent node
    :param step: the object step with info the step
    :param parent_node: the parent node the step
    :return: None
    """
    step_info = dict()
    for attrib in ('step_type', 'name', 'text', 'duration', 'status'):
        step_info[attrib] = getattr(step, attrib)
    if step.table:
        step_info['table'] = {}
        for heading in step.table.headings:
            step_info['table'][heading] = []
            for row in step.table:
                step_info['table'][heading].append(row[heading])
    parent_node.append(step_info)


def add_step_info_background(step, parent_node):
    """
    craate dictionary with info the background step and add in parent node
    :param step: the object step with info the step
    :param parent_node: the parent node the step
    :return: None
    """
    step_info = dict()
    for attrib in ('step_type', 'name', 'text', 'status'):
        step_info[attrib] = getattr(step, attrib)
    step_info['duration'] = step.duration
    if step.table:
        step_info['table'] = {}
        for heading in step.table.headings:
            step_info['table'][heading] = []
            for row in step.table:
                step_info['table'][heading].append(row[heading])
    parent_node.append(step_info)


def generate_execution_info(context):
    """
    Generate json with representation data of the ejecution
    :param context: the context of the ejecution
    :return: object json with representation of the ejecution
    """
    # Get tags filter
    print "\n*******************************************"
    # Generate scenario list
    feature_list = []
    for feature in context._runner.features:
        scenario_list = []

        for feature_scenario in feature.scenarios:
            scenarios = []
            if feature_scenario.keyword == "Scenario":
                scenarios = [feature_scenario]
            elif feature_scenario.keyword == "Scenario Outline":
                scenarios = feature_scenario.scenarios

            overall_status, scenario_list = _processing_scenarios(scenarios,
                                                                 scenario_list)
        if scenario_list:
            feature_info = dict()
            for attrib in ('name', 'status', 'duration'):
                feature_info[attrib] = getattr(feature, attrib)
            feature_info['scenarios'] = scenario_list
            feature_info['background'] = _processing_background_feature(feature)
            feature_list.append(feature_info)
    environment_info = ""
    if hasattr(context, 'environment'):
        environment_info = context.environment
    output = {
        "environment": environment_info,
        "features": feature_list
    }
    path_info = os.path.join(os.path.abspath(os.environ.get('OUTPUT')),
                             INFO_FILE)
    file_info = open(path_info, 'w')
    file_info.write(json.dumps(output))
    file_info.close()

    path_info = os.path.join(os.path.abspath(os.environ.get('OUTPUT')),
                             OVERALL_STATUS_FILE)
    file_info = open(path_info, 'w')
    file_info.write(json.dumps({"status": overall_status}))
    file_info.close()


def _processing_background(scenario):
    """
    Proccesing data background of the scenario
    :param scenario: object scenario
    :return: dictionary with background scenario
    """
    scenario_background = {'duration': 0.0, 'steps': []}
    if scenario.background:
        steps = []
        for step in scenario._background_steps:
            add_step_info_background(step, steps)
            scenario_background['duration'] += step.duration
        scenario_background['name'] = scenario.background.name
        scenario_background['steps'] = steps
    # return scenario_background


def _processing_background_feature(feature):
    """
    Proccesing data background of the feature
    :param feature: object feature
    :return: dictionary with background feature
    """
    feature_background = {}
    if feature.background:
        steps = []
        for step in feature.background.steps:
            add_step_info_background(step, steps)
        feature_background['name'] = feature.background.name
        feature_background['steps'] = steps
    return feature_background


def _processing_scenarios(scenarios, scenario_list):
    """ Proccesing all scenario of the one feature
    :param scenarios: all scenarios
    :param scenario_list: list of the scenario to add
    :return: tuble with overall_status and scenario list
    """
    tag_re = re.compile('@?[\w\d\-\_\.]+')
    scenario_outline_index = 0
    overall_status = "passed"
    for scenario in scenarios:
        # Set MANUAL to False in order filter regarless of it
        tags_filter = os.environ['TAGS']
        # Set scenarioio tags in filter
        for tag in scenario.tags:
            # Try with and without @
            tags_filter = tags_filter.replace(' @' + tag + ' ', ' True ').replace(' ' + tag + ' ', ' True ')
        # Set all other tags to False
        for tag in tag_re.findall(tags_filter):
            if tag not in ("not", "and", "or", "True", "False"):
                tags_filter = tags_filter.replace(tag + " ", "False ")
        # Check filter
        if tags_filter == '' or tags_filter:
            # Scenario was selectable
            scenario_info = dict()
            for attrib in ('name', 'filename', 'status', 'tags',
                           'duration'):
                scenario_info[attrib] = getattr(scenario, attrib)
            scenario_info['feature'] = scenario.feature.name
            steps = []
            for step in scenario.steps:
                add_step_info(step, steps)
            scenario_info['steps'] = steps
            scenario_info['outline_index'] = scenario_outline_index
            if scenario_info['status'] == "failed":
                overall_status = "failed"
            scenario_outline_index += 1
            scenario_info['background'] = _processing_background(scenario)
            scenario_list.append(scenario_info)

    return overall_status, scenario_list
