from configobj import ConfigObj
import os
import re
import json

BEHAVE_TAGS_FILE = "behave.tags"
INFO_FILE = "test_results.json"
OVERALL_STATUS_FILE = "overall_status.json"
REPORTS = '..\\reports\\'

def add_step_info(step, parent_node):
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


def generate_execution_info(context):
    # Get tags filter
    print "\n*******************************************"
    test_results_path = os.path.join(os.path.abspath(REPORTS))
    behave_tags_path = os.path.join(test_results_path, BEHAVE_TAGS_FILE)
    behave_tags_file = open(behave_tags_path, 'r')
    behave_tags = behave_tags_file.readline().strip()
    behave_tags_file.close()
    tag_re = re.compile("@?[\w\d\-\_\.]+")
    overall_status="passed"
    # Generate scenario list
    feature_list = []
    for feature in context._runner.features:
        scenario_list = []
        for feature_scenario in feature.scenarios:
            scenarios = []
            if(feature_scenario.keyword == "Scenario"):
                scenarios = [feature_scenario]
            elif(feature_scenario.keyword == "Scenario Outline"):
                scenarios = feature_scenario.scenarios
            scenario_outline_index = 0
            for scenario in scenarios:
                # Check if scenario was selectable based on tags
                tags = scenario.tags
                # Set MANUAL to False in order filter regarless of it
                tags_filter = behave_tags.replace("@MANUAL", "False")
                # Set scenario tags in filter
                for tag in tags:
                    # Try with and without @
                    tags_filter = tags_filter.replace(' @' + tag + ' ', ' True ')
                    tags_filter = tags_filter.replace(' ' + tag + ' ', ' True ')
                # Set all other tags to False
                other_tags = tag_re.findall(tags_filter)
                for tag in other_tags:
                    if(tag not in ("not", "and", "or", "True", "False")):
                        tags_filter = tags_filter.replace(tag + " ", "False ")
                # Check filter
                if(tags_filter == '' or eval(tags_filter)):
                    # Scenario was selectable
                    scenario_info = dict()
                    for attrib in ('name', 'filename', 'status', 'tags', 'duration'):
                        scenario_info[attrib] = getattr(scenario, attrib)
                    scenario_info['feature'] = scenario.feature.name
                    steps = []
                    for step in scenario.steps:
                        add_step_info(step, steps)
                    scenario_info['steps'] = steps
                    scenario_info['outline_index'] = scenario_outline_index
                    if scenario_info['status'] == "failed":
                        overall_status="failed"
                    scenario_outline_index += 1
                    scenario_list.append(scenario_info)
        if(scenario_list):
            feature_info = dict()
            for attrib in ('name', 'status', 'duration'):
                feature_info[attrib] = getattr(feature, attrib)

            scenario_background = dict()
            if scenario.background:
                steps = []
                for step in scenario.background.steps:
                    add_step_info(step, steps)
                scenario_background['name'] = scenario.background.name
                scenario_background['steps'] = steps
            feature_info['background'] = scenario_background

            feature_info['scenarios'] = scenario_list
            feature_list.append(feature_info)
    environment_info = ""
    if hasattr(context, 'environment'):
        environment_info = context.environment
    output = {
        "environment": environment_info,
        "features": feature_list
    }
    jsonString = json.dumps(output)
    info_file = os.path.join(os.path.abspath(REPORTS), INFO_FILE)
    f = open(info_file, 'w')
    f.write(jsonString)
    f.close()

    status = {"status": overall_status}
    jsonString = json.dumps(status)
    info_file = os.path.join(os.path.abspath(REPORTS), OVERALL_STATUS_FILE)
    f = open(info_file, 'w')
    f.write(jsonString)
    f.close()
