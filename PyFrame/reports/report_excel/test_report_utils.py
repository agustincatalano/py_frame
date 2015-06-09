import re
import os
import xml.etree.ElementTree as ET

def __gather_steps(features):
    steps = {}
    for feature in features:
        for scenario in feature['scenarios']:
            for step in scenario['steps']:
                if not step['name'] in steps:
                    steps[step['name']] = 0
                steps[step['name']] += 1
    return steps

def __gather_stories(features, tag_selector, tag_prefix):
    stories = {}
    storyless_features = []
    pattern = re.compile(tag_selector,re.IGNORECASE)
    for feature in features:
        for scenario in feature['scenarios']:
            matching_stories = filter(pattern.match,scenario['tags'])
            for story in matching_stories:
                story = story.replace(tag_prefix,"")
                if not story in stories:
                    stories[story] = []
                stories[story].append({'name' : scenario['name'],'status' : scenario['status']})

            if not matching_stories:
                storyless_features.append({'name' : scenario['name'],'status' : scenario['status']})
    return stories, storyless_features

def _gather_errors(scenario, retrieve_step_name=False):
    error_msg = None
    error_lines = []
    error_step = None
    filename = os.path.splitext(os.path.basename(scenario['filename']))[0]
    folders_in_path = os.path.split(scenario['filename'])[0].split(os.sep)
    total_folders_in_path = len(folders_in_path)
    if total_folders_in_path > 1:
        for x in range(total_folders_in_path, 1, -1):
            filename = folders_in_path[x-1] + "." + filename
    results_xml = ET.parse(os.path.join( os.path.abspath(os.environ.get('OUTPUT')),"TESTS-"+filename+".xml"))
    #print "Results file: " + os.path.join( os.path.abspath(os.environ.get('OUTPUT')),"TESTS-"+filename+".xml")
    report_root = results_xml.getroot()
    for testcase in report_root.iter('testcase'):
        if testcase.get('name') == scenario['name']:
            if testcase.find('failure') is not None:
                error_msg = testcase.find('failure').get('message') 
                if testcase.find('failure').text:
                    error_lines = testcase.find('failure').text.split('\n')
            elif testcase.find('error') is not None:
                error_msg = testcase.find('error').get('message') 
                if testcase.find('error').text:
                    error_lines = testcase.find('error').text.split('\n')
    if retrieve_step_name:
        for line in error_lines:
            regex = "Failing\sstep:\s[given|when|then|and|or]+\s+(.+)\s\.\.\.\s.+"
            matchObj = re.match(regex, line, re.M|re.I)
            if matchObj:
                error_step = matchObj.group(1)
                break
        return error_msg, error_lines, error_step
    else:
        return error_msg, error_lines
