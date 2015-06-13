"""This module generate xlsx file with reports the ejecutions
"""
import os

from openpyxl import load_workbook, Workbook
from openpyxl import charts
import openpyxl.cell.cell as Cell
import string
from math import ceil
from reports.report_excel.patched_charts import PieChartWriter
from openpyxl.styles import Color
from openpyxl.styles import colors
from openpyxl.worksheet.worksheet import AutoFilter
from openpyxl.charts import (
    PieChart,
    Reference,
    Series
)

from reports.report_utils import (
    gather_steps,
    gather_errors,
    # get_stories_v1,
    # get_coverage,
    second_user_format,
    get_tags_process,
    # get_traceability
)

from reports.report_excel.report_xlsx_utils import (
    __create_styles,
    # __create_borders,
    __copy_template_table__
)

from configobj import ConfigObj

TRACEABILITY = 'Traceability'
END_COLUMN_UTIL = 9
EXECUTION_REPORT = "Test Execution Report"
EXECUTION_REPORT_SEL = "'%s'" % EXECUTION_REPORT + "!$%s"
EXECUTION_REPORT_FUN = "=+'%s'" % EXECUTION_REPORT + "!$%s"


charts.PieChartWriter = PieChartWriter


CONFIG = ConfigObj(os.path.join(os.getcwd(), "..", "config", "config.cfg"))


def generate_report(output):
    """genereate report in format xlsx"""
    features = output["features"]
    # open template for copying xls
    work_book, work_sheet, template = _instance_template()

    # Set first row
    row_index = 1
    #
    # row_index = __to_xlsx_feature_summary(
    #     template, work_sheet, features, row_index)
    #
    # row_index = __to_xlsx_general_summary(work_sheet, features, row_index - 1)
    row_summary = row_index - 1
    # row_index += 3

    # row_index = __export_to_xlsx_test_cases(template, work_sheet, features,
    #                                         row_index, row_summary)
    #
    __add_conditional_rule_scenario(work_book, row_index, row_summary + 2)
    #
    # # Save file
    scenarios = sum([feature['scenarios'] for feature in features], [])
    # rows, stories_with_ref = _processing_rows_traceability(features, scenarios)
    # ref_scenaries = {scenary['name']: scenary['row'] for scenary in scenarios}
    # ref_features = {feature['name']: ['G' + str(row['row'])
    #                                   for row in feature['scenarios']]
    #                 for feature in features}
    #
    # # __export_to_xlsx_traceability(work_book, rows, stories_with_ref,
    # #                               ref_scenaries)
    #
    # __add_formatting_summary(ref_features, work_book, row_summary)

    __export_to_xlsx_steps(work_book, features)

    __export_to_xlsx_charts(work_book, features, row_summary + 1,)
    #
    work_book.save(
           os.path.join(os.path.abspath(os.environ.get('OUTPUT')),
                      'results.xlsx'))


def _instance_template():
    """instance template for creacion of the exel file"""
    # Open template sheet
    template_file = os.path.join(os.path.dirname(os.path.realpath(__file__)),
                                 "user_report_template.xlsx")
    template_book = load_workbook(template_file)
    template = template_book.get_active_sheet()

    # Create sheet
    work_book = Workbook()
    work_sheet = work_book.get_active_sheet()
    work_sheet.title = EXECUTION_REPORT
    # Copy dimensions
    work_sheet.row_dimensions = template.row_dimensions
    work_sheet.column_dimensions = template.column_dimensions
    # Set first row

    return work_book, work_sheet, template


def __add_conditional_rule_scenario(work_book, row_index, start_index):
    """
    creating rule for each scenario
    :param work_sheet:
    :param row_index:
    :param start_index:
    :return:
    """
    work_sheet = work_book.get_active_sheet()
    # pallette = __create_styles(work_book, work_sheet)[0]

    work_sheet.conditional_formatting. \
        add("A" + str(start_index) + ":I" + str(row_index),
                      {'type': 'expression',
                       'formula': [EXECUTION_REPORT_SEL % 'G' +
                                   str(start_index) + '="Skipped"'],
                       'stopIfTrue': '1'})

    work_sheet.conditional_formatting. \
        add("A" + str(start_index) + ":I" + str(row_index),
                      {'type': 'expression',
                       'formula': [EXECUTION_REPORT_SEL % 'G' +
                                   str(start_index) + '="Passed"'],
                       'stopIfTrue': '1'})

    work_sheet.conditional_formatting. \
        add("A" + str(start_index) + ":I" + str(row_index),
                      {'type': 'expression',
                       'formula': [EXECUTION_REPORT_SEL % 'G' +
                                   str(start_index) + '="Failed"'],
                       'stopIfTrue': '1'})


def __add_formatting_summary(ref_features, work_book, row_summary):
    """add formatting conditional for general summary """
    work_sheet = work_book[EXECUTION_REPORT]
    # pallette = __create_styles(work_book, work_sheet)[0]
    for index in range(row_summary - 2):
        cell = work_sheet.cell(row=(2 + index), column=0)
        value = cell.value
        passed, failed, skip = __generate_formula_for_rule(ref_features[value])

        work_sheet.conditional_formatting. \
            addCustomRule('A%i:I%i' % (cell.row, cell.row),
                          {'type': 'expression',
                           'formula': [skip],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule('A%i:I%i' % (cell.row, cell.row),
                          {'type': 'expression',
                           'formula': [passed],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule('A%i:I%i' % (cell.row, cell.row),
                          {'type': 'expression',
                           'formula': [failed],
                           'stopIfTrue': '1'})


def __export_to_xlsx_tag_mapping(work_book, all_scenarios, ref_scenaries):
    """Export tag to xlsx file"""
    work_sheet = Workbook.create_sheet(work_book)
    work_sheet.title = 'Tags'
    work_sheet['B1'].value = 'Tag'
    tags_process = get_tags_process(all_scenarios)
    # formated for create table
    rows = [{'value': scenary[0], 'color': scenary[1], 'tag': tag,
             'color_tag': value[1]}
            for tag, value in tags_process.items() for scenary in value[0]]

    dict_ref = _get_ref_for_tags(tags_process)
    __create_table(rows, work_book, work_sheet, dict_ref, ref_scenaries)


def _get_ref_for_tags(tags):
    """get references for tags"""
    dict_ref = {}
    for tag, value in tags.items():
        dict_ref[tag] = ['G' + str(row[2]) for row in value[0]]
    return dict_ref


def __export_to_xlsx_steps(work_book, features):
    """export to xlsx steps"""
    steps = gather_steps(features)
    work_sheet = Workbook.create_sheet(work_book)
    work_sheet.title = 'Execution Steps'
    row_index = 1
    # borders = __create_borders()
    work_sheet['B1'].value = 'Step'
    # work_sheet['B1'].style.borders = borders['full']
    work_sheet['C1'].value = 'Appearances'
    # work_sheet['C1'].style.borders = borders['full']
    work_sheet['D1'].value = 'Executions'
    # work_sheet['D1'].style.borders = borders['full']
    work_sheet['E1'].value = 'Avg Time'
    # work_sheet['E1'].style.borders = borders['full']
    work_sheet['F1'].value = 'Total Time'
    # work_sheet['F1'].style.borders = borders['full']
    for step in sorted(steps):
        cell = work_sheet.cell(row=row_index, column=1)
        cell.value = step
        # cell.style.borders = borders['full']
        cell.offset(column=1).value = steps[step]['appearances']
        # cell.offset(column=1).style.borders = borders['full']
        cell.offset(column=2).value = steps[step]['quantity']
        # cell.offset(column=2).style.borders = borders['full']
        cell.offset(column=3).value = '%.2fs' % \
                                      (steps[step]['total_duration'] /
                                       (steps[step]['quantity'] or 1))

        # cell.offset(column=3).style.borders = borders['full']
        cell.offset(column=4).value = '%.2fs' % steps[step]['total_duration']
        # cell.offset(column=4).style.borders = borders['full']
        if len(step) > work_sheet.column_dimensions['B'].width:
            work_sheet.column_dimensions['B'].width = len(step)
        work_sheet.column_dimensions['C'].width = 10
        work_sheet.column_dimensions['D'].width = 10
        work_sheet.column_dimensions['E'].width = 10
        row_index += 1
    # work_sheet.auto_filter = "B1:B" + str(row_index)


def __generate_dict_for_rules(work_book, stories_with_ref):
    """
    generates a Dictionary where each value is a list of references columns of
    scenarios of the execution report. Examples of the result:
    {{'B-10101':['B1','B2']}*}
    :param work_book:
    :param stories_both: dict the stories that both
    :return: dict
    """
    work_sheet = work_book[EXECUTION_REPORT]
    dict_ref = {}
    for number, value in stories_with_ref.items():
        list_ref = []
        scenarios = [scenario['name'] for scenario in value['scenarios']]
        for i, cell in enumerate(work_sheet.columns[1]):
            if cell.value in scenarios:
                list_ref.append('G' + str(i + 1))
        dict_ref[number] = list_ref
    return dict_ref


def __generate_formula_for_rule(cells):
    """
    formulated to generate the rule based on the state of other columns

    :param cells: las celdas que debe tener en cuenta
    :return: tuple with string for rule
    """
    ors = ["""OR('%s'!$%s="Passed",'%s'!$%s="Skipped")""" %
           (EXECUTION_REPORT, cell, EXECUTION_REPORT, cell) for cell in cells]

    rule_passed = 'AND(%s)' % ','.join(ors)

    ors = ["""'%s'!$%s="Failed",'%s'!$%s="Skipped" """ %
           (EXECUTION_REPORT, cell, EXECUTION_REPORT, cell) for cell in cells]

    rule_failed = 'OR(%s)' % ','.join(ors)

    is_skipped = ["""'%s'!$%s="Skipped" """ % (EXECUTION_REPORT, cell)
                  for cell in cells]

    rule_skipped = 'AND(%s)' % ','.join(is_skipped)

    return rule_passed, rule_failed, rule_skipped


def __generate_formula_for_tags(cells, work_sheet, row):
    """
    formulated to generate the rule based on the state of other columns

    :param cells: las celdas que debe tener en cuenta
    :param work_sheet: this work sheet to jobs
    :param cell: cell to apply rules
    :return: tuple with string for rule
    """
    index_of_letter = string.ascii_uppercase[5:]  # Generate ABC From F
    start = 0
    cant = 30
    partitioned = 0
    portions = len(cells) / cant + 1
    while start <= len(cells):
        portion = cells[start:start + cant]
        ors = ["""OR('%s'!$%s="Passed",'%s'!$%s="Skipped")""" %
               (EXECUTION_REPORT, cell, EXECUTION_REPORT, cell)
               for cell in portion]
        # storing in cell for that not passed of the characters limit
        # for after joining
        cell_aux = work_sheet.cell(index_of_letter[partitioned] + row[1:])
        cell_aux.value = '=AND(%s)' % ','.join(ors)

        # generate formula for failed
        ors = ["""'%s'!$%s="Failed",'%s'!$%s="Skipped" """ %
               (EXECUTION_REPORT, cell, EXECUTION_REPORT, cell)
               for cell in portion]
        cell_aux.style.font.color.index = Color(colors.WHITE)

        cell_aux = work_sheet.cell(index_of_letter[portions + partitioned]
                                   + row[1:])
        cell_aux.value = '=OR(%s)' % ','.join(ors)
        cell_aux.style.font.color.index = Color(colors.WHITE)

        # generate formula for skipped
        is_skipped = ["""'%s'!$%s="Skipped" """ % (EXECUTION_REPORT, cell)
                      for cell in portion]
        cell_aux = work_sheet.cell(index_of_letter[2 * portions + partitioned]
                                   + row[1:])
        cell_aux.style.font.color.index = Color(colors.WHITE)
        cell_aux.value = '=AND(%s)' % ','.join(is_skipped)
        start += cant
        partitioned += 1

    references = range(portions) or [0]

    # join references column, return rule for passed, rule for skipped,
    #  and rule for failed

    return 'AND(%s)' % ','.join([index_of_letter[col] + row[1:]
                                 for col in references]), \
           'OR(%s)' % ','.join([index_of_letter[portions + col] + row[1:]
                                for col in references]), \
           'AND(%s)' % ','.join([index_of_letter[portions * 2 + col]
                                 + row[1:] for col in references])


def _add_rule_for_status(work_sheet, cell, row, dict_ref, ref_scenaries):
    """Add rule for status dependent of testcase """
    pallette = __create_styles(work_sheet.parent, work_sheet)[0]
    if not dict_ref:
        pass
    elif not dict_ref.get(row['tag'], False):
        pass
    else:
        if work_sheet.title == 'Tags':
            rule_passed, rule_faill, rule_skipped = \
                __generate_formula_for_tags(dict_ref[row['tag']],
                                            work_sheet, cell)
        else:
            rule_passed, rule_faill, rule_skipped = \
                __generate_formula_for_rule(dict_ref[row['tag']])
        #  This following formating is for each cell and dependent the one set
        #  of the scenarios
        work_sheet.conditional_formatting. \
            addCustomRule('B%s' % cell[1:],
                          {'type': 'expression',
                           'dxfId': pallette['grey'],
                           'formula': [rule_skipped],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule('B%s' % cell[1:],
                          {'type': 'expression',
                           'dxfId': pallette['green'],
                           'formula': [rule_passed],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule('B%s' % cell[1:],
                          {'type': 'expression',
                           'dxfId': pallette['red'],
                           'formula': [rule_faill],
                           'stopIfTrue': '1'})



        # The following formating is for each cell and dependent of the one
        # specific scenario
        work_sheet.conditional_formatting. \
            addCustomRule(cell,
                          {'type': 'expression',
                           'dxfId': pallette['grey'],
                           'formula': [EXECUTION_REPORT_SEL % 'G' +
                                       str(ref_scenaries[row['value']])
                                       + '="Skipped"'],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule(cell,
                          {'type': 'expression',
                           'dxfId': pallette['green'],
                           'formula': [EXECUTION_REPORT_SEL % 'G' +
                                       str(ref_scenaries[row['value']])
                                       + '="Passed"'],
                           'stopIfTrue': '1'})

        work_sheet.conditional_formatting. \
            addCustomRule(cell,
                          {'type': 'expression',
                           'dxfId': pallette['red'],
                           'formula': [EXECUTION_REPORT_SEL % 'G' +
                                       str(ref_scenaries[row['value']])
                                       + '="Failed"'],
                           'stopIfTrue': '1'})


def __export_to_xlsx_traceability(work_book, rows, stories_ref, ref_scenaries):
    """export to xlsx traceability"""
    work_sheet = Workbook.create_sheet(work_book)

    work_sheet.title = TRACEABILITY
    work_sheet['B1'].value = 'Story/bug'
    dict_ref = __generate_dict_for_rules(work_book, stories_ref)
    __create_table(rows, work_book, work_sheet, dict_ref, ref_scenaries)


def __to_xlsx_general_summary(work_sheet, features, row_index):
    """export to xlsx general symmary"""
    style_attrs = ("font", "fill", "alignment", "number_format")
    fields = {"total": 1, "number": 2, "executed": 3, "passed": 4, "failed": 5,
              "skipped": 6, "status": 7, "rate": 8, "duration": 9}

    model_row_index = row_index - 1
    values = dict()
    for field in fields.keys():
        values[field] = 0

    # Fill values
    first_test_index = 7 + len(features)
    total_scenarios = sum(len(feature['scenarios']) for feature in features)
    status_calc_range = 'G' + str(first_test_index) + ':G' + \
                        str(first_test_index + total_scenarios)

    values['total'] = 'Total' + ' ' * 9 + str(len(features))
    values['number'] = total_scenarios
    values['passed'] = '=COUNTIF(' + status_calc_range + ',"Passed")'
    values['failed'] = '=COUNTIF(' + status_calc_range + ',"Failed")'
    values['executed'] = '=COUNTIF(' + status_calc_range + \
                         ',"Failed") + COUNTIF(' + status_calc_range + ',"Passed")'
    values['skipped'] = '=COUNTIF(' + status_calc_range + ',"Skipped")'
    values['rate'] = '=FIXED(COUNTIF(' + status_calc_range + \
                     ',"Passed")/' + str(total_scenarios) + '*100, 2)'
    values['duration'] = second_user_format(sum(feature['duration']
                                                for feature in features))
    values['status'] = '=FIXED(COUNTIF(' + status_calc_range + \
                       ',"<>Skipped")/' + str(total_scenarios) + '*100, 2)'
    # Fill info
    for field in fields.keys():
        cell = work_sheet.cell(row=row_index, column=fields[field])
        model_cell = work_sheet.cell(row=model_row_index,
                                     column=fields[field])
        for attr in style_attrs:
            setattr(cell.style, attr, getattr(model_cell.style, attr))

        cell.value = values[field]

    # Advance row
    row_index += 1
    # Return next free row
    return row_index


def __create_table(rows, work_book, work_sheet, dict_ref, ref_scenaries):
    """Create table with rows of the two column"""
    fills = __create_styles(work_book, work_sheet)[1]
    fills = {'failed': fills['red'], 'passed': fills[
        'green'], 'skipped': fills['grey'], 'yellow': fills['yellow']}
    # borders = __create_borders()
    row_index = 1
    # work_sheet['B1'].style.borders = borders['full']
    work_sheet.column_dimensions['B'].width = len(work_sheet['B1'].value) + 5
    work_sheet['C1'].value = 'Scenario'
    # work_sheet['C1'].style.borders = borders['full']
    for i, row in enumerate(rows):
        cell = work_sheet.cell(row=row_index, column=2)
        cell.value = row['value']
        cell.offset(column=0).style.fill = fills[row['color']]
        cell.style.font.color.index = Color(colors.BLACK)
        # cell.style.borders = borders['full']

        cell = work_sheet.cell(row=row_index, column=1)
        # cell.offset(column=0).style.fill = fills[row['color_tag']]
        value_old = row['tag'] if i == 0 else rows[i - 1]['tag']
        if len(row['tag']) > work_sheet.column_dimensions['B'].width:
            work_sheet.column_dimensions['B'].width = len(row['tag']) + 5
        if value_old != row['tag'] or i == 0:
            cell.offset(column=0).value = row['tag']
            # cell.offset(column=0).style.borders = borders['top']
        else:
            cell.offset(column=0).value = ''
        #     # cell.offset(column=0).style.borders = borders['side']
        # if len(rows) - 1 == i and row['tag'] == value_old:
        #     # cell.offset(column=0).style.borders = borders['bottom']
        # if len(rows) - 1 == i and row['tag'] != value_old:
        #     # cell.offset(column=0).style.borders = borders['full']

        cell.offset(column=0).style.fill = fills[row['color_tag']]
        cell.offset(column=0).style.font.color.index = Color(colors.BLACK)
        if len(row['value']) > work_sheet.column_dimensions['C'].width:
            work_sheet.column_dimensions['C'].width = len(row['value']) + 5
        row_index += 1
        _add_rule_for_status(work_sheet, 'C' + str(i + 2), row, dict_ref,
                             ref_scenaries)

    work_sheet.auto_filter = "B1:C1" + str(row_index)
    work_sheet['E2'].value = work_sheet['B1'].value + ' Legend: '
    work_sheet['E3'].value = 'All scenarios passed'
    work_sheet['E3'].style.fill = fills['passed']
    work_sheet['E4'].value = 'At least one scenario failed'
    work_sheet['E4'].style.fill = fills['failed']
    work_sheet['E5'].value = 'At least one scenario skipped'
    work_sheet['E5'].style.fill = fills['skipped']
    work_sheet.column_dimensions['E'].width = 60

    #work_sheet.column_dimensions['E'].width = len(work_sheet['E6'].value)


def __to_xlsx_feature_summary(template, work_sheet, features, row_index):
    """ export to xlsx feature summary
    """
    fields = {"feature": 1, "number": 2, "executed": 3, "passed": 4,
              "failed": 5, "skipped": 6, "status": 7, "rate": 8, "duration": 9}
    row_index = __copy_template_table__(template, work_sheet, row_index,
                                        "Summary")

    last_test_index = row_index + 5 + len(features)
    model_row_index = row_index
    style_attrs = ("font", "fill", "alignment", "number_format") #"borders",
    # Fill values
    for feature in features:
        # Get info
        values = dict()
        for field in fields.keys():
            values[field] = 0
        values['feature'] = feature['name']
        values['number'] = len(feature['scenarios'])

        status_calc_range = 'G' + str(last_test_index) + ':G' + \
                            str(last_test_index + len(feature['scenarios']) - 1)
        last_test_index += len(feature['scenarios'])

        values['status'] = '=COUNTIF(' + status_calc_range + \
                           ',"<>Skipped")/' + str(len(feature['scenarios'])) + '*100'
        values['passed'] = '=COUNTIF(' + status_calc_range + ',"Passed")'
        values['failed'] = '=COUNTIF(' + status_calc_range + ',"Failed")'
        values['executed'] = '=COUNTIF(' + status_calc_range + \
                             ',"Failed") + COUNTIF(' + status_calc_range + ',"Passed")'
        values['skipped'] = '=COUNTIF(' + status_calc_range + ',"Skipped")'
        values['rate'] = '=COUNTIF(' + status_calc_range + \
                         ',"Passed")/' + str(len(feature['scenarios'])) + '*100'
        values['duration'] = second_user_format(feature['duration'])
        # Fill info
        for field in fields.keys():
            model_cell = work_sheet.cell(
                # row=model_row_index - 1, column=fields[field] - 1)
                row=model_row_index, column=fields[field])
            # cell = work_sheet.cell(row=row_index - 1, column=fields[field] - 1)
            cell = work_sheet.cell(row=row_index, column=fields[field])
            for attr in style_attrs:
                setattr(cell.style, attr, getattr(model_cell.style, attr))
            cell.value = values[field]
        if model_row_index != row_index:
            _set_row_dimension(work_sheet, row_index - 1)
        # Advance row
        row_index += 1
    # Return next free row
    return row_index


def _set_row_dimension(work_sheet, row_index):
    """Set row dimension"""
    big = max([1] + [len(cell.value) for cell in work_sheet.rows[row_index]
                     if isinstance(cell.value, unicode) or
                     isinstance(cell.value, str)])

    dimension = (ceil(big / 40.0) or 1) * 16
    work_sheet.row_dimensions[row_index].height = int(dimension)


def __export_to_xlsx_test_cases(template, work_sheet, features, row_index,
                                row_summary):
    """export to xlsx test cases"""
    fields = {"feature": 1, "scenario": 2, "steps": 3,
              "automatic": 4, "status": 5, "duration": 6,
              "tester": 7}
    row_index = __copy_template_table__(template, work_sheet, row_index,
                                        "Test Cases")
    model_row_index = row_index
    # Fill values
    for feature in features:
        # Get info
        values = dict()
        for field in fields.keys():
            values[field] = 0
        values['feature'] = feature['name']
        # Get first automatic cases
        scenarios_auto = []
        scenarios_manual = []
        for scenario in feature['scenarios']:
            if "MANUAL" in scenario['tags']:
                scenarios_manual.append(scenario)
            else:
                scenarios_auto.append(scenario)
        scenarios = scenarios_auto + scenarios_manual
        for scenario in scenarios:
            values['scenario'] = scenario['name']
            values['status'] = scenario['status'].capitalize()
            values['duration'] = second_user_format(scenario['duration'])
            if "MANUAL" in scenario['tags']:
                values['automatic'] = "Manual"
                values['tester'] = ''
            else:
                values['automatic'] = "Automatic"
                values['tester'] = "Automatic"
            values['steps'] = ''
            _processing_step(scenario, values)
            # Fill info
            scenario['row'] = row_index
            _fill_info(work_sheet, {}, model_row_index, row_index, values)
            if model_row_index != row_index:
                _set_row_dimension(work_sheet, row_index - 1)
            # Advance row
            row_index += 1
        for index in range(row_index - row_summary):
            work_sheet.merge_cells(range_string="C{0}:E{1}". \
                                   format(row_summary + index + 3,
                                          row_summary + index + 3))

    # Return next free row
    return row_index


def _fill_info(work_sheet, scenarios, model_row_index, row_index, values):
    """ Fill info """
    # style_attrs = ("borders", "font", "fill", "alignment", "number_format")
    style_attrs = ("font", "fill", "alignment", "number_format")
    fields = {"feature": 1, "scenario": 2, "steps": 3, "empty1": 4, "empty2": 5,
              "automatic": 6, "status": 7, "duration": 8,
              "tester": 9}
    #work_sheet.row_dimensions[model_row_index -1].height = 15
    for field in fields.keys():
        model_cell = work_sheet.cell(
            row=model_row_index, column=fields[field])
        cell = work_sheet.cell(row=row_index,
                               column=fields[field])
        if field == 'status':
            scenarios[values['scenario']] = \
                Cell.absolute_coordinate(cell.coordinate)
        for attr in style_attrs:
            setattr(cell.style, attr, getattr(model_cell.style, attr))
        if field in ("empty1", "empty2"):
            cell.value = ''
        else:
            cell.value = values[field]


def _processing_step(scenario, values):
    """Processing step of the scenary"""
    for step in scenario['steps']:
        fail_string = ""
        if step['status'] == "failed":
            error_msg, error_lines = gather_errors(scenario)[0:2]
            fail_string = " (FAILED!)\n"
            if error_msg:
                fail_string += "-" + error_msg + "\n"
                for line in error_lines:
                    fail_string = fail_string + "-" + line + "\n"
        values['steps'] = values[
                              'steps'] + step['step_type'].capitalize() + ' ' \
                          + step['name'] + fail_string + '\n'
        if 'table' in step:
            table_length = 0
            for key in step['table']:
                values['steps'] = values['steps'] + '|' + key
                table_length = len(step['table'][key])
            values['steps'] += '|\n'
            for i in range(table_length):
                for key in step['table']:
                    values['steps'] += '|' + step['table'][key][i]
                values['steps'] += '|\n'


def __export_to_xlsx_charts(work_book, scenarios, row_summary):
    """export to xlsx charts"""

    scenarios = [scenario for scenario in scenarios
                 if not any([tag in scenario.get('tags', [])
                             for tag in os.environ.get('SKIP_TAGS',
                                                       '').split(',')])]

    total_passed = sum(scenario['status'].count("passed")
                       for scenario in scenarios)

    total_failed = sum(scenario['status'].count("failed")
                       for scenario in scenarios)

    total_skipped = sum(scenario['status'].count("skipped")
                        for scenario in scenarios)

    work_sheet2 = Workbook.create_sheet(work_book)
    work_sheet2.title = "Metrics"

    work_sheet2.cell(
        "A1").value = "Charts are not dinamically updated.If test status" \
                      " changes in test execution please update the " \
                      "summarized value"

    # work_sheet2.cell("O2").value = "Test automation Rate"
    work_sheet2.cell("P3").value = total_passed + total_failed
    # work_sheet2.cell("P3").style.f = Color(colors.(colors.WHITE)
    work_sheet2.cell("P4").value = total_skipped
    # work_sheet2.cell("P4").style.font.color.index = Color(colors..WHITE
    # work_sheet2.cell("O3").value = "Automated"
    # work_sheet2.cell("O4").value = "Not Automated"
    values = Reference(work_sheet2, (2, 15), pos2=(3, 15))
    labels = Reference(work_sheet2, (2, 14), pos2=(3, 14))
    # __add_chart(work_sheet2, values, labels, "Test Automation Rate", top=60)

    # work_sheet2.cell("O6").value = "Test execution status"
    work_sheet2.cell("P7").value = total_passed
    # work_sheet2.cell("P7").style.font.color.index = Color(colors..WHITE
    work_sheet2.cell("P8").value = total_failed
    # work_sheet2.cell("P8").style.font.color.index = Color(colors..WHITE
    work_sheet2.cell("P9").value = total_skipped
    # work_sheet2.cell("P9").style.font.color.index = Color(colors..WHITE
    work_sheet2.cell("O7").value = "Passed"
    work_sheet2.cell("O8").value = "Failed"
    work_sheet2.cell("O9").value = "Skipped"
    values = Reference(work_sheet2, (6, 15), pos2=(8, 15))
    labels = Reference(work_sheet2, (6, 14), pos2=(8, 14))
    __add_chart(work_sheet2, values, labels, "Test Execution Status", top=460)


    # Executed
    work_sheet2.cell("P3").value = EXECUTION_REPORT_FUN % 'C' + str(row_summary)

    # Passed
    work_sheet2.cell("P7").value = EXECUTION_REPORT_FUN % 'D' + str(row_summary)

    # Failed
    work_sheet2.cell("P8").value = EXECUTION_REPORT_FUN % 'E' + str(row_summary)

    # Skipped
    work_sheet2.cell("P9").value = EXECUTION_REPORT_FUN % 'F' + str(row_summary)



def __add_chart(work_sheet, values, labels, title, top=400):
    """add chart"""
    series = Series(values, title=title, labels=labels, color=Color(colors.GREEN))
    # series.color = Color(colors.BLACK)
    chart = PieChart()
    chart.append(series)
    chart.drawing.top = top
    chart.drawing.left = 10
    work_sheet.add_chart(chart)


if __name__ == "__main__":
    # Make sure we are in the common run path (this script should by in
    # (root)/utils/, run path should be (root)/)
    SCRIPT_DIR, _ = os.path.split(os.path.abspath(__file__))
    RUN_DIR, _ = os.path.split(SCRIPT_DIR)
    PATH_OUTPUT = os.path.join(
        SCRIPT_DIR.split('\\yarara\\yarara', 1)[0], 'output', 'results')
    PATH_LOGS = os.path.join(PATH_OUTPUT, "reports", "logs")
    os.environ.setdefault('OUTPUT', PATH_OUTPUT)
    os.environ.setdefault('https_proxy', 'http://proxy-us.intel.com:911')
    os.environ.setdefault('LOGS', PATH_LOGS)
    import platform
    if platform.system() == 'Windows':
        YARARA_PATH = os.path.join(SCRIPT_DIR, '..\\', '..\\')
        DEST_PATH = 'C://virtualenvs/yarara/Lib/site-packages/yarara'
    else:
        YARARA_PATH = os.path.join(SCRIPT_DIR, '../', '../')
        VIRTUAL_PATH = '/'.join(SCRIPT_DIR.split('/', 4)[0:3])
        VTS_YARARA = 'virtualenvs/yarara/lib/python2.7/site-packages/yarara'
        DEST_PATH = os.path.join(VIRTUAL_PATH, VTS_YARARA)
    import shutil

    shutil.rmtree(DEST_PATH)
    shutil.copytree(YARARA_PATH, DEST_PATH)
    # Call export function
    INFO_FILE = open(os.path.join(PATH_OUTPUT, 'results.json'), 'r')
    import json

    JSON_OUTPUT = INFO_FILE.read()
    JSON_OUTPUT = json.loads(JSON_OUTPUT)
    INFO_FILE.close()
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    generate_report(JSON_OUTPUT)
