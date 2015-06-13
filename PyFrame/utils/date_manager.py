from datetime import datetime
import re

SEPARATOR_PATTERN = '[/-]'
DATE_PATTERN = '(\d+%s\d+%s\d+)' % (SEPARATOR_PATTERN, SEPARATOR_PATTERN)


DEFAULT_FORMAT = '%m/%d/%Y'
DAY_MONTH_YEAR_FORMAT = '%d/%m/%Y'
YEAR_MONTH_DAY_FORMAT = '%Y/%m/%d'


def get_iso_format_date(date_as_string, _format=DEFAULT_FORMAT):
    match = re.match(DATE_PATTERN, date_as_string)
    assert match, 'Invalid date'
    date_formatted = re.sub(SEPARATOR_PATTERN, '/', date_as_string)
    return datetime.strptime(date_formatted, _format).date().isoformat()

#TEST
# print get_iso_format_date('04-27-1988')
# print get_iso_format_date('04/27/5465')
# print get_iso_format_date('30/05/1945', DAY_MONTH_YEAR_FORMAT)
# print get_iso_format_date('1586/05/12', YEAR_MONTH_DAY_FORMAT)
# print get_iso_format_date('30-05-4566', DAY_MONTH_YEAR_FORMAT)
# print get_iso_format_date('7865-05-12', YEAR_MONTH_DAY_FORMAT)
