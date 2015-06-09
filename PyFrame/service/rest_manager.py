# -*- coding: utf-8 -*-
import requests
import json

try:
    from requests import *
except ImportError, e:
    print 'Unable to load requests:', e

REQUEST_INFO = ['url', 'method', 'headers', 'body']
RESPONSE_INFO = ['status_code', 'reason', 'content']
SEPARATOR = '\n%s\n' % ('-'*80)


def post(url, data=None, headers=None, **kwargs):
    return requests.post(url, data=parse_data_as_json(data), headers=headers or get_headers(), **kwargs)


def get(url, **kwargs):
    return requests.get(url, **kwargs)


def put(url, data=None, **kwargs):
    return requests.put(url, data=parse_data_as_json(data), **kwargs)


def delete(url, **kwargs):
    return requests.delete(url, **kwargs)


def parse_data_as_json(data):
    if data is None:
        return None
    elif isinstance(data, dict):
            return json.dumps(data)
    else:
        return str(data)


def response_info(response):
    req_info = '\n'.join(['%s: %s' % (attr.upper(), getattr(response.request, attr)) for
                          attr in REQUEST_INFO if getattr(response.request, attr, None)])

    res_info = '\n'.join(['%s: %s' % (attr.upper(), getattr(response, attr)) for
                          attr in RESPONSE_INFO if getattr(response, attr, None)])

    return SEPARATOR.join(['\n', 'REQUEST_INFO:', req_info, 'RESPONSE_INFO:', res_info])


def get_json_headers():
    return get_headers('data')


def get_headers(rest="data"):
    header = {'Content-type': 'application/json'}
    if rest == "data":
        header['Accept'] = 'application/json'
    return header


class HttpStatusCodes():
    OK = 200
    CREATED = 201
    ACCEPTED = 202
    REQUEST_ACCEPTED = 204  # NO_CONTENT
    BAD_REQUEST = 400
    UNAUTHORIZED = 401
    NOT_FOUND = 404
    METHOD_NOT_ALLOWED = 405
    ERROR = 500  # INTERNAL_SERVER_ERROR


def validate_status_in_response(response, status_code):
    logging.debug("Status:  " + str(response.status_code))
    return response.status_code == int(status_code)