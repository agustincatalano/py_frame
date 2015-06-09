# -*- coding: utf-8 -*-
from service import rest_manager

url = 'http://www.mocky.io/v2/5567723d11ea660515cf131c'

response = rest_manager.get(url)
assert rest_manager.validate_status_in_response(response, rest_manager.HttpStatusCodes.OK), \
    rest_manager.response_info(response)
assert response.json() == {"pyframe": "test"}
#assert response.json() == {"pyframe": "FAIL"}, rest_manager.response_info(response)

put_url = 'http://www.mocky.io/v2/5185415ba171ea3a00704eed'
response_put = rest_manager.put(put_url)
assert rest_manager.validate_status_in_response(response, rest_manager.HttpStatusCodes.OK), \
    rest_manager.response_info(response)