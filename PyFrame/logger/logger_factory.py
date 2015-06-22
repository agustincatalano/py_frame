import logging
from os import environ
import os


LOG_FILENAME = os.path.join(environ.get('OUTPUT'), 'logs', 'py_frame.log')
my_logger = None


def __initialize():
    global my_logger
    my_logger = logging.getLogger('py_frame')


    # create handlers for loggers
    console_handler = logging.StreamHandler()
    handler = logging.FileHandler(LOG_FILENAME, mode='w')

    # set logging levels
    console_handler.setLevel(logging.DEBUG)
    my_logger.setLevel(logging.DEBUG)

    # define format for logger
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # set format to handlers
    handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # add handlers to logger
    my_logger.addHandler(handler)
    my_logger.addHandler(console_handler)


def get_logger():
    if my_logger is not None:
        return my_logger
    else:
        __initialize()
        return my_logger