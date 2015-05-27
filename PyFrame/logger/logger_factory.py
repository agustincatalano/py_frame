import logging

LOG_FILENAME = '..\logger\py_frame.log'
my_logger = None


def initialize():
    global my_logger
    my_logger = logging.getLogger('py_frame')
    my_logger.setLevel(logging.DEBUG)
    # Add the log message handler to the logger
    handler = logging.FileHandler(LOG_FILENAME, mode='w')
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    my_logger.addHandler(handler)


def get_logger():
    if my_logger is not None:
        return my_logger
    else:
        initialize()
        return my_logger