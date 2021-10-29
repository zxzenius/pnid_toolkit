import logging


def init():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    log_handler = logging.StreamHandler(stream=None)
    logger.addHandler(log_handler)
