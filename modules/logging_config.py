import logging
import sys

DEFAULT_LOG_FORMAT = "%(asctime)s - [%(levelname)s] [%(threadName)s] - %(name)s - %(message)s"


def setup_logging(log_level=logging.INFO, log_file=None):
    """Setup centralized logging for entire application."""
    formatter = logging.Formatter(DEFAULT_LOG_FORMAT)
    handlers = [logging.StreamHandler(sys.stdout)]

    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))

    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(log_level)

    for handler in handlers:
        handler.setFormatter(formatter)
        root_logger.addHandler(handler)

    return root_logger
