__author__ = "Ofner Mario"

import logging


__all__ = ( 'set_rootLogger_log_level' )

# --------------------------------------------------------------------------------
# - Initializations
# --------------------------------------------------------------------------------

# default Logging Environment
# use the "root"-Logger Definition to get Log-Messages from Libraries
# default Format String for LogMessages
formatter = logging.Formatter('%(asctime)s - %(levelname)-.1s: - %(message)s - [%(name)s:%(lineno)s]')

# default handler for stdout
console = logging.StreamHandler()
console.setFormatter(formatter)
logging.getLogger().addHandler(console)

# Set RootLogger to "Debug"
logging.getLogger().setLevel(logging.DEBUG)

# Set "Console" to whatever you want to see
# Use "set_rootLogger_log_level()" in your Application to override the default_level
console.setLevel(logging.INFO)


def set_rootLogger_log_level(_level=logging.DEBUG):
    console.setLevel(_level)
