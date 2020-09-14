# https://stackoverflow.com/questions/7787120/python-check-if-a-process-is-running-or-not

import subprocess
import sys


def process_exists(process_name):
    if sys.platform.startswith('win32'):
        # Use tasklist to reduce package dependency.

        call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
        # use buildin check_output right away
        output = subprocess.check_output(call).decode()
        # check in last line for process name
        last_line = output.strip().split('\r\n')[-1]
        # because Fail message could be translated
        return last_line.lower().startswith(process_name.lower())
    else:
        raise NotImplementedError('process_exists is not implemented.')
