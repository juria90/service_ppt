"""Process existence checking utility.

This module provides cross-platform functionality to check if a process is
currently running by name.
"""

import subprocess
import sys


def process_exists(process_name: str) -> bool:
    """Check if a process is currently running by name.

    :param process_name: Name of the process to check
    :returns: True if process is running, False otherwise
    """
    if sys.platform.startswith("win32"):
        # https://stackoverflow.com/questions/7787120/python-check-if-a-process-is-running-or-not

        # Use tasklist to reduce package dependency.

        call = "TASKLIST", "/FI", f"imagename eq {process_name}"
        # use buildin check_output right away
        output = subprocess.check_output(call).decode()
        # check in last line for process name
        last_line = output.strip().split("\r\n")[-1]
        # because Fail message could be translated
        return last_line.lower().startswith(process_name.lower())

    call = f"""pgrep "{process_name}" """
    retcode = subprocess.call(call)
    return retcode == 0
