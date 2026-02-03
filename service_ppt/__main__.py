#!python3
"""Main application entry point.

This module contains the main() function that initializes the wxPython
application and displays the main window for the service_ppt application.
"""

import sys
import traceback
from pathlib import Path

# Ensure the parent directory is in the path so we can import service_ppt
# This is needed when running as a module: python -m service_ppt
_package_dir = Path(__file__).parent.parent
if str(_package_dir) not in sys.path:
    sys.path.insert(0, str(_package_dir))

import wx

from service_ppt.mainframe import Frame


def activate_macos_app():
    """Activate macOS application using native APIs to bring window to front."""
    try:
        # Try using pyobjc if available (comes with py-applescript dependency)
        from AppKit import NSApplication
        _macos_app = NSApplication.sharedApplication()
        _macos_app.activateIgnoringOtherApps_(True)
    except ImportError:
        # Fallback: Use ctypes to call NSApplication directly
        try:
            import ctypes
            from ctypes import c_void_p, c_bool
            from ctypes.util import find_library

            objc = ctypes.cdll.LoadLibrary(find_library("objc"))

            # Define Objective-C runtime functions
            objc.objc_getClass.restype = c_void_p
            objc.sel_registerName.restype = c_void_p
            objc.objc_msgSend.restype = c_void_p
            objc.objc_msgSend.argtypes = [c_void_p, c_void_p]

            NSApplication = objc.objc_getClass(b"NSApplication")
            sharedApplication = objc.sel_registerName(b"sharedApplication")
            app_obj = objc.objc_msgSend(NSApplication, sharedApplication)
            activate = objc.sel_registerName(b"activateIgnoringOtherApps:")
            objc.objc_msgSend(app_obj, activate, c_bool(True))
        except Exception:
            # If all else fails, silently continue
            pass


def main():
    """Main entrance function."""
    redirect = False
    app = wx.App(redirect=redirect)  # Error messages go to popup window
    # root, _ext = os.path.splitext(os.path.abspath(__file__))
    # logfn = root + '.log'
    # app.RedirectStdio(logfn)
    frame = Frame()
    app.SetTopWindow(frame)
    
    # macOS-specific: Use native APIs to activate the application
    if sys.platform == "darwin":
        activate_macos_app()
    
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
        traceback.print_exc()
        sys.exit(1)
