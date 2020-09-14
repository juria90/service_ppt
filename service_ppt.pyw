#!python3
'''service_ppt.py has the main() starting function.
'''

import os
import sys
import traceback

sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'bible'))
import wx

from mainframe import Frame


def main():
    '''Main entrance function.
    '''
    redirect = False
    app = wx.App(redirect=redirect)   # Error messages go to popup window
    # root, _ext = os.path.splitext(os.path.abspath(__file__))
    # logfn = root + '.log'
    # app.RedirectStdio(logfn)
    frame = Frame()
    frame.Show()
    app.SetTopWindow(frame)
    app.MainLoop()

if __name__ == '__main__':
    try:
        main()
    except:
        e = sys.exc_info()[0]
        print("Error: %s" % e)
        traceback.print_exc()
