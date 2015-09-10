"""PyWrap is a command line utility that runs a python
program with additional runtime tools, such as PyCrust."""

__author__ = "Patrick K. O'Brien  pobrien@orbtech.com "
__cvsid__ = "$Id: PyCrust.txt,v 1.15 2005/03/29 23:39:27 robind Exp $"
__revision__ = "$Revision: 1.15 $"[11:-2]

import os
import sys
import wx
from wx.py.crust import CrustFrame

def wrap(app):
    wx.InitAllImageHandlers()
    frame = CrustFrame()
    frame.SetSize((750, 525))
    frame.Show(True)
    frame.shell.interp.locals['app'] = app
    app.MainLoop()

def main(modulename=None):
    sys.path.insert(0, os.curdir)
    if not modulename:
        if len(sys.argv) < 2:
            print "Please specify a module name."
            raise SystemExit
        modulename = sys.argv[1]
        if modulename.endswith('.py'):
            modulename = modulename[:-3]
    module = __import__(modulename)
    # Find the App class.
    App = None
    d = module.__dict__
    for item in d.keys():
        try:
            if issubclass(d[item], wx.App):
                App = d[item]
        except (NameError, TypeError):
            pass
    if App is None:
        print "No App class was found."
        raise SystemExit
    app = App()
    wrap(app)
if __name__ == '__main__':
    main()
