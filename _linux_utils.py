#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Functions and classes to help with manipulating a remote Linux instance."""

import xmlrpclib
import sys
import time

from dragonfly import *

class LinuxHelper(object):
    def __init__(self):
        self.server = xmlrpclib.ServerProxy("http://127.0.0.1:12400", allow_none=True)
        self.last_update = None

    def GetActiveWindowTitle(self):
        now = time.clock()
        if (not self.last_update) or now - self.last_update > 0.05:
            try:
                self.remote_title = self.server.GetActiveWindowTitle()
                duration = time.clock() - now
                if duration > 0.05:
                    print("Slow duration: %f" % duration)
            except:
                self.remote_title = ""
            self.last_update = now
        return self.remote_title

    def ActivateWindow(self, title):
        try:
            self.server.ActivateWindow(title)
        except:
            pass

linux_helper = LinuxHelper()

class UniversalAppContext(AppContext):
    def matches(self, executable, title, handle):
        if AppContext.matches(self, executable, title, handle):
            return True
        # Only check Linux if it is active.
        if title.find("Oracle VM VirtualBox") != -1 or title.find(" - Chrome Remote Desktop") != -1:
            remote_title = linux_helper.GetActiveWindowTitle().lower()
            found = remote_title.find(self._title) != -1
            if self._exclude != found:
                return True
        return False

class ActivateLinuxWindow(ActionBase):
    """Activate the provided window within Linux."""

    def __init__(self, title):
        super(ActivateLinuxWindow, self).__init__()
        self.title = title

    def _execute(self, data=None):
        linux_helper.ActivateWindow(self.title)
