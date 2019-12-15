#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Eye tracker functions."""

import win32gui
import sys

from dragonfly import (Mouse, Text, Window)
import _dragonfly_local as local


class Tracker(object):

    def __init__(self):
        self.is_available = False
        self.host = None
        self.last_gaze_point = None
        self.gaze_state = None

    def connect(self):
        if not self.is_available:
            # Attempt to load eye tracker DLLs.
            global clr, Action, Double, Host, GazeTracking
            try:
                import clr
                from System import Action, Double
                sys.path.append(local.DLL_DIRECTORY)
                clr.AddReference("Tobii.Interaction.Model")
                clr.AddReference("Tobii.Interaction.Net")
                from Tobii.Interaction import Host
                from Tobii.Interaction.Framework import GazeTracking
                self.is_available = True
            except:
                print("Tracker not available.")
                return False

        if self.host:
            print("Tracker already connected.")
            return True
        self.host = Host()
        gaze_points = self.host.Streams.CreateGazePointDataStream()
        action = Action[Double, Double, Double](self.handle_gaze_point)
        gaze_points.GazePoint(action)
        gaze_state = self.host.States.CreateGazeTrackingObserver()
        gaze_state.Changed += self.handle_gaze_state
        print("Tracker connected.")
        return True

    def disconnect(self):
        self.host.DisableConnection()
        self.host = None
        print("Tracker disconnected.")

    def handle_gaze_point(self, x, y, timestamp):
        self.last_gaze_point = (x, y, timestamp)

    def handle_gaze_state(self, sender, state):
        if not state.IsValid:
            print("Invalid gaze state")
            return
        self.gaze_state = state.Value

    def has_gaze_point(self):
        return self.gaze_state == GazeTracking.GazeTracked and self.last_gaze_point

    def get_gaze_point_or_default(self):
        if self.has_gaze_point():
            return self.last_gaze_point[:2]
        else:
            window_position = Window.get_foreground().get_position()
            return (window_position.x_center, window_position.y_center)

    def screen_to_foreground(self, position):
        return win32gui.ScreenToClient(win32gui.GetForegroundWindow(), position)

    def print_position(self):
        if not self.has_gaze_point():
            return False
        print("(%f, %f)" % self.last_gaze_point[:2])
        return True

    def move_to_position(self, offset=(0, 0)):
        gaze = self.get_gaze_point_or_default()
        x = max(0, int(gaze[0]) + offset[0])
        y = max(0, int(gaze[1]) + offset[1])
        print("Moving to last gaze: {}".format(gaze))
        Mouse("[%d, %d]" % (x, y)).execute()
        return True

    def type_position(self, format):
        Text(format % self.get_gaze_point_or_default()).execute()


tracker = Tracker()


def get_tracker():
    return tracker
