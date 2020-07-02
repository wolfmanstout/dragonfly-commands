#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Tobii eye tracker wrapper."""

import sys

from dragonfly import (Mouse, Text, Window)
import _dragonfly_local as local


class Tracker(object):
    _instance = None

    @classmethod
    def get_connected_instance(cls):
        if not cls._instance:
            cls._instance = cls()
        if not cls._instance.is_connected:
            cls._instance.connect()
        return cls._instance

    def __init__(self):
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
            self.is_mock = False
        except:
            print("Eye tracking libraries are unavailable.")
            self.is_mock = True
        self._host = None
        self._gaze_point = None
        self._gaze_state = None
        self.is_connected = False

    def connect(self):
        if self.is_mock:
            return
        self._host = Host()
        gaze_points = self._host.Streams.CreateGazePointDataStream()
        action = Action[Double, Double, Double](self._handle_gaze_point)
        gaze_points.GazePoint(action)
        gaze_state = self._host.States.CreateGazeTrackingObserver()
        gaze_state.Changed += self._handle_gaze_state
        self.is_connected = True
        print("Eye tracker connected.")

    def disconnect(self):
        if not self.is_connected:
            return
        self._host.DisableConnection()
        self._host = None
        self._gaze_point = None
        self._gaze_state = None
        self.is_connected = False
        print("Eye tracker disconnected.")

    def _handle_gaze_point(self, x, y, timestamp):
        self._gaze_point = (x, y, timestamp)

    def _handle_gaze_state(self, sender, state):
        if not state.IsValid:
            print("Ignoring invalid gaze state.")
            return
        self._gaze_state = state.Value

    def has_gaze_point(self):
        return (not self.is_mock and
                self._gaze_state == GazeTracking.GazeTracked and
                self._gaze_point)

    def get_gaze_point_or_default(self):
        if self.has_gaze_point():
            return self._gaze_point[:2]
        else:
            window_position = Window.get_foreground().get_position()
            return (window_position.x_center, window_position.y_center)

    def print_gaze_point(self):
        if not self.has_gaze_point():
            print("No valid gaze point.")
            return
        print("Gaze point: (%f, %f)" % self._gaze_point[:2])

    def move_to_gaze_point(self, offset=(0, 0)):
        gaze = self.get_gaze_point_or_default()
        x = max(0, int(gaze[0]) + offset[0])
        y = max(0, int(gaze[1]) + offset[1])
        Mouse("[%d, %d]" % (x, y)).execute()

    def type_gaze_point(self, format):
        Text(format % self.get_gaze_point_or_default()).execute()
