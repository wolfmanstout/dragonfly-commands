#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Eye tracker functions."""

from ctypes import (byref, c_double, CDLL)
import win32gui

from dragonfly import (Mouse, Text)
import _dragonfly_local as local

# Attempt to load eye tracker DLLs.
try:
    eyex_dll = CDLL(local.DLL_DIRECTORY + "/Tobii.EyeX.Client.dll")
    tracker_dll = CDLL(local.DLL_DIRECTORY + "/Tracker.dll")
except:
    print("Tracker not loaded.")


def connect():
    try:
        result = tracker_dll.connect()
        print("connect: %d" % result)
    except:
        print("Could not connect to tracker.")


def disconnect():
    result = tracker_dll.disconnect()
    print("disconnect: %d" % result)


def get_position():
    x = c_double()
    y = c_double()
    tracker_dll.last_position(byref(x), byref(y))
    return (x.value, y.value)


def screen_to_foreground(position):
    return win32gui.ScreenToClient(win32gui.GetForegroundWindow(), position);


def print_position():
    print("(%f, %f)" % get_position())


def move_to_position(offset=(0, 0)):
    position = get_position()
    x = max(0, int(position[0]) + offset[0])
    y = max(0, int(position[1]) + offset[1])
    Mouse("[%d, %d]" % (x, y)).execute()


def type_position(format):
    position = get_position()
    Text(format % (position[0], position[1])).execute()


def activate_position():
    tracker_dll.activate()


def panning_step_position():
    tracker_dll.panning_step()
