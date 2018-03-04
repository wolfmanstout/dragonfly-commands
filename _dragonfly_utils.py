#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Various utility functions for working with dragonfly.

It requires the presence of dragonfly_local.py in the core subdirectory to define
the following variables:

HOME: Path to your home directory.
DLL_DIRECTORY: Path to directory containing DLLs used in this module. Missing
    DLLs will be ignored.
CHROME_DRIVER_PATH: Path to chrome driver executable.
"""

import json
import os
import os.path
import platform
import tempfile

from dragonfly import (
    ActionBase,
    DynStrActionBase,
    MappingRule,
    Pause,
    Repetition,
    Sequence,
#    StartApp,
#    WaitWindow,
)
from aenea.lax import (
    Key,
    Text,
)
# from dragonfly.windows.window import Window

import _dragonfly_local as local

#-------------------------------------------------------------------------------
# Utility functions and classes for manipulating grammars and their components.
# To make creating rules a bit easier, we define a few new types:
# Action map: Mapping from command spec (string) to action.
# Element map: Mapping from element name to either element or tuple of element
#              and default value. The preset element names will be overwritten
#              when used.
#
# The advantage to this separation is we can easily bind a single element to
# different names to be used in multiple rules, and we can easily create rules
# on the fly without defining a new class.

def combine_maps(*maps):
    """Merge the contents of multiple maps, giving precedence to later maps."""
    result = {}
    for map in maps:
        if map:
            result.update(map)
    return result


def text_map_to_action_map(text_map):
    """Converts string values in a map to text actions."""
    return dict((k, Text(v.replace("%", "%%")))
                for (k, v) in text_map.iteritems())


class JoinedRepetition(Repetition):
    """Like Repetition, except the results are joined with the given delimiter
    instead of returned as a list.
    """

    def __init__(self, delimiter, *args, **kwargs):
        Repetition.__init__(self, *args, **kwargs)
        self.delimiter = delimiter

    def value(self, node):
        return self.delimiter.join(Repetition.value(self, node))

class JoinedSequence(Sequence):
    """Like Sequence, except the results are joined with the given delimiter instead
    of returned as a list.
    """

    def __init__(self, delimiter, *args, **kwargs):
        Sequence.__init__(self, *args, **kwargs)
        self.delimiter = delimiter

    def value(self, node):
        return self.delimiter.join(str(v)
                                   for v in Sequence.value(self, node)
                                   if v)


class ElementWrapper(Sequence):
    """Identity function on element, useful for renaming."""

    def __init__(self, name, child):
        Sequence.__init__(self, (child, ), name)

    def value(self, node):
        return Sequence.value(self, node)[0]


def element_map_to_extras(element_map):
    """Converts an element map to a standard named element list that may be used in
    MappingRule.
    """
    return [ElementWrapper(name, element[0] if isinstance(element, tuple) else element)
            for (name, element) in element_map.items()]


def element_map_to_defaults(element_map):
    """Converts an element map to a map of element names to default values."""
    return dict([(name, element[1])
                 for (name, element) in element_map.items()
                 if isinstance(element, tuple)])


def create_rule(name, action_map, element_map, exported=False, context=None):
    """Creates a rule with the given name, binding the given element map to the
    action map.
    """
    return MappingRule(name,
                       action_map,
                       element_map_to_extras(element_map),
                       element_map_to_defaults(element_map),
                       exported,
                       context=context)


def combine_contexts(context1, context2):
    """Combine two contexts using "&", treating None as equivalent to a context that
    matches everything.
    """
    if not context1:
        return context2
    if not context2:
        return context1
    return context1 & context2


class SwitchWindows(DynStrActionBase):
    """Simulates the effects of alt-tab. The constructor argument should be a string
    representing the number of times to effectively press the "tab" button if
    alt-tab were actually being used.
    """

    def _parse_spec(self, spec):
        return int(spec)

    def _execute_events(self, repeat):
        if platform.release() >= "8":
            # Work around security restrictions in Windows 8.
            # Credit: https://autohotkey.com/board/topic/84771-alttab-mapping-isnt-working-anymore-in-windows-8/
            os.startfile("C:/Users/Default/AppData/Roaming/Microsoft/Internet Explorer/Quick Launch/Window Switcher.lnk")
            Pause("10").execute()
            if platform.release() == "8":
                Key("tab:%d/25, enter" % (repeat - 1)).execute()
            else:
                Key("right:%d/25, enter" % repeat).execute()
        else:
            Key("alt:down, tab:%d/25, alt:up" % repeat).execute()


class RunApp(ActionBase):
    """Starts an app and waits for it to be the foreground app."""

    def __init__(self, *args):
        super(RunApp, self).__init__()
        self.args = args

    def _execute(self, data=None):
        pass
        #StartApp(*self.args).execute()
        #WaitWindow(None, os.path.basename(self.args[0]), 3).execute()


class RunEmacs(ActionBase):
    """Runs Emacs on a temporary file with the given suffix."""

    def __init__(self, suffix):
        super(RunEmacs, self).__init__()
        self.suffix = suffix

    def _execute(self, data=None):
        cygwin_tmp = local.CYGWIN_ROOT + "/tmp"
        (f, path) = tempfile.mkstemp(prefix="emacs", suffix=self.suffix, dir=cygwin_tmp)
        os.close(f)
        os.unlink(path)
        cygwin_path = path[path.index("tmp") - 1:].replace("\\", "/")
        emacs_path = local.CYGWIN_ROOT + "/bin/emacsclient-w32.exe"
        RunApp(emacs_path, "-c", "-n", cygwin_path).execute()


class UniversalPaste(ActionBase):
    """Paste action that works everywhere, including Emacs."""

    def _execute(self, data=None):
        Key("c-v").execute()
        # foreground = Window.get_foreground()
        # if foreground.title.find("Emacs editor") != -1:
        #     Key("c-y").execute()
        # else:
        #     Key("c-v").execute()


class FormattedText(DynStrActionBase):
    """Types text after running through formatter function."""

    def __init__(self, spec, formatter):
        DynStrActionBase.__init__(self, spec)
        self.formatter = formatter

    def _parse_spec(self, spec):
        return spec

    def _execute_events(self, events):
        Text(self.formatter(events)).execute()


def lowercase_text_action(spec):
    return FormattedText(spec, lambda text: text.lower())


def uncapitalize_text_action(spec):
    return FormattedText(spec, lambda text: text[0].lower() + text[1:])


def capitalize_text_action(spec):
    return FormattedText(spec, lambda text: text.capitalize())


def byteify(input):
    """Convert unicode to str. Dragonfly grammars don't play well with Unicode."""
    if isinstance(input, dict):
        return dict((byteify(key), byteify(value)) for key,value in input.iteritems())
    elif isinstance(input, list):
        return [byteify(element) for element in input]
    elif isinstance(input, unicode):
        return input.encode('utf-8')
    else:
        return input


def load_json(filename):
    try:
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)), filename)) as json_file:
            return byteify(json.load(json_file))
    except IOError:
        print filename + " not found"
        return None
