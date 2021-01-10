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

from collections import OrderedDict
import copy
import json
import os
import os.path
import platform
from six import text_type
import tempfile

from dragonfly import (
    ActionBase,
    DynStrActionBase,
    Function,
    Grammar,
    Key,
    MappingRule,
    Pause,
    Repetition,
    Sequence,
    StartApp,
    Text,
    WaitWindow,
)
from dragonfly.windows.window import Window

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

class Override(object):
    """Used as a dictionary key to override with combine_maps_checked."""

    def __init__(self, key):
        self.key = key

    def __hash__(self):
        return hash(self.key)

    def __str__(self):
        return "Override(" + self.key + ")"


class Delete(object):
    """Used as a dictionary key to delete with combine_maps_checked."""

    def __init__(self, key):
        self.key = key

    def __hash__(self):
        return hash(self.key)

    def __str__(self):
        return "Delete(" + self.key + ")"


def combine_maps(*maps):
    """Merge the contents of multiple maps.

    Does not allow deletions or overrides. Skips empty maps.
    """
    # Use OrderedDict to maintain possible ordering in the source maps.
    result = OrderedDict()
    for map in maps:
        if not map:
            continue
        for key, value in map.items():
            if key in result and result[key] != value:
                raise ValueError("Key already exists: {}. Use combine_maps_checked to override.".format(key))
            result[key] = value
    return result


def combine_maps_checked(*maps):
    """Merge the contents of multiple maps allowing deletion and override.

    Wrap keys in Override and Delete to perform those operations. Skips empty
    maps.
    """
    # Use OrderedDict to maintain possible ordering in the source maps.
    result = OrderedDict()
    for map in maps:
        if not map:
            continue
        for key, value in map.items():
            if isinstance(key, Delete):
                if value is not None:
                    raise ValueError("Delete key has non-None value: {}".format(key))
                if key.key not in result:
                    raise ValueError("Delete key cannot be applied to missing key: {}".format(key))
                del result[key.key]
            elif isinstance(key, Override):
                if key.key not in result:
                    raise ValueError("Override key cannot be applied to missing key: {}".format(key))
                result[key.key] = value
            else:
                if key in result:
                    raise ValueError("Key already exists: {}. Wrap key in Delete or Override.".format(key))
                result[key] = value
    return result


def text_map_to_action_map(text_map):
    """Converts string values in a map to text actions."""
    return dict((k, Text(v.replace("%", "%%")))
                for (k, v) in text_map.items())


def _printable_to_key_action_spec(printable):
    if len(printable) != 1:
        raise ValueError("Printable must have a single character: %s" % printable)
    if printable == "/":
        return "slash"
    if printable == ":":
        return "colon"
    if printable == ",":
        return "comma"
    if printable == "-":
        return "minus"
    if printable == "%":
        return "%%"
    return printable


def text_map_to_key_action_map(text_map):
    """Converts string values in a map to key actions."""
    return dict((k, Key(_printable_to_key_action_spec(v)))
                for (k, v) in text_map.items())


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
        return self.delimiter.join(text_type(v)
                                   for v in Sequence.value(self, node)
                                   if v)


def renamed_element(name, element):
    element_copy = copy.copy(element)
    element_copy.name = name
    return element_copy

def element_map_to_extras(element_map):
    """Converts an element map to a standard named element list that may be used in
    MappingRule.
    """
    return [renamed_element(name, element[0] if isinstance(element, tuple) else element)
            for (name, element) in element_map.items()]


def element_map_to_defaults(element_map):
    """Converts an element map to a map of element names to default values."""
    return dict([(name, element[1])
                 for (name, element) in element_map.items()
                 if isinstance(element, tuple)])


def create_rule(name, action_map, element_map=None, exported=False, context=None):
    """Creates a rule with the given name, binding the given element map to the
    action map.
    """
    element_map = element_map if element_map else {}
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


class ModifiedAction(ActionBase):
    def __init__(self, name, action):
        ActionBase.__init__(self)
        self.name = name
        self.action = action

    def _execute(self, data=None):
        modifier = data[self.name]
        modified_action = modifier(self.action)
        modified_action.execute(data)


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
                Key("tab:%d/10, enter" % (repeat - 1)).execute()
            else:
                Key("right:%d/10, enter" % repeat).execute()
        else:
            Key("alt:down, tab:%d/10, alt:up" % repeat).execute()


class RunApp(ActionBase):
    """Starts an app and waits for it to be the foreground app."""

    def __init__(self, *args):
        super(RunApp, self).__init__()
        self.args = args

    def _execute(self, data=None):
        StartApp(*self.args).execute()
        WaitWindow(None, os.path.basename(self.args[0]), 3).execute()


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
        foreground = Window.get_foreground()
        if foreground.title.find("Emacs editor") != -1:
            Key("c-y").execute()
        else:
            Key("c-v").execute()


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
    return FormattedText(spec, lambda text: text[0].upper() + text[1:])


def load_json(filename):
    try:
        with open(os.path.join(os.path.dirname(os.path.realpath(__file__)), filename)) as json_file:
            return json.load(json_file)
    except IOError:
        print(filename + " not found")
        return None


class GrammarController(object):
    """Wraps grammars so they can be turned on and off by command."""

    def __init__(self, name, grammars):
        self._controlled_grammars = grammars
        self.enabled = True
        rule = create_rule(name + "_mode",
                           {
                               name + " (off|close)": Function(lambda: self.disable()),
                               name + " (on|open)": Function(lambda: self.enable()),
                           },
                           exported=True)
        self._command_grammar = Grammar(name + "_mode")
        self._command_grammar.add_rule(rule)

    def enable(self):
        if not self.enabled:
            for grammar in self._controlled_grammars:
                grammar.enable()
        self.enabled = True

    def disable(self):
        if self.enabled:
            for grammar in self._controlled_grammars:
                grammar.disable()
        self.enabled = False

    def load(self):
        for grammar in self._controlled_grammars:
            grammar.load()
        self._command_grammar.load()

    def unload(self):
        for grammar in self._controlled_grammars:
            grammar.unload()
        self._command_grammar.unload()
