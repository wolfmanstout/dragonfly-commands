#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
This contains all commands which may be spoken continuously or repeated. It
requires the presence of dragonfly_local.py in the core subdirectory to define
the following variables:

HOME: Path to your home directory.
DLL_DIRECTORY: Path to directory containing DLLs used in this module. Missing
    DLLs will be ignored.
CHROME_DRIVER_PATH: Path to chrome driver executable.

This is heavily modified from _multiedit.py, found here:
https://code.google.com/p/dragonfly-modules/

"""

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r99")
except ImportError:
    pass

from dragonfly import *
import dragonfly.log

import dragonfly_words
from dragonfly_local import *

import BaseHTTPServer
import Queue
import SocketServer
import cProfile
import json
import socket
import threading
import time
import urllib
import urllib2
import win32gui

# Make sure dragonfly errors show up in NatLink messages.
dragonfly.log.setup_log()

# Load _repeat.txt.
config = Config("repeat")
namespace = config.load()

#-------------------------------------------------------------------------------
# Eye tracker functions.
# TODO: Move these into a separate module.

# Attempt to load eye tracker DLLs.
from ctypes import *
try:
    eyex_dll = CDLL(DLL_DIRECTORY + "/Tobii.EyeX.Client.dll")
    tracker_dll = CDLL(DLL_DIRECTORY + "/Tracker.dll")
except:
    print("Tracker not loaded.")

def connect():
    result = tracker_dll.connect()
    print("connect: %d" % result)

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

def move_to_position():
    position = get_position()
    Mouse("[%d, %d]" % (max(0, int(position[0])), max(0, int(position[1])))).execute()

def type_position(format):
    position = get_position()
    Text(format % (position[0], position[1])).execute()

def activate_position():
    tracker_dll.activate()

#-------------------------------------------------------------------------------
# Actions for manipulating Chrome via WebDriver.
# TODO: Move to a separate module.

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

def create_driver():
    global driver
    chrome_options = Options()
    chrome_options.experimental_options["debuggerAddress"] = "127.0.0.1:9222"
    driver = webdriver.Chrome(CHROME_DRIVER_PATH, chrome_options=chrome_options)

def quit_driver():
    global driver
    if driver:
        driver.quit()
    driver = None

def switch_to_active_tab():
    tabs = json.load(urllib2.urlopen("http://127.0.0.1:9222/json"))
    # Chrome seems to order the tabs by when they were last updated, so we find
    # the first one that is not an extension.
    for tab in tabs:
        if not tab["url"].startswith("chrome-extension://"):
            active_tab = tab["id"]
            break
    for window in driver.window_handles:
        # ChromeDriver adds to the raw ID, so we just look for substring match.
        if active_tab in window:
            driver.switch_to_window(window);
            print "Switched to: " + driver.title.encode('ascii', 'backslashreplace')
            return

def test_driver():
    switch_to_active_tab()
    driver.get('http://www.google.com/xhtml');

class ClickElementAction(ActionBase):
    def __init__(self, xpath=None):
        ActionBase.__init__(self)
        self.xpath = xpath

    def _execute(self, data=None):
        switch_to_active_tab()
        if self.xpath:
            element = driver.find_element_by_xpath(self.xpath)
        element.click()


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
        result.update(map)
    return result

def text_map_to_action_map(text_map):
    """Converts string values in a map to text actions."""
    return dict((k, Text(v.replace("%", "%%"))) for (k, v) in text_map.iteritems())

class JoinedRepetition(Repetition):
    """Accepts multiple repetitions of the given element, and joins them with the
    given delimiter. See Repetition class for available arguments."""
    def __init__(self, delimiter, *args, **kwargs):
        Repetition.__init__(self, *args, **kwargs)
        self.delimiter = delimiter

    def value(self, node):
        return self.delimiter.join(Repetition.value(self, node))

class ElementWrapper(Sequence):
    """Identity function on element, useful for renaming."""

    def __init__(self, name, child): 
        Sequence.__init__(self, (child, ), name)

    def value(self, node):
        return Sequence.value(self, node)[0]

def element_map_to_extras(element_map):
    """Converts an element map to a standard named element list that may be used in
    MappingRule."""
    return [ElementWrapper(name, element[0] if isinstance(element, tuple) else element)
            for (name, element) in element_map.items()]

def element_map_to_defaults(element_map):
    """Converts an element map to a map of element names to default values."""
    return dict([(name, element[1])
                 for (name, element) in element_map.items()
                 if isinstance(element, tuple)])

def create_rule(name, action_map, element_map, exported=False):
    """Creates a rule with the given name, binding the given element map to the action map."""
    return MappingRule(name,
                       action_map,
                       element_map_to_extras(element_map),
                       element_map_to_defaults(element_map),
                       exported)

def combine_contexts(context1, context2):
    """Combine two contexts using "&", treating None as equivalent to a context that
    matches everything."""
    if not context1:
        return context2
    if not context2:
        return context1
    return context1 & context2


#-------------------------------------------------------------------------------
# Common maps and lists.
symbol_map = {
    "plus": " + ",
    "dub plus": "++",
    "minus": " - ",
    "nad": ", ",
    "coal": ": ",
    "equals": " = ",
    "dub equals": " == ",
    "not equals": " != ",
    "increment by": " += ",
    "greater than": " > ",
    "less than": " < ",
    "greater equals": " >= ",
    "less equals": " <= ", 
    "dot": ".",
    "leap": "(",
    "reap": ")",
    "lake": "{",
    "rake": "}",
    "lobe": "[",
    "robe": "]",
    "luke": "<",
    "dub luke": " << ", 
    "ruke": ">",
    "quote": "\"",
    "dash": "-",
    "semi": ";",
    "bang": "!",
    "percent": "%",
    "star": "*",
    "backslash": "\\", 
    "slash": "/",
    "tilde": "~", 
    "underscore": "_",
    "sick quote": "'",
    "dollar": "$",
    "carrot": "^",
    "arrow": "->",
    "fat arrow": "=>",
    "dub coal": "::",
    "amper": "&",
    "dub amper": " && ",
    "pipe": "|",
    "dub pipe": " || ",
    "hash": "#",
    "at symbol": "@",
    "question": "?",
}

numbers_map = {
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
    "point": ".",
    "minus": "-",
    "slash": "/",
    "coal": ":",
    "nad": ",",
}

short_letters_map = {
    "A": "a",
    "B": "b",
    "C": "c",
    "D": "d",
    "E": "e",
    "F": "f",
    "G": "g",
    "H": "h",
    "I": "i",
    "J": "j",
    "K": "k",
    "L": "l",
    "M": "m",
    "N": "n",
    "O": "o",
    "P": "p",
    "Q": "q",
    "R": "r",
    "S": "s",
    "T": "t",
    "U": "u",
    "V": "v",
    "W": "w",
    "X": "x",
    "Y": "y",
    "Z": "z",
}

long_letters_map = {
    "alpha": "a",
    "bravo": "b",
    "charlie": "c",
    "delta": "d",
    "echo": "e",
    "foxtrot": "f",
    "golf": "g",
    "hotel": "h",
    "india": "i",
    "juliet": "j",
    "kilo": "k",
    "lima": "l",
    "mike": "m",
    "november": "n",
    "oscar": "o",
    "poppa": "p",
    "quebec": "q",
    "romeo": "r",
    "sierra": "s",
    "tango": "t",
    "uniform": "u",
    "victor": "v",
    "whiskey": "w",
    "x-ray": "x",
    "yankee": "y",
    "zulu": "z",
    "dot": ".",
}

letters_map = combine_maps(short_letters_map, long_letters_map)

char_map = dict((k, v.strip()) for (k, v) in combine_maps(letters_map, numbers_map, symbol_map).iteritems())

# Load commonly misrecognized words saved to a file.
saved_words = []
try:
    with open(dragonfly_words.WORDS_PATH) as file:
        for line in file:
            word = line.strip()
            if len(word) > 2 and word not in letters_map:
                saved_words.append(line.strip())
except:
    print("Unable to open: " + dragonfly_words.WORDS_PATH)

#-------------------------------------------------------------------------------
# Action maps to be used in rules.

# Key actions which may be used anywhere in any command.
global_key_action_map = {
    "slap [<n>]": Key("enter/5:%(n)d"),
    "pat [<n>]": Key("space/5:%(n)d"),
    "tab [<n>]": Key("tab/5:%(n)d"),
}

# Actions of commonly used text navigation and mousing commands. These can be
# used anywhere except after commands which include arbitrary dictation.
release = Key("shift:up, ctrl:up")
key_action_map = {
    "up [<n>]":                         Key("up/5:%(n)d"),
    "down [<n>]":                       Key("down/5:%(n)d"),
    "left [<n>]":                       Key("left/5:%(n)d"),
    "right [<n>]":                      Key("right/5:%(n)d"),
    "fomble [<n>]": Key("c-right/5:%(n)d"),
    "bamble [<n>]": Key("c-left/5:%(n)d"),
    "dumbble [<n>]": Key("c-backspace/5:%(n)d"),
    "kimble [<n>]": Key("c-del/5:%(n)d"),
    "dird [<n>]": Key("a-backspace/5:%(n)d"),
    "kill [<n>]": Key("c-k/5:%(n)d"),
    "top [<n>]":                        Key("pgup/5:%(n)d"),
    "pown [<n>]":                       Key("pgdown/5:%(n)d"),
    "up <n> (page | pages)":            Key("pgup/5:%(n)d"),
    "down <n> (page | pages)":          Key("pgdown/5:%(n)d"),
    "left <n> (word | words)":          Key("c-left/5:%(n)d"),
    "right <n> (word | words)":         Key("c-right/5:%(n)d"),
    "west":                             Key("home"),
    "east":                              Key("end"),
    "north":                            Key("c-home"),
    "south":                           Key("c-end"),
    "yankee|Y":                           Key("y"),
    "november|N":                           Key("n"),

    "crack [<n>]":                     release + Key("del/5:%(n)d"),
    "delete [<n> | this] (line|lines)": release + Key("home, s-down/5:%(n)d, del"),
    "snap [<n>]":                  release + Key("backspace/5:%(n)d"),
    "pop up":                           release + Key("apps"),
    "cancel":                             release + Key("escape"),
    "(volume|audio|turn it) up": Key("volumeup"), 
    "(volume|audio|turn it) down": Key("volumedown"), 
    "mute": Key("volumemute"),

    "paste":                            release + Key("c-v"),
    "duplicate <n>":                    release + Key("c-c, c-v/5:%(n)d"),
    "copy":                             release + Key("c-c"),
    "cut":                              release + Key("c-x"),
    "select everything":                       release + Key("c-a"),
    "[hold] shift":                     Key("shift:down"),
    "release shift":                    Key("shift:up"),
    "[hold] control":                   Key("ctrl:down"),
    "release control":                  Key("ctrl:up"),
    "release [all]":                    release,

    "(I|eye) connect": Function(connect),
    "(I|eye) disconnect": Function(disconnect),
    "(I|eye) print position": Function(print_position),
    "(I|eye) move": Function(move_to_position),
    "(I|eye) click": Function(move_to_position) + Mouse("left"),
    "(I|eye) act": Function(activate_position), 
    "(I|eye) right click": Function(move_to_position) + Mouse("right"),
    "(I|eye) middle click": Function(move_to_position) + Mouse("middle"),
    "(I|eye) double click": Function(move_to_position) + Mouse("left:2"),
    "(I|eye) triple click": Function(move_to_position) + Mouse("left:3"),
    "(I|eye) start drag": Function(move_to_position) + Mouse("left:down"),
    "(I|eye) stop drag": Function(move_to_position) + Mouse("left:up"),
}

# Actions for speaking out sequences of characters.
character_action_map = {
    "plain <char>": Text("%(char)s"),
    "number <numerals>": Text("%(numerals)s"),
    "print <letters>": Text("%(letters)s"),
    "shout <letters>": Function(lambda letters: Text(letters.upper()).execute()),
}

# Actions that can be used anywhere in any command.
global_action_map = combine_maps(global_key_action_map,
                                 text_map_to_action_map(symbol_map))

# Actions that can be used anywhere except after a command with arbitrary
# dictation.
command_action_map = combine_maps(global_action_map, key_action_map)

# Here we prepare the action map of formatting functions from the config file.
# Retrieve text-formatting functions from this module's config file. Each of
# these functions must have a name that starts with "format_".
format_functions = {}
if namespace:
    for name, function in namespace.items():
        if name.startswith("format_") and callable(function):
            spoken_form = function.__doc__.strip()

            # We wrap generation of the Function action in a function so
            #  that its *function* variable will be local.  Otherwise it
            #  would change during the next iteration of the namespace loop.
            def wrap_function(function):
                def _function(dictation):
                    formatted_text = function(dictation)
                    Text(formatted_text).execute()
                return Function(_function)

            action = wrap_function(function)
            format_functions[spoken_form] = action

#-------------------------------------------------------------------------------
# Simple elements that may be referred to within a rule.

numbers_dict_list  = DictList("numbers_dict_list", numbers_map)
letters_dict_list = DictList("letters_dict_list", letters_map)
short_letters_dict_list = DictList("short_letters_dict_list", short_letters_map)
long_letters_dict_list = DictList("long_letters_dict_list", long_letters_map)
char_dict_list = DictList("char_dict_list", char_map)
saved_word_list = List("saved_word_list", saved_words)
# Lists which will be populated later via RPC.
# Dummy value is to work around dragonfly bug where list is not added to a
# grammar if it is equal to an existing list.
context_phrase_list = List("context_phrase_list", ["dummy_context_phrase_list"])
context_word_list = List("context_word_list", ["dummy_context_word_list"])

# Dictation consisting of sources of contextually likely words.
custom_dictation = Alternative([
    ListRef(None, saved_word_list),
    ListRef(None, context_phrase_list),
])

# Either arbitrary dictation or letters.
mixed_dictation = Alternative([
    Dictation(),
    DictListRef(None, short_letters_dict_list), 
    DictListRef(None, long_letters_dict_list), 
])

# A sequence of either short letters or long letters.
letters_element = Alternative([
    JoinedRepetition("", DictListRef(None, short_letters_dict_list), min = 1, max = 10),
    JoinedRepetition("", DictListRef(None, long_letters_dict_list), min = 1, max = 10),
])

# Simple element map corresponding to keystroke action maps from earlier.
keystroke_element_map = {
    "n": (IntegerRef(None, 1, 100), 1),
    "text": Dictation(),
    "char": DictListRef(None, char_dict_list),
}

#-------------------------------------------------------------------------------
# Rules which we will refer to within other rules.

# Rule for formatting mixed_dictation elements.
format_rule = create_rule(
    "FormatRule",
    format_functions,
    {"dictation": mixed_dictation}
)

# Rule for formatting custom_dictation elements.
custom_format_rule = create_rule(
    "CustomFormatRule",
    dict([("my " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": custom_dictation}
)

# Rule for handling raw dictation.
dictation_rule = create_rule(
    "DictationRule",
    {
        "say <text>":                       release + Text("%(text)s"),
        "mimic <text>":                     release + Mimic(extra="text"),
    },
    {
        "text": Dictation()
    }
)

# Rule for printing single characters.
single_character_rule = create_rule(
    "SingleCharacterRule",
    character_action_map,
    {
        "numerals": DictListRef(None, numbers_dict_list),
        "letters": Alternative([DictListRef(None, short_letters_dict_list),
                                DictListRef(None, long_letters_dict_list)]), 
        "char": DictListRef(None, char_dict_list),
    }
)

# Rule for spelling a word letter by letter and formatting it.
spell_format_rule = create_rule(
    "SpellFormatRule",
    dict([("spell " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": letters_element}
)

# Rule for printing a sequence of characters.
character_rule = create_rule(
    "CharacterRule",
    character_action_map,
    {
        "numerals": JoinedRepetition("", DictListRef(None, numbers_dict_list),
                                       min = 0, max = 10),
        "letters": letters_element,
        "char": DictListRef(None, char_dict_list),
    }
)
    
#-------------------------------------------------------------------------------
# Elements that are composed of rules. Note that the value of these elements are
# actions which will have to be triggered manually.

# Element matching simple commands.
# For efficiency, this should not contain any repeating elements.
single_action = RuleRef(rule=create_rule("CommandKeystrokeRule",
                                         command_action_map,
                                         keystroke_element_map))

# Element matching dictation and commands allowed at the end of an utterance.
# For efficiency, this should not contain any repeating elements. For accuracy,
# few custom commands should be included to avoid clashes with dictation
# elements.
terminal_element = Alternative([
    RuleRef(rule=dictation_rule),
    RuleRef(rule=format_rule),
    RuleRef(rule=custom_format_rule),
    RuleRef(rule=create_rule("GlobalKeystrokeRule",
                             global_action_map,
                             keystroke_element_map)), 
    RuleRef(rule=single_character_rule),
])


#---------------------------------------------------------------------------
# Here we define the top-level rule which the user can say.

# This is the rule that actually handles recognitions.
#  When a recognition occurs, its _process_recognition()
#  method will be called.  It receives information about the
#  recognition in the "extras" argument: the sequence of
#  actions and the number of times to repeat them.
class RepeatRule(CompoundRule):
    def __init__(self, name, repeated, terminal, context):
        # Here we define this rule's spoken-form and special elements. Note that
        # the middle element is the only one that contains Repetitions, and it
        # is not itself repeated. This is for performance purposes. We also
        # include a special escape command "terminal <terminal_command>" in case
        # recognition problems occur with repeated terminal commands.
        spec     = "[<sequence>] [<middle_element>] [<terminal_sequence>] [terminal <terminal>] [[[and] repeat [that]] <n> times]"
        extras   = [
            Repetition(repeated, min=1, max = 5, name="sequence"), 
            Alternative([RuleRef(rule=character_rule), RuleRef(rule=spell_format_rule)],
                        name="middle_element"),
            Repetition(terminal, min = 1, max = 5, name = "terminal_sequence"), 
            Alternative([terminal], name = "terminal"), 
            IntegerRef("n", 1, 100),  # Times to repeat the sequence.
        ]
        defaults = {
            "n": 1,                   # Default repeat count.
            "sequence": [], 
            "middle_element": None, 
            "terminal_sequence": [],
            "terminal": None, 
        }

        CompoundRule.__init__(self, name=name, spec=spec,
                              extras=extras, defaults=defaults, exported=True, context=context)

    # This method gets called when this rule is recognized.
    # Arguments:
    #  - node -- root node of the recognition parse tree.
    #  - extras -- dict of the "extras" special elements:
    #     . extras["sequence"] gives the sequence of actions.
    #     . extras["n"] gives the repeat count.
    def _process_recognition(self, node, extras):
        sequence = extras["sequence"]   # A sequence of actions.
        middle_element = extras["middle_element"]
        terminal_sequence = extras["terminal_sequence"]
        terminal = extras["terminal"]
        count = extras["n"]             # An integer repeat count.
        for i in range(count):
            for action in sequence:
                action.execute()
                Pause("5").execute()
            if middle_element:
                middle_element.execute()
            for action in terminal_sequence:
                action.execute()
                Pause("5").execute()
            if terminal:
                terminal.execute()
        release.execute()

#-------------------------------------------------------------------------------
# Define top-level rules for different contexts. Note that Dragon only allows
# top-level rules to be context-specific, but we want control over sub-rules. To
# work around this limitation, we compile a mutually exclusive top-level rule
# for each context.

class ContextHelper:
    """Helper to define a context hierarchy in terms of sub-rules but pass it to
    dragonfly as top-level rules."""

    def __init__(self, name, context, element):
        """Associate the provided context with the element to be repeated."""
        self.name = name
        self.context = context
        self.element = element
        self.children = []

    def add_child(self, child):
        """Add child ContextHelper."""
        self.children.append(child)

    def add_rules(self, grammar, parent_context):
        """Walk the ContextHelper tree and add exclusive top-level rules to the
        grammar."""
        full_context = combine_contexts(parent_context, self.context)
        exclusive_context = full_context
        for child in self.children:
            child.add_rules(grammar, full_context)
            exclusive_context = combine_contexts(exclusive_context, ~child.context)
        grammar.add_rule(RepeatRule(self.name + "RepeatRule",
                                    self.element,
                                    terminal_element,
                                    exclusive_context))

global_context_helper = ContextHelper("Global", None, single_action)

shell_command_map = combine_maps({
        "five sync": Text("5 sync "),
        "five merge": Text("5 merge "),
        "five diff": Text("5 diff "),
        "five mail": Text("5 mail -m "),
        "five lint": Text("5 lint "),
        "five submit": Text("5 submit "),
        "five cleanup": Text("5 c "),
        "five pending": Text("5 p "),
        "five export": Text("5 e "), 
        "five fix": Text("5 fix "), 
        "git commit": Text("git commit -am "),
        "git commit done": Text("git commit -am done "),
        "git checkout new": Text("git checkout -b "),
        "git reset hard head": Text("git reset --hard HEAD "), 
        "soft link": Text("ln -s "),
        "list": Text("ls -l "), 
        "make dir": Text("mkdir "), 
    }, dict((command, Text(command + " ")) for command in [
        "cd",
        "ls",
        "git status",
        "git branch",
        "git diff",
        "git checkout",
        "git stash",
        "git stash pop", 
    ]))


def Exec(command):
    return Key("c-c, a-x") + Text(command) + Key("enter")

# Work in progress.
def FastExec(command):
    return Function(lambda: urllib.urlopen("http://127.0.0.1:9091/" + command).close())

emacs_action_map = combine_maps(
    command_action_map,
    {
        "up [<n>]": Key("c-u") + Text("%(n)s") + Key("up"),
        "down [<n>]": Key("c-u") + Text("%(n)s") + Key("down"),
        "crack [<n>]": Key("c-u") + Text("%(n)s") + Key("c-d"),
        "kimble [<n>]": Key("c-u") + Text("%(n)s") + Key("as-d"),
        "(shuck|undo)": Key("c-slash"),
        "scratch": Key("c-x, r, U, U"),  
        "redo": Key("c-question"),
        "split fub": Key("c-x, 3"),
        "clote fub": Key("c-x, 0"),
        "only fub": Key("c-x, 1"), 
        "other fub": Key("c-x, o"),
        "die fub": Key("c-x, k"),
        "even fub": Key("c-x, plus"), 
        "open bookmark": Key("c-x, r, b"),
        "indent region": Key("ca-backslash"), 
        "comment region": Key("a-semicolon"), 
        "project file": Key("c-c, p, h"),
        "switch project": Key("c-c, p, s"),
        "build file": Key("c-c/10, c-g"),
        "test file": Key("c-c, c-t"),
        "helm": Key("c-x, c"),
        "helm resume": Key("c-x, c, b"), 
        "line <line>": Key("a-g, a-g") + Text("%(line)s") + Key("enter"),
        "re-center": Key("c-l"),
        "set mark": Key("c-backtick"), 
        "jump mark": Key("c-langle"),
        "jump symbol": Key("a-i"), 
        "select region": Key("c-x, c-x"),
        "swap mark": Key("c-c, c-x"),
        "slap above": Key("a-enter"),
        "slap below": Key("c-enter"), 
        "move (line|lines) up [<n>]": Key("c-u") + Text("%(n)d") + Key("a-up"),
        "move (line|lines) down [<n>]": Key("c-u") + Text("%(n)d") + Key("a-down"),
        "copy (line|lines) up [<n>]": Key("c-u") + Text("%(n)d") + Key("as-up"),
        "copy (line|lines) down [<n>]": Key("c-u") + Text("%(n)d") + Key("as-down"),
        "clear line": Key("c-a, c-c, c, k"), 
        "join (line|lines)": Key("as-6"), 
        "white": Key("a-m"),
        "buff": Key("c-x, b"),
        "oaf": Key("c-x, c-f"),
        "dired": Key("c-d"),
        "furred [<n>]": Key("a-f/5:%(n)d"),
        "bird [<n>]": Key("a-b/5:%(n)d"),
        "kurd [<n>]": Key("a-d/5:%(n)d"),
        "nope": Key("c-g"),
        "no way": Key("c-g/5:3"),
        "(prev|preev) [<n>]": Key("c-r/5:%(n)d"),
        "next [<n>]": Key("c-s/5:%(n)d"),
        "edit search": Key("a-e"),
        "word search": Key("a-s, w"),
        "symbol search": Key("a-s, underscore"), 
        "replace": Key("as-5"),
        "replace symbol": Key("a-apostrophe"),
        "(prev|preev) symbol": Key("c-c, c, c-r, a-s, underscore, c-y"), 
        "(next symbol|neck symbol)": Key("c-c, c, c-s, a-s, underscore, c-y"),
        "jump before <context_word>": Key("c-r") + Text("%(context_word)s") + Key("enter"),
        "jump after <context_word>": Key("c-s") + Text("%(context_word)s") + Key("enter"),
        "next result": Key("a-comma"),
        "preev error": Key("f11"), 
        "next error": Key("f12"),
        "cut": Key("c-w"),
        "copy": Key("a-w"),
        "yank": Key("c-y"),
        "sank": Key("a-y"), 
        "Mark": Key("c-space"),
        "nasper": Key("ca-f"),
        "pesper": Key("ca-b"),
        "moosper": Key("cas-2"),
        "kisper": Key("ca-k"),
        "dowsper": Key("ca-d"),
        "usper": Key("ca-u"),
        "exec": Key("a-x"),
        "preelin": Key("a-p"),
        "nollin": Key("a-n"),
        "before [preev] <char>": Key("c-c, c, b") + Text("%(char)s"),
        "after [next] <char>": Key("c-c, c, f") + Text("%(char)s"),
        "before next <char>": Key("c-c, c, s") + Text("%(char)s"),
        "after preev <char>": Key("c-c, c, e") + Text("%(char)s"),
        "surround parens": Key("a-lparen"),
        "plate <template>": Key("c-c, ampersand, c-s") + Text("%(template)s") + Key("enter"), 
        "open template <template>": Key("c-c, ampersand, c-v") + Text("%(template)s") + Key("enter"),
        "open template": Key("c-c, ampersand, c-v"), 
        "prefix": Key("c-u"), 
        "quit": Key("q"),
        "save": Key("c-x, c-s"),
        "open definition": Key("c-c, comma, d"),
        "open cross references": Key("c-c, comma, x"),
        "clang format": Key("ca-q"),
        "format comment": Key("a-q"),
        "other top": Key("c-minus, ca-v"),
        "other pown": Key("ca-v"),
        "close tag": Key("c-c, c-e"),
        "I jump <char>": Key("c-c, c, j") + Text("%(char)s") + Function(lambda: type_position("%d\n%d\n")),
        "R grep": Exec("rgrep"),
        "code search": Exec("cs"),
        "code search car": Exec("csc"),
        "set position <char>": Key("c-x, r, space, %(char)s"),
        "jump position <char>": Key("c-x, r, j, %(char)s"),
        "copy register <char>": Key("c-x, r, s, %(char)s"),
        "yank register <char>": Key("c-u, c-x, r, i, %(char)s"),
        "expand diff": Key("a-4"),
        "refresh": Key("g"),
        "JavaScript mode": Exec("js-mode"),
        "HTML mode": Exec("html-mode"),
        "toggle details": Exec("dired-hide-details-mode"),
        "up fub": Exec("windmove-up"), 
        "down fub": Exec("windmove-down"), 
        "left fub": Exec("windmove-left"), 
        "right fub": Exec("windmove-right"),
        "shell (preev|back)": Key("a-r"),
        "hello": FastExec("hello-world"),
        "kill emacs server": Exec("ws-stop-all"), 
        "closure compile": Key("c-c, c-k"),
        "closure namespace": Key("c-c, a-n"), 
    })

templates = {
    "car": "car",
    "class": "class",
    "def": "function",
    "each": "each",
    "eval": "eval",
    "for": "for",
    "fun declaration": "fun_declaration",
    "function": "function",
    "if": "if",
    "method": "method", 
    "var": "vardef",
    "while": "while",
}
template_dict_list = DictList("template_dict_list", templates)
emacs_element_map = combine_maps(
    keystroke_element_map,
    {
        "line": IntegerRef(None, 1, 10000),
        "template": DictListRef(None, template_dict_list),
        "context_word": ListRef(None, context_word_list),
    })

emacs_element = RuleRef(rule=create_rule("EmacsKeystrokeRule", emacs_action_map, emacs_element_map))
emacs_context_helper = ContextHelper("Emacs", AppContext(title = "Emacs editor"), emacs_element)
global_context_helper.add_child(emacs_context_helper)

emacs_shell_action_map = combine_maps(
    emacs_action_map,
    shell_command_map,
    {
        "show output": Key("c-c, c-r"), 
    })
emacs_shell_element = RuleRef(rule=create_rule("EmacsShellKeystrokeRule", emacs_shell_action_map, emacs_element_map))
emacs_shell_context_helper = ContextHelper("EmacsShell", AppContext(title="- Shell -"), emacs_shell_element)
emacs_context_helper.add_child(emacs_shell_context_helper)


shell_action_map = combine_maps(
    command_action_map,
    shell_command_map,
    {
        "copy": Key("cs-c"), 
        "paste": Key("cs-v"), 
        "cut": Key("cs-x"), 
        "top [<n>]": Key("s-pgup/5:%(n)d"), 
        "pown [<n>]": Key("s-pgdown/5:%(n)d"), 
        "crack [<n>]": Key("c-d/5:%(n)d"),
        "pret [<n>]": Key("cs-left/5:%(n)d"),
        "net [<n>]": Key("cs-right/5:%(n)d"),
        "move tab left [<n>]": Key("cs-pgup/5:%(n)d"), 
        "move tab right [<n>]": Key("cs-pgdown/5:%(n)d"), 
        "(prev|preev|back)": Key("c-r"),
        "(next|frack)": Key("c-s"), 
        "(nope|no way)": Key("c-g"),
        "new tab": Key("cs-t"),
        "clote": Key("cs-w"),
        "forward": Key("f"),
        "backward": Key("b"),
        "quit": Key("q"),
        "kill process": Key("c-c"),
    })

shell_element = RuleRef(rule=create_rule("ShellKeystrokeRule", shell_action_map, keystroke_element_map))
shell_context_helper = ContextHelper("Shell", AppContext(title = "Terminal"), shell_element)
global_context_helper.add_child(shell_context_helper)


chrome_action_map = combine_maps(
    command_action_map,
    {
        "link": Key("c-comma"),
        "new link": Key("c-dot"),
        "background links": Key("a-f"), 
        "new tab":            Key("c-t"),
        "new incognito":            Key("cs-n"),
        "new window": Key("c-n"),
        "clote":          Key("c-w"),
        "(search|address) bar":        Key("c-l"),
        "back [<n>]":               Key("a-left/15:%(n)d"),
        "Frak [<n>]":            Key("a-right/15:%(n)d"),
        "reload": Key("c-r"),
        "shot <tab_n>": Key("c-%(tab_n)d"),
        "shot last": Key("c-9"), 
        "net [<n>]":           Key("c-tab:%(n)d"),
        "pret [<n>]":           Key("cs-tab:%(n)d"),
        "move tab left [<n>]": Key("cs-pgup/5:%(n)d"), 
        "move tab right [<n>]": Key("cs-pgdown/5:%(n)d"),
        "move tab <tab_n>": Key("cs-%(tab_n)d"),
        "move tab last": Key("cs-9"), 
        "reote":         Key("cs-t"),
        "duplicate tab": Key("c-l/15, a-enter"), 
        "find":               Key("c-f"),
        "<link>":          Text("%(link)s"), 
        "search <text>":        Key("c-l/15") + Text("%(text)s") + Key("enter"),
        "moma": Key("c-l/15") + Text("moma") + Key("tab"),
        "code search car": Key("c-l/15") + Text("csc") + Key("tab"),
        "code search": Key("c-l/15") + Text("cs") + Key("tab"),
        "go to calendar": Key("c-l/15") + Text("c/") + Key("enter"),
        "go to critique": Key("c-l/15") + Text("cr/") + Key("enter"),
        "go to (buganizer|bugs)": Key("c-l/15") + Text("b/") + Key("enter"),
        "go to presubmits": Key("c-l/15, b, tab") + Text("One shot") + Key("enter:2"), 
        "go to postsubmits": Key("c-l/15, b, tab") + Text("Continuous") + Key("enter:2"), 
        "go to latest test results": Key("c-l/15, b, tab") + Text("latest test results") + Key("enter:2"), 
        "go to docs": Key("c-l/15") + Text("docs.google.com") + Key("enter"),
        "go to slides": Key("c-l/15") + Text("slides.google.com") + Key("enter"),
        "go to sheets": Key("c-l/15") + Text("sheets.google.com") + Key("enter"),
        "go to drive": Key("c-l/15") + Text("drive.google.com") + Key("enter"),
        "go to amazon": Key("c-l/15") + Text("smile.amazon.com") + Key("enter"),
        "(new|insert) row": Key("a-i/15, r"),
        "delete row": Key("a-e/15, d"),
        "strikethrough": Key("as-5"),
        "bullets": Key("cs-8"),
        "bold": Key("c-b"), 
        "text box": Key("a-i/15, t"),
        "copy format": Key("ca-c"), 
        "paste format": Key("ca-v"),
        "select column": Key("c-space"), 
        "select row": Key("s-space"),
        "row up": Key("a-e/15, k"), 
        "row down": Key("a-e/15, j"),
        "column left": Key("a-e/15, m"), 
        "column right": Key("a-e/15, m"), 
        "next match": Key("c-g"),
        "preev match": Key("cs-g"),
        "(go to|open) bookmark": Key("c-semicolon"),
        "new bookmark": Key("c-apostrophe"),
        "save bookmark": Key("c-d"), 
        "next frame": Key("c-lbracket"),
        "developer tools": Key("cs-j"),
        "create driver": Function(create_driver),
        "quit driver": Function(quit_driver),
        "test driver": Function(test_driver),
    })

link_char_map = {
    "zero": "0",
    "one": "1",
    "two": "2",
    "three": "3",
    "four": "4",
    "five": "5",
    "six": "6",
    "seven": "7",
    "eight": "8",
    "nine": "9",
}
link_char_dict_list  = DictList("link_char_dict_list", link_char_map)
chrome_element_map = combine_maps(
    keystroke_element_map,
    {
        "tab_n": IntegerRef(None, 1, 9),
        "link": JoinedRepetition("", DictListRef(None, link_char_dict_list), min = 0, max = 5),
    })

chrome_element = RuleRef(rule=create_rule("ChromeKeystrokeRule", chrome_action_map, chrome_element_map))
chrome_context_helper = ContextHelper("Chrome", AppContext(executable="chrome"), chrome_element)
global_context_helper.add_child(chrome_context_helper)

critique_action_map = combine_maps(
    chrome_action_map,
    {
        "preev": Key("p"), 
        "next": Key("n"),
        "preev file": Key("k"),
        "next file": Key("j"),
        "open": Key("o"),
        "list": Key("u"),
        "comment": Key("c"),
        "save": Key("c-s"),
        "expand|collapse": Key("e"),
        "reply": Key("r"), 
    })
critique_element = RuleRef(rule=create_rule("CritiqueKeystrokeRule", critique_action_map, chrome_element_map))
critique_context_helper = ContextHelper("Critique", AppContext(title = "<critique.corp.google.com>"), critique_element)
chrome_context_helper.add_child(critique_context_helper)

code_search_action_map = combine_maps(
    chrome_action_map,
    {
        "header": Key("r/25, h"),
        "source": Key("r/25, c"), 
    })
code_search_element = RuleRef(rule=create_rule("CodeSearchKeystrokeRule", code_search_action_map, chrome_element_map))
code_search_context_helper = ContextHelper("CodeSearch", AppContext(title = "<cs.corp.google.com>"), code_search_element)
chrome_context_helper.add_child(code_search_context_helper)

gmail_action_map = combine_maps(
    chrome_action_map,
    {
        "open": Key("o"),
        "(archive|done)": Text("["),
        "list": Key("u"),
        "preev": Key("k"),
        "next": Key("j"),
        "preev message": Key("p"),
        "next message": Key("n"),
        "compose": Key("c"),
        "reply": Key("r"),
        "reply all": Key("a"),
        "forward": Key("f"),
        "important": Key("plus"),
        "next section": Key("backtick"),
        "preev section": Key("tilde"),
        "not important|don't care": Key("minus"),
        "label waiting": Key("l/25") + Text("waiting") + Key("enter"),
        "select": Key("x"),
        "select next <n>": Key("x, j") * Repeat(extra="n"), 
        "new messages": Key("N"),
        "go to inbox": Key("g, i"), 
        "go to starred": Key("g, s"), 
        "go to sent": Key("g, t"),
        "expand all": ClickElementAction(xpath="//*[@aria-label='Expand all']"),
    })

gmail_element = RuleRef(rule=create_rule("GmailKeystrokeRule", gmail_action_map, chrome_element_map))
gmail_context_helper = ContextHelper("Gmail",
                                     (AppContext(title = "Gmail") |
                                      AppContext(title = "Google.com Mail") |
                                      AppContext(title = "<mail.google.com>") |
                                      AppContext(title = "<inbox.google.com>")),
                                     gmail_element)
chrome_context_helper.add_child(gmail_context_helper)

#-------------------------------------------------------------------------------
# Populate and load the grammar.

grammar = Grammar("repeat")   # Create this module's grammar.
global_context_helper.add_rules(grammar, None)
grammar.load()

#-------------------------------------------------------------------------------
# Start a server which lets Emacs send us nearby text being edited, so we can
# use it for contextual recognition. Note that the server is only exposed to
# the local machine, except it is still potentially vulnerable to CSRF
# attacks. Consider this when adding new functionality.

# Register timer to run arbitrary callbacks added by the server.
callbacks = Queue.Queue()

def RunCallbacks():
    while not callbacks.empty():
        callbacks.get_nowait()()

timer = get_engine().create_timer(RunCallbacks, 0.1)

# Update the context words and phrases.
def UpdateWords(words, phrases):
    context_word_list.set(words)
    context_phrase_list.set(phrases)

# HTTP handler for receiving a block of text. The file type can be provided with
# header My-File-Type.
# TODO: Use JSON instead of custom headers and dispatch based on path.
class TextRequestHandler(BaseHTTPServer.BaseHTTPRequestHandler):
    def do_POST(self):
    # Uncomment to enable profiling.
    #     cProfile.runctx("self.PostInternal()", globals(), locals())
    # def PostInternal(self):
        start_time = time.time()
        length = self.headers.getheader("content-length")
        file_type = self.headers.getheader("My-File-Type")
        text = self.rfile.read(int(length)) if length else ""
        # print("received text: %s" % text)
        words = dragonfly_words.ExtractWords(text, file_type)
        phrases = dragonfly_words.ExtractPhrases(text, file_type)
        # Asynchronously update word lists available to Dragon.
        callbacks.put_nowait(lambda: UpdateWords(words, phrases))
        self.send_response(204)  # no content
        self.end_headers()
        # The following sequence of low-level socket commands was needed to get
        # this working properly with the Emacs client. Gory details:
        # http://blog.netherlabs.nl/articles/2009/01/18/the-ultimate-so_linger-page-or-why-is-my-tcp-not-reliable
        self.wfile.flush()
        self.request.shutdown(socket.SHUT_WR)
        self.rfile.read()
        print("Processed words: %.10f" % (time.time() - start_time))
    def do_GET(self):
        self.do_POST()

# Start a single-threaded HTTP server in a separate thread. Bind the server to
# localhost so it cannot be accessed outside the local computer (except by SSH
# tunneling).
HOST, PORT = "127.0.0.1", 9090
server = BaseHTTPServer.HTTPServer((HOST, PORT), TextRequestHandler)
server_thread = threading.Thread(target = server.serve_forever)
server_thread.start()
print("started server")

#-------------------------------------------------------------------------------
# Unload function which will be called by NatLink.
def unload():
    global grammar, server, server_thread, timer
    if grammar:
        grammar.unload()
        grammar = None
    timer.stop()
    server.shutdown()
    server_thread.join()
    print("shutdown server")
