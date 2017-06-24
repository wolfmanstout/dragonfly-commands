#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""This contains all commands which may be spoken continuously or repeated.

This is heavily modified from _multiedit.py, found here:
https://github.com/t4ngo/dragonfly-modules/blob/master/command-modules/_multiedit.py
"""

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r99")
except ImportError:
    pass

import BaseHTTPServer
import Queue
import platform
import socket
import threading
import time
import webbrowser
import win32clipboard

from dragonfly import (
    ActionBase,
    Alternative,
    AppContext,
    CompoundRule,
    Config,
    DictList,
    DictListRef,
    Dictation,
    Empty,
    Function,
    Grammar,
    IntegerRef,
    Key,
    List,
    ListRef,
    Mimic,
    Mouse,
    Optional,
    Pause,
    Repeat,
    Repetition,
    RuleRef,
    RuleWrap,
    Text,
    get_engine,
)
import dragonfly.log
from selenium.webdriver.common.by import By

import _dragonfly_utils as utils
import _eye_tracker_utils as eye_tracker
import _linux_utils as linux
import _text_utils as text
import _webdriver_utils as webdriver

# Load local hooks if defined.
try:
    import _dragonfly_local_hooks as local_hooks
    def run_local_hook(name, *args, **kwargs):
        """Function to run local hook if defined."""
        try:
            hook = getattr(local_hooks, name)
            return hook(*args, **kwargs)
        except AttributeError:
            pass
except:
    print("Local hooks not loaded.")
    def run_local_hook(name, *args, **kwargs):
        pass

# Make sure dragonfly errors show up in NatLink messages.
dragonfly.log.setup_log()

# Load _repeat.txt.
config = Config("repeat")
namespace = config.load()

#-------------------------------------------------------------------------------
# Common maps and lists.
symbol_map = {
    "plus": " + ",
    "dub plus": "++",
    "minus": " - ",
    "nad": ", ",
    "coal": ":",
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

quick_letters_map = {
    "arch": "a",
    "brov": "b",
    "chair": "c",
    "dell": "d",
    "etch": "e",
    "fomp": "f",
    "goof": "g",
    "hark": "h",
    "ice": "i",
    "jinks": "j",
    "koop": "k",
    "lug": "l",
    "mowsh": "m",
    "nerb": "n",
    "ork": "o",
    "pooch": "p",
    "quash": "q",
    "rosh": "r",
    "souk": "s",
    "teek": "t",
    "unks": "u",
    "verge": "v",
    "womp": "w",
    "trex": "x",
    "yang": "y",
    "zooch": "z",
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

prefixes = [
    "num",
    "min",
]

suffixes = [
    "bytes",
]

letters_map = utils.combine_maps(quick_letters_map, long_letters_map)

char_map = dict((k, v.strip())
                for (k, v) in utils.combine_maps(letters_map, numbers_map, symbol_map).iteritems())

# Load commonly misrecognized words saved to a file.
saved_words = []
try:
    with open(text.WORDS_PATH) as file:
        for line in file:
            word = line.strip()
            if len(word) > 2 and word not in letters_map:
                saved_words.append(line.strip())
except:
    print("Unable to open: " + text.WORDS_PATH)

#-------------------------------------------------------------------------------
# Action maps to be used in rules.

# Key actions which may be used anywhere in any command.
global_key_action_map = {
    "slap [<n>]": Key("enter/5:%(n)d"),
    "spooce [<n>]": Key("space/5:%(n)d"),
    "tab [<n>]": Key("tab/5:%(n)d"),
}

# Actions of commonly used text navigation and mousing commands. These can be
# used anywhere except after commands which include arbitrary dictation.
release = Key("shift:up, ctrl:up, alt:up")
key_action_map = {
    "up [<n>]":                         Key("up/5:%(n)d"),
    "down [<n>]":                       Key("down/5:%(n)d"),
    "left [<n>]":                       Key("left/5:%(n)d"),
    "right [<n>]":                      Key("right/5:%(n)d"),
    "fomble [<n>]": Key("c-right/5:%(n)d"),
    "bamble [<n>]": Key("c-left/5:%(n)d"),
    "dumbbell [<n>]": Key("c-backspace/5:%(n)d"),
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
    "yankee|yang":                           Key("y"),
    "november|nerb":                           Key("n"),

    "crack [<n>]":                     release + Key("del/5:%(n)d"),
    "delete [<n> | this] (line|lines)": release + Key("home, s-down/5:%(n)d, del"),
    "snap [<n>]":                  release + Key("backspace/5:%(n)d"),
    "pop up":                           release + Key("apps"),
    "cancel|escape":                             release + Key("escape"),
    "(volume|audio|turn it) up": Key("volumeup"),
    "(volume|audio|turn it) down": Key("volumedown"),
    "(volume|audio) mute": Key("volumemute"),
    "next track": Key("tracknext"),
    "prev track": Key("trackprev"),
    "play pause|pause play": Key("playpause"),

    "paste":                            release + Key("c-v"),
    "copy":                             release + Key("c-c"),
    "cut":                              release + Key("c-x"),
    "select everything":                       release + Key("c-a"),
    "edit text": utils.RunApp("notepad"),
    "edit emacs": utils.RunEmacs(".txt"),
    "edit everything": Key("c-a, c-x") + utils.RunApp("notepad") + Key("c-v"),
    "edit region": Key("c-x") + utils.RunApp("notepad") + Key("c-v"),
    "[hold] shift":                     Key("shift:down"),
    "release shift":                    Key("shift:up"),
    "[hold] control":                   Key("ctrl:down"),
    "release control":                  Key("ctrl:up"),
    "[hold] (meta|alt)":                   Key("alt:down"),
    "release (meta|alt)":                  Key("alt:up"),
    "release [all]":                    release,

    "(I|eye) connect": Function(eye_tracker.connect),
    "(I|eye) disconnect": Function(eye_tracker.disconnect),
    "(I|eye) print position": Function(eye_tracker.print_position),
    "(I|eye) move": Function(eye_tracker.move_to_position),
    "(I|eye) click": Function(eye_tracker.move_to_position) + Mouse("left"),
    "(I|eye) act": Function(eye_tracker.activate_position),
    "(I|eye) pan": Function(eye_tracker.panning_step_position),
    "(I|eye) right click": Function(eye_tracker.move_to_position) + Mouse("right"),
    "(I|eye) middle click": Function(eye_tracker.move_to_position) + Mouse("middle"),
    "(I|eye) double click": Function(eye_tracker.move_to_position) + Mouse("left:2"),
    "(I|eye) triple click": Function(eye_tracker.move_to_position) + Mouse("left:3"),
    "(I|eye) start drag": Function(eye_tracker.move_to_position) + Mouse("left:down"),
    "(I|eye) stop drag": Function(eye_tracker.move_to_position) + Mouse("left:up"),
    "scrup": Function(lambda: eye_tracker.move_to_position((0, 50))) + Mouse("scrollup:8"), 
    "half scrup": Function(lambda: eye_tracker.move_to_position((0, 50))) + Mouse("scrollup:4"), 
    "scrown": Function(lambda: eye_tracker.move_to_position((0, -50))) + Mouse("scrolldown:8"), 
    "half scrown": Function(lambda: eye_tracker.move_to_position((0, -50))) + Mouse("scrolldown:4"), 
    "do click": Mouse("left"),
    "do right click": Mouse("right"),
    "do middle click": Mouse("middle"),
    "do double click": Mouse("left:2"),
    "do triple click": Mouse("left:3"),
    "do start drag": Mouse("left:down"),
    "do stop drag": Mouse("left:up"),

    "create driver": Function(webdriver.create_driver),
    "quit driver": Function(webdriver.quit_driver),
}

# Actions for speaking out sequences of characters.
character_action_map = {
    "plain <chars>": Text("%(chars)s"),
    "numbers <numerals>": Text("%(numerals)s"),
    "print <letters>": Text("%(letters)s"),
    "shout <letters>": Function(lambda letters: Text(letters.upper()).execute()),
}

# Actions that can be used anywhere in any command.
global_action_map = utils.combine_maps(global_key_action_map,
                                       utils.text_map_to_action_map(symbol_map))

# Actions that can be used anywhere except after a command with arbitrary
# dictation.
command_action_map = utils.combine_maps(global_action_map, key_action_map)

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
char_dict_list = DictList("char_dict_list", char_map)
saved_word_list = List("saved_word_list", saved_words)
# Lists which will be populated later via RPC.
context_phrase_list = List("context_phrase_list", [])
context_word_list = List("context_word_list", [])
prefix_list = List("prefix_list", prefixes)
suffix_list = List("suffix_list", suffixes)

# Dictation consisting of sources of contextually likely words.
custom_dictation = RuleWrap(None, Alternative([
    ListRef(None, saved_word_list),
    ListRef(None, context_phrase_list),
]))

# Either arbitrary dictation or letters.
mixed_dictation = RuleWrap(None, utils.JoinedSequence(" ", [
    Optional(ListRef(None, prefix_list)),
    Alternative([
        Dictation(),
        DictListRef(None, letters_dict_list),
        ListRef(None, saved_word_list),
    ]),
    Optional(ListRef(None, suffix_list)),
]))

# A sequence of either short letters or long letters.
letters_element = RuleWrap(None, utils.JoinedRepetition(
    "", DictListRef(None, letters_dict_list), min=1, max=10))

# A sequence of numbers.
numbers_element = RuleWrap(None, utils.JoinedRepetition(
    "", DictListRef(None, numbers_dict_list), min=0, max=10))

# A sequence of characters.
chars_element = RuleWrap(None, utils.JoinedRepetition(
    "", DictListRef(None, char_dict_list), min=0, max=10))

# Simple element map corresponding to keystroke action maps from earlier.
keystroke_element_map = {
    "n": (IntegerRef(None, 1, 21), 1),
    "text": Dictation(),
    "char": DictListRef(None, char_dict_list),
    "custom_text": RuleWrap(None, Alternative([
        Dictation(),
        DictListRef(None, char_dict_list),
        ListRef(None, prefix_list),
        ListRef(None, suffix_list),
        ListRef(None, saved_word_list),
    ])),
}

#-------------------------------------------------------------------------------
# Rules which we will refer to within other rules.

# Rule for formatting mixed_dictation elements.
format_rule = utils.create_rule(
    "FormatRule",
    format_functions,
    {"dictation": mixed_dictation}
)

# Rule for formatting pure dictation elements.
pure_format_rule = utils.create_rule(
    "PureFormatRule",
    dict([("pure " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": Dictation()}
)

# Rule for formatting custom_dictation elements.
custom_format_rule = utils.create_rule(
    "CustomFormatRule",
    dict([("my " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": custom_dictation}
)

# Rule for handling raw dictation.
dictation_rule = utils.create_rule(
    "DictationRule",
    {
        "(mim|mimic) text <text>": release + Text("%(text)s"),
        "mim small <text>": release + utils.uncapitalize_text_action("%(text)s"),
        "mim big <text>": release + utils.capitalize_text_action("%(text)s"),
        "mimic <text>": release + Mimic(extra="text"),
    },
    {
        "text": Dictation()
    }
)

# Rule for printing single characters.
single_character_rule = utils.create_rule(
    "SingleCharacterRule",
    character_action_map,
    {
        "numerals": DictListRef(None, numbers_dict_list),
        "letters": DictListRef(None, letters_dict_list),
        "chars": DictListRef(None, char_dict_list),
    }
)

# Rule for spelling a word letter by letter and formatting it.
spell_format_rule = utils.create_rule(
    "SpellFormatRule",
    dict([("spell " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": letters_element}
)

# Rule for printing a sequence of characters.
character_rule = utils.create_rule(
    "CharacterRule",
    character_action_map,
    {
        "numerals": numbers_element,
        "letters": letters_element,
        "chars": chars_element,
    }
)

#-------------------------------------------------------------------------------
# Elements that are composed of rules. Note that the value of these elements are
# actions which will have to be triggered manually.

# Element matching simple commands.
# For efficiency, this should not contain any repeating elements.
single_action = RuleRef(rule=utils.create_rule("CommandKeystrokeRule",
                                               command_action_map,
                                               keystroke_element_map))

# Element matching dictation and commands allowed at the end of an utterance.
# For efficiency, this should not contain any repeating elements. For accuracy,
# few custom commands should be included to avoid clashes with dictation
# elements.
dictation_element = RuleWrap(None, Alternative([
    RuleRef(rule=dictation_rule),
    RuleRef(rule=format_rule),
    RuleRef(rule=pure_format_rule),
    RuleRef(rule=custom_format_rule),
    RuleRef(rule=utils.create_rule("DictationKeystrokeRule",
                                   global_action_map,
                                   keystroke_element_map)),
    RuleRef(rule=single_character_rule),
]))


### Final commands that can be used once after everything else. These change the
### application context so it is important that nothing else gets run after
### them.

# Ordered list of pinned taskbar items. Sublists refer to windows within a specific application.
windows = [
    "explorer",
    ["dragonbar", "dragon [messages]", "dragonpad"],
    "home chrome",
    "home terminal",
    "home emacs",
]
json_windows = utils.load_json("windows.json")
if json_windows:
    windows = json_windows

windows_prefix = "go to"
windows_mapping = {}
for i, window in enumerate(windows):
    if isinstance(window, str):
        window = [window]
    for j, words in enumerate(window):
        windows_mapping[windows_prefix + " (" + words + ")"] = Key("win:down, %d:%d/20, win:up" % (i + 1, j + 1))

# Work around security restrictions in Windows 8.
if platform.release() == "8":
    swap_action = Mimic("press", "alt", "tab")
else:
    swap_action = Key("alt:down, tab:%(n)d/25, alt:up")

final_action_map = utils.combine_maps(windows_mapping, {
    "swap [<n>]": swap_action,
})
final_element_map = {
    "n": (IntegerRef(None, 1, 20), 1)
}
final_rule = utils.create_rule("FinalRule",
                               final_action_map,
                               final_element_map)


#---------------------------------------------------------------------------
# Here we define the top-level rule which the user can say.

# This is the rule that actually handles recognitions.
#  When a recognition occurs, its _process_recognition()
#  method will be called.  It receives information about the
#  recognition in the "extras" argument: the sequence of
#  actions and the number of times to repeat them.
class RepeatRule(CompoundRule):

    def __init__(self, name, command, terminal_command, context):
        # Here we define this rule's spoken-form and special elements. Note that
        # nested_repetitions is the only one that contains Repetitions, and it
        # is not itself repeated. This is for performance purposes. We also
        # include a special escape command "terminal <dictation>" in case
        # recognition problems occur with repeated dictation commands.
        spec = ("[<sequence>] "
                "[<nested_repetitions>] "
                "([<dictation_sequence>] [terminal <dictation>] | <terminal_command>) "
                "[[[and] repeat [that]] <n> times] "
                "[<final_command>]")
        extras = [
            Repetition(command, min=1, max = 5, name="sequence"),
            Alternative([RuleRef(rule=character_rule), RuleRef(rule=spell_format_rule)],
                        name="nested_repetitions"),
            Repetition(dictation_element, min=1, max=5, name="dictation_sequence"),
            utils.ElementWrapper("dictation", dictation_element),
            utils.ElementWrapper("terminal_command", terminal_command),
            IntegerRef("n", 1, 100),  # Times to repeat the sequence.
            RuleRef(rule=final_rule, name="final_command"),
        ]
        defaults = {
            "n": 1,                   # Default repeat count.
            "sequence": [],
            "nested_repetitions": None,
            "dictation_sequence": [],
            "dictation": None,
            "terminal_command": None,
            "final_command": None,
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
        nested_repetitions = extras["nested_repetitions"]
        dictation_sequence = extras["dictation_sequence"]
        dictation = extras["dictation"]
        terminal_command = extras["terminal_command"]
        final_command = extras["final_command"]
        count = extras["n"]             # An integer repeat count.
        for i in range(count):
            for action in sequence:
                action.execute()
                Pause("5").execute()
            if nested_repetitions:
                nested_repetitions.execute()
            for action in dictation_sequence:
                action.execute()
                Pause("5").execute()
            if dictation:
                dictation.execute()
            if terminal_command:
                terminal_command.execute()
        release.execute()
        if final_command:
            final_command.execute()


#-------------------------------------------------------------------------------
# Define top-level rules for different contexts. Note that Dragon only allows
# top-level rules to be context-specific, but we want control over sub-rules. To
# work around this limitation, we compile a mutually exclusive top-level rule
# for each context.

class Environment(object):
    """Environment where voice commands can be spoken. Combines grammar and context
    and adds hierarchy. When installed, will produce a top-level rule for each
    environment.
    """

    def __init__(self,
                 name,
                 environment_map,
                 context=None,
                 parent=None):
        self.name = name
        self.children = []
        if parent:
            parent.add_child(self)
            self.context = utils.combine_contexts(parent.context, context)
            self.environment_map = {}
            for key in set(environment_map.keys()) | set(parent.environment_map.keys()):
                action_map, element_map = environment_map.get(key, ({}, {}))
                parent_action_map, parent_element_map = parent.environment_map.get(key, ({}, {}))
                self.environment_map[key] = (utils.combine_maps(parent_action_map, action_map),
                                             utils.combine_maps(parent_element_map, element_map))
        else:
            self.context = context
            self.environment_map = environment_map

    def add_child(self, child):
        self.children.append(child)

    def install(self, grammar, exported_rule_factory):
        exclusive_context = self.context
        for child in self.children:
            child.install(grammar, exported_rule_factory)
            exclusive_context = utils.combine_contexts(exclusive_context, ~child.context)
        rule_map = dict([(key, RuleRef(rule=utils.create_rule(self.name + "_" + key, action_map, element_map)) if action_map else Empty())
                         for (key, (action_map, element_map)) in self.environment_map.items()])
        grammar.add_rule(exported_rule_factory(self.name + "_exported", exclusive_context, **rule_map))


class MyEnvironment(object):
    """Specialization of Environment for convenience with my exported rule factory
    (RepeatRule).
    """

    def __init__(self,
                 name,
                 parent=None,
                 context=None,
                 action_map=None,
                 terminal_action_map=None,
                 element_map=None):
        self.environment = Environment(
            name,
            {"command": (action_map or {}, element_map or {}),
             "terminal_command": (terminal_action_map or {}, element_map or {})},
            context,
            parent.environment if parent else None)

    def add_child(self, child):
        self.environment.add_child(child.environment)

    def install(self, grammar):
        def create_exported_rule(name, context, command, terminal_command):
            return RepeatRule(name, command or Empty(), terminal_command or Empty(), context)
        self.environment.install(grammar, create_exported_rule)


### Global

global_environment = MyEnvironment(name="Global",
                                   action_map=command_action_map,
                                   element_map=keystroke_element_map)


### Shell commands

shell_command_map = utils.combine_maps({
    "git commit": Text("git commit -am "),
    "git commit done": Text("git commit -am done "),
    "git checkout new": Text("git checkout -b "),
    "git reset hard head": Text("git reset --hard HEAD "),
    "(soft|sym) link": Text("ln -s "),
    "list": Text("ls -l "),
    "make dir": Text("mkdir "),
    "ps (a UX|aux)": Text("ps aux "),
    "kill command": Text("kill "),
    "pipe": Text(" | "),
    "CH mod": Text("chmod "),
    "TK diff": Text("tkdiff "),
    "MV": Text("mv "),
    "CP": Text("cp "),
    "RM": Text("rm "),
    "CD": Text("cd "),
    "LS": Text("ls "),
    "PS": Text("ps "),
    "reset terminal": Text("exec bash\n"),
    "pseudo": Text("sudo "),
    "apt get": Text("apt-get "),
}, dict((command, Text(command + " ")) for command in [
    "echo", 
    "grep",
    "ssh",
    "diff",
    "cat",
    "man",
    "less",
    "git status",
    "git branch",
    "git diff",
    "git checkout",
    "git stash",
    "git stash pop",
    "git push",
    "git pull",
]))
run_local_hook("AddShellCommands", shell_command_map)


### Emacs

def Exec(command):
    return Key("c-c, a-x") + Text(command) + Key("enter")


def jump_to_line(line_string):
    return Key("c-u") + Text(line_string) + Key("c-c, c, g")


class OpenClipboardUrlAction(ActionBase):
    """Open a URL in the clipboard in the default browser."""

    def _execute(self, data=None):
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        print "Opening link: %s" % data
        webbrowser.open(data)


class MarkLinesAction(ActionBase):
    """Mark several lines within a range."""

    def __init__(self, tight=False):
        super(MarkLinesAction, self).__init__()
        self.tight = tight

    def _execute(self, data=None):
        jump_to_line("%(n1)d" % data).execute()
        if self.tight:
            Key("a-m").execute()
        Key("c-space").execute()
        if "n2" in data:
            jump_to_line("%d" % (data["n2"])).execute()
        if self.tight:
            Key("c-e").execute()
        else:
            Key("down").execute()


class UseLinesAction(ActionBase):
    """Make use of lines within a range."""

    def __init__(self, pre_action, post_action, tight=False):
        super(UseLinesAction, self).__init__()
        self.pre_action = pre_action
        self.post_action = post_action
        self.tight = tight

    def _execute(self, data=None):
        # Set mark without activating.
        Key("c-backtick").execute()
        MarkLinesAction(self.tight).execute(data)
        self.pre_action.execute(data)
        # Jump to mark twice then to the beginning of the line.
        (Key("c-langle") + Key("c-langle")).execute()
        if not self.tight:
            Key("c-a").execute()
        self.post_action.execute(data)


emacs_action_map = {
    # Overrides
    "up [<n>]": Key("c-u") + Text("%(n)s") + Key("up"),
    "down [<n>]": Key("c-u") + Text("%(n)s") + Key("down"),
    # NX doesn't forward <delete> properly, so we avoid those bindings.
    "crack [<n>]": Key("c-d:%(n)d"),
    "kimble [<n>]": Key("as-d:%(n)d"),
    "select everything": Key("c-x, h"),
    "edit everything": Key("c-x, h, c-w") + utils.RunApp("notepad") + Key("c-v"),
    "edit region": Key("c-w") + utils.RunApp("notepad") + Key("c-v"),

    # General
    "exec": Key("a-x"),
    "helm": Key("c-x, c"),
    "helm resume": Key("c-x, c, b"),
    "preelin": Key("a-p"),
    "nollin": Key("a-n"),
    "prefix": Key("c-u"),
    "quit": Key("q"),
    "refresh": Key("g"),
    "open link": Key("c-c, c, u/25") + OpenClipboardUrlAction(),

    # Emacs
    "help variable": Key("c-h, v"),
    "help function": Key("c-h, f"),
    "help key": Key("c-h, k"),
    "help mode": Key("c-h, m"),
    "help back": Key("c-c, c-b"),
    "customize": Exec("customize-apropos"),
    "kill emacs server": Exec("ws-stop-all"),

    # Undo
    "nope": Key("c-g"),
    "no way": Key("c-g/5:3"),
    "(shuck|undo)": Key("c-slash"),
    "redo": Key("c-question"),

    # Window manipulation
    "split fub": Key("c-x, 3"),
    "clote fub": Key("c-x, 0"),
    "done fub": Key("c-x, hash"),
    "only fub": Key("c-x, 1"),
    "other fub": Key("c-x, o"),
    "die fub": Key("c-x, k"),
    "even fub": Key("c-x, plus"),
    "up fub": Exec("windmove-up"),
    "down fub": Exec("windmove-down"),
    "left fub": Exec("windmove-left"),
    "right fub": Exec("windmove-right"),
    "split header": Key("c-x, 3, c-x, o, c-x, c-h"),

    # Filesystem
    "save": Key("c-x, c-s"),
    "save as": Key("c-x, c-w"),
    "save all": Key("c-x, s"),
    "save all now": Key("c-u, c-x, s"),
    "buff": Key("c-x, b"),
    "oaf|oafile": Key("c-x, c-f"),
    "no ido": Key("c-f"),
    "dear red": Key("c-d"),
    "project file": Key("c-c, p, f"),
    "simulator file": Key("c-c, c, p, s"),
    "switch project": Key("c-c, p, p"),
    "swap project": Key("c-c, s"),
    "next result": Key("a-comma"),
    "open (definition|def)": Key("c-backtick, c-c, comma, d"),
    "toggle (definition|def)": Key("c-c, comma, D"),
    "open cross (references|ref)": Key("c-c, comma, x"),
    "open tag": Key("a-dot, enter"),
    "R grep": Exec("rgrep"),
    "code search": Exec("cs"),
    "code search car": Exec("csc"),

    # Bookmarks
    "open bookmark": Key("c-x, r, b"),
    "(save|set) bookmark": Key("c-x, r, m"),
    "list bookmarks": Key("c-x, r, l"),

    # Movement
    "furred [<n>]": Key("a-f/5:%(n)d"),
    "bird [<n>]": Key("a-b/5:%(n)d"),
    "nasper": Key("ca-f"),
    "pesper": Key("ca-b"),
    "dowsper": Key("ca-d"),
    "usper": Key("ca-u"),
    "fopper": Key("c-c, c, c-f"),
    "bapper": Key("c-c, c, c-b"),
    "white": Key("a-m"),
    "full line <line>": Key("a-g, a-g") + Text("%(line)s") + Key("enter"),
    "line <n1>": jump_to_line("%(n1)s"),
    "re-center": Key("c-l"),
    "set mark": Key("c-backtick"),
    "jump mark": Key("c-langle"),
    "jump change": Key("c-c, c, c"),
    "jump symbol": Key("a-i"),
    "swap mark": Key("c-c, c-x"),
    "(prev|preev) [<n>]": Key("c-r/5:%(n)d"),
    "next [<n>]": Key("c-s/5:%(n)d"),
    "edit search": Key("a-e"),
    "word search": Key("a-s, w"),
    "symbol search": Key("a-s, underscore"),
    "regex search": Key("ca-s"),
    "occur": Key("a-s, o"),
    "(prev|preev) symbol": Key("a-s, dot, c-r, c-r"),
    "(next|neck) symbol": Key("a-s, dot, c-s"),
    "before [preev] <char>": Key("c-c, c, b") + Text("%(char)s"),
    "after [next] <char>": Key("c-c, c, f") + Text("%(char)s"),
    "before next <char>": Key("c-c, c, s") + Text("%(char)s"),
    "after preev <char>": Key("c-c, c, e") + Text("%(char)s"),
    "other top": Key("c-minus, ca-v"),
    "other pown": Key("ca-v"),
    "I jump <char>": Key("c-c, c, j") + Text("%(char)s") + Function(lambda: eye_tracker.type_position("%d\n%d\n")),

    # Editing
    "slap above": Key("a-enter"),
    "slap below": Key("c-enter"),
    "move (line|lines) up [<n>]": Key("c-u") + Text("%(n)d") + Key("a-up"),
    "move (line|lines) down [<n>]": Key("c-u") + Text("%(n)d") + Key("a-down"),
    "copy (line|lines) up [<n>]": Key("c-u") + Text("%(n)d") + Key("as-up"),
    "copy (line|lines) down [<n>]": Key("c-u") + Text("%(n)d") + Key("as-down"),
    "clear line": Key("c-a, c-c, c, k"),
    "join (line|lines)": Key("as-6"),
    "open line <n1>": jump_to_line("%(n1)s") + Key("a-enter"),
    "select region": Key("c-x, c-x"),
    "indent region": Key("ca-backslash"),
    "comment region": Key("a-semicolon"),
    "(clang format|format region)": Key("ca-q"),
    "format <n1> [(through|to) <n2>]": MarkLinesAction() + Key("ca-q"),
    "format comment": Key("a-q"),
    "kurd [<n>]": Key("a-d/5:%(n)d"),
    "replace": Key("as-5"),
    "regex replace": Key("cas-5"),
    "replace symbol": Key("a-apostrophe"),
    "narrow region": Key("c-x, n, n"),
    "widen buffer": Key("c-x, n, w"),
    "cut": Key("c-w"),
    "copy": Key("a-w"),
    "yank": Key("c-y"),
    "yank <n1> [(through|to) <n2>]": UseLinesAction(Key("a-w"), Key("c-y")),
    "yank tight <n1> [(through|to) <n2>]": UseLinesAction(Key("a-w"), Key("c-y"), True),
    "grab <n1> [(through|to) <n2>]": UseLinesAction(Key("c-w"), Key("c-y")),
    "grab tight <n1> [(through|to) <n2>]": UseLinesAction(Key("c-w"), Key("c-y"), True),
    "copy <n1> [(through|to) <n2>]": MarkLinesAction() + Key("a-w"),
    "copy tight <n1> [(through|to) <n2>]": MarkLinesAction(True) + Key("c-w"),
    "cut <n1> [(through|to) <n2>]": MarkLinesAction() + Key("c-w"),
    "cut tight <n1> [(through|to) <n2>]": MarkLinesAction(True) + Key("c-w"),
    "sank": Key("a-y"),
    "Mark": Key("c-space"),
    "Mark <n1> [(through|to) <n2>]": MarkLinesAction(),
    "Mark tight <n1> [(through|to) <n2>]": MarkLinesAction(True),
    "moosper": Key("cas-2"),
    "kisper": Key("ca-k"),
    "expand region": Key("c-equals"),
    "contract region": Key("c-plus"),
    "surround parens": Key("a-lparen"),
    "close tag": Key("c-c, c-e"),

    # Registers
    "set mark (reg|rej) <char>": Key("c-x, r, space, %(char)s"),
    "save mark [(reg|rej)] <char>": Key("c-c, c, m, %(char)s"),
    "jump mark (reg|rej) <char>": Key("c-x, r, j, %(char)s"),
    "copy (reg|rej) <char>": Key("c-x, r, s, %(char)s"),
    "save copy [(reg|rej)] <char>": Key("c-c, c, w, %(char)s"),
    "yank (reg|rej) <char>": Key("c-u, c-x, r, i, %(char)s"),

    # Templates
    "plate <template>": Key("c-c, ampersand, c-s") + Text("%(template)s") + Key("enter"),
    "open (snippet|template) <template>": Key("c-c, ampersand, c-v") + Text("%(template)s") + Key("enter"),
    "open (snippet|template)": Key("c-c, ampersand, c-v"),
    "new (snippet|template)": Key("c-c, ampersand, c-n"),
    "reload (snippets|templates)": Exec("yas-reload-all"),

    # Compilation
    "build file": Key("c-c/10, c-g"),
    "test file": Key("c-c, c-t"),
    "preev error": Key("f11"),
    "next error": Key("f12"),
    "recompile": Exec("recompile"),

    # Dired
    "toggle details": Exec("dired-hide-details-mode"),

    # Web editing
    "JavaScript mode": Exec("js-mode"),
    "HTML mode": Exec("html-mode"),

    # C++
    "header": Key("c-x, c-h"),
    "copy import": Key("f5"),
    "paste import": Key("f6"),

    # Python
    "pie flakes": Key("c-c, c-v"),

    # Shell
    "create shell": Exec("shell"),
    "durr shell": Key("c-c, c, dollar"),

    # Clojure
    "closure compile": Key("c-c, c-k"),
    "closure namespace": Key("c-c, a-n"),

    # Lisp
    "eval defun": Key("ca-x"),
    "eval region": Exec("eval-region"),

    # Version control
    "magit status": Key("c-c, m"),
    "expand diff": Key("a-4"),
    "submit comment": Key("c-c, c-c"),
    "show diff": Key("c-x, v, equals"),
}

emacs_terminal_action_map = {
    "boof <custom_text>": Key("c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
    "ooft <custom_text>": Key("left, c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("c-s, enter"),
    "baif <custom_text>": Key("right, c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("c-r, enter"),
    "aift <custom_text>": Key("c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
}

templates = {
    "beginend": "beginend",
    "car": "car",
    "class": "class",
    "const ref": "const_ref",
    "const pointer": "const_pointer",
    "def": "function",
    "each": "each",
    "else": "else",
    "entry": "entry",
    "error": "error",
    "eval": "eval",
    "fatal": "fatal",
    "for": "for",
    "fun declaration": "fun_declaration",
    "function": "function",
    "if": "if",
    "info": "info",
    "inverse if": "inverse_if",
    "key": "key",
    "map": "map",
    "method": "method",
    "ref": "ref",
    "set": "set",
    "shared pointer": "shared_pointer",
    "ternary": "ternary",
    "text": "text",
    "to do": "todo",
    "unique pointer": "unique_pointer",
    "var": "vardef",
    "vector": "vector",
    "warning": "warning",
    "while": "while",
}
template_dict_list = DictList("template_dict_list", templates)
emacs_element_map = {
    "n1": IntegerRef(None, 0, 100),
    "n2": IntegerRef(None, 0, 100),
    "line": IntegerRef(None, 1, 10000),
    "template": DictListRef(None, template_dict_list),
}

emacs_environment = MyEnvironment(name="Emacs",
                                  parent=global_environment,
                                  context=linux.UniversalAppContext(title = "Emacs editor"),
                                  action_map=emacs_action_map,
                                  terminal_action_map=emacs_terminal_action_map,
                                  element_map=emacs_element_map)


### Emacs: Python

emacs_python_action_map = {
    "[python] indent": Key("c-c, rangle"),
    "[python] dedent": Key("c-c, langle"),
}
emacs_python_environment = MyEnvironment(name="EmacsPython",
                                         parent=emacs_environment,
                                         context=linux.UniversalAppContext(title="- Python -"),
                                         action_map=emacs_python_action_map)


### Emacs: Org-Mode

emacs_org_action_map = {
    "new heading above": Key("c-a, a-enter"),
    "new heading": Key("c-e, a-enter"),
    "brand new heading": Key("c-e, a-enter, c-c, c, a-left"),
    "new heading below": Key("c-e, c-enter"),
    "subheading": Key("c-e, a-enter, a-right"),
    "split heading": Key("a-enter"),
    "new to do above": Key("c-a, as-enter"),
    "new to do": Key("c-e, as-enter"),
    "brand new to do": Key("c-e, as-enter, c-c, c, a-left"),
    "new to do below": Key("c-e, cs-enter"),
    "sub to do": Key("c-e, as-enter, a-right"),
    "split to do": Key("as-enter"),
    "toggle heading": Key("c-c, asterisk"),
    "to do": Key("c-1, c-c, c-t"),
    "done": Key("c-2, c-c, c-t"),
    "clear to do": Key("c-3, c-c, c-t"),
    "indent tree": Key("as-right"),
    "indent": Key("a-right"),
    "dedent tree": Key("as-left"),
    "dedent": Key("a-left"),
    "move tree down": Key("as-down"),
    "move tree up": Key("as-up"),
    "open org link": Key("c-c, c-o"),
    "show to do's": Key("c-c, slash, t"),
    "archive": Key("c-c, c-x, c-a"),
    "org (West|white)": Key("c-c, c, c-a"),
    "tag <tag>": Key("c-c, c-q") + Text("%(tag)s") + Key("enter"),
}
tags = {
    "new": "new",
    "Q1": "q1",
    "Q2": "q2",
    "Q3": "q3",
    "Q4": "q4",
    "low": "low",
    "high": "high",
}
tag_dict_list = DictList("tag_dict_list", tags)
emacs_org_element_map = {
    "tag": DictListRef(None, tag_dict_list),
}
emacs_org_environment = MyEnvironment(name="EmacsOrg",
                                      parent=emacs_environment,
                                      context=linux.UniversalAppContext(title="- Org -"),
                                      action_map=emacs_org_action_map,
                                      element_map=emacs_org_element_map)


### Emacs: Shell

emacs_shell_action_map = utils.combine_maps(
    shell_command_map,
    {
        "shell (preev|back)": Key("a-r"),
        "show output": Key("c-c, c-r"),
    })
emacs_shell_environment = MyEnvironment(name="EmacsShell",
                                        parent=emacs_environment,
                                        context=linux.UniversalAppContext(title="- Shell -"),
                                        action_map=emacs_shell_action_map)


### Shell

shell_action_map = utils.combine_maps(
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
        "shot <tab_n>": Key("a-%(tab_n)d"),
        "shot last": Key("a-1, cs-left"),
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

shell_element_map = {
    "tab_n": IntegerRef(None, 1, 10),
}

shell_environment = MyEnvironment(name="Shell",
                                  parent=global_environment,
                                  context=linux.UniversalAppContext(title=" - Terminal"),
                                  action_map=shell_action_map,
                                  element_map=shell_element_map)


### Chrome

chrome_action_map = {
    "link": Key("c-comma"),
    "new link": Key("c-dot"),
    "background links": Key("a-f"),
    "new tab":            Key("c-t"),
    "new incognito":            Key("cs-n"),
    "new window": Key("c-n"),
    "clote":          Key("c-w"),
    "address bar":        Key("c-l"),
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
    "(caret|carrot) browsing": Key("f7"),
    "moma": Key("c-l/15") + Text("moma") + Key("tab"),
    "code search car": Key("c-l/15") + Text("csc") + Key("tab"),
    "code search simulator": Key("c-l/15") + Text("css") + Key("tab"),
    "code search": Key("c-l/15") + Text("cs") + Key("tab"),
    "go to calendar": Key("c-l/15") + Text("calendar.google.com") + Key("enter"),
    "go to critique": Key("c-l/15") + Text("cr/") + Key("enter"),
    "go to (buganizer|bugs)": Key("c-l/15") + Text("b/") + Key("enter"),
    "go to presubmits": Key("c-l/15, b, tab") + Text("One shot") + Key("enter:2"),
    "go to postsubmits": Key("c-l/15, b, tab") + Text("Continuous") + Key("enter:2"),
    "go to latest test results": Key("c-l/15, b, tab") + Text("latest test results") + Key("enter:2"),
    "go to docs": Key("c-l/15") + Text("docs.google.com") + Key("enter"),
    "go to slides": Key("c-l/15") + Text("slides.google.com") + Key("enter"),
    "go to sheets": Key("c-l/15") + Text("sheets.google.com") + Key("enter"),
    "go to new doc": Key("c-l/15") + Text("go/newdoc") + Key("enter"),
    "go to new (slide|slides)": Key("c-l/15") + Text("go/newslide") + Key("enter"),
    "go to new sheet": Key("c-l/15") + Text("go/newsheet") + Key("enter"),
    "go to new script": Key("c-l/15") + Text("go/newscript") + Key("enter"),
    "go to drive": Key("c-l/15") + Text("drive.google.com") + Key("enter"),
    "go to amazon": Key("c-l/15") + Text("smile.amazon.com") + Key("enter"),
    "strikethrough": Key("as-5"),
    "bullets": Key("cs-8"),
    "bold": Key("c-b"),
    "create link": Key("c-k"),
    "text box": Key("a-i/15, t"),
    "paste raw": Key("cs-v"),
    "next match": Key("c-g"),
    "preev match": Key("cs-g"),
    "(go to|open) bookmark": Key("c-semicolon"),
    "new bookmark": Key("c-apostrophe"),
    "save bookmark": Key("c-d"),
    "next frame": Key("c-lbracket"),
    "developer tools": Key("cs-j"),
    "test driver": Function(webdriver.test_driver),
    "search bar": webdriver.ClickElementAction(By.NAME, "q"),
    "add bill": webdriver.ClickElementAction(By.LINK_TEXT, "Add a bill"),
}

chrome_terminal_action_map = {
    "search <text>":        Key("c-l/15") + Text("%(text)s") + Key("enter"),
}

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
chrome_element_map = {
    "tab_n": IntegerRef(None, 1, 9),
    "link": utils.JoinedRepetition(
        "", DictListRef(None, link_char_dict_list), min=0, max=5),
}

chrome_environment = MyEnvironment(name="Chrome",
                                   parent=global_environment,
                                   context=AppContext(title=" - Google Chrome"),
                                   action_map=chrome_action_map,
                                   terminal_action_map=chrome_terminal_action_map,
                                   element_map=chrome_element_map)


### Chrome: Amazon

amazon_action_map = {
    "search bar": webdriver.ClickElementAction(By.NAME, "field-keywords"),
}

amazon_environment = MyEnvironment(name="Amazon",
                                   parent=chrome_environment, 
                                   context=(AppContext(title="<www.amazon.com>") |
                                            AppContext(title="<smile.amazon.com>")),
                                   action_map=amazon_action_map)


### Chrome: Critique

critique_action_map = {
    "preev": Key("p"),
    "next": Key("n"),
    "preev comment": Key("P"),
    "next comment": Key("N"),
    "preev file": Key("k"),
    "next file": Key("j"),
    "open": Key("o"),
    "list": Key("u"),
    "comment": Key("c"),
    "resolve": Key("c-j"),
    "done": Key("d"),
    "save": Key("c-s"),
    "expand|collapse": Key("e"),
    "reply": Key("r"),
    "comment <line_n>": webdriver.DoubleClickElementAction(
        By.XPATH, ("//span[contains(@class, 'stx-line') and "
                   "starts-with(@id, 'c') and "
                   "substring-after(@id, '_') = '%(line_n)s']")),
    "click LGTM": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='LGTM']"),
    "click action required": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Action required']"),
    "click send": webdriver.ClickElementAction(By.XPATH, "//*[starts-with(@aria-label, 'Send')]"),
    "search bar": Key("slash"),
}
critique_element_map = {
    "line_n": IntegerRef(None, 1, 10000),
}
critique_environment = MyEnvironment(name="Critique",
                                     parent=chrome_environment,
                                     context=AppContext(title="<critique.corp.google.com>"),
                                     action_map=critique_action_map,
                                     element_map=critique_element_map)


### Chrome: Calendar

calendar_action_map = {
    "click <name>": webdriver.ClickElementAction(
        By.XPATH, "//*[@role='option' and contains(string(.), '%(name)s')]"),
    "today": Key("t"),
    "preev": Key("k"),
    "next": Key("j"),
    "day": Key("d"),
    "week": Key("w"),
    "month": Key("m"),
    "agenda": Key("a"), 
}
names_dict_list = DictList(
    "name_dict_list",
    {
        "Sonica": "Sonica"
    })
calendar_element_map = {
    "name": DictListRef(None, names_dict_list),
}
calendar_environment = MyEnvironment(name="Calendar",
                                     parent=chrome_environment,
                                     context=(AppContext(title="Google Calendar") |
                                              AppContext(title="Google.com - Calendar")),
                                     action_map=calendar_action_map,
                                     element_map=calendar_element_map)


### Chrome: Code search

code_search_action_map = {
    "header": Key("r/25, h"),
    "source": Key("r/25, c"),
    "search bar": Key("slash"),
}
code_search_environment = MyEnvironment(name="CodeSearch",
                                        parent=chrome_environment,
                                        context=AppContext(title="<cs.corp.google.com>"),
                                        action_map=code_search_action_map)


### Chrome: Gmail

gmail_action_map = {
    "open": Key("o"),
    "archive": Text("{"),
    "done": Text("["),
    "mark unread": Text("_"),
    "undo": Key("z"),
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
    "mark starred": Key("s"),
    "next section": Key("backtick"),
    "preev section": Key("tilde"),
    "not important|don't care": Key("minus"),
    "label waiting": Key("l/50") + Text("waiting") + Key("enter"),
    "label snooze": Key("l/50") + Text("snooze") + Key("enter"),
    "label candidates": Key("l/50") + Text("candidates") + Key("enter"),
    "check": Key("x"),
    "check next <n>": Key("x, j") * Repeat(extra="n"),
    "new messages": Key("N"),
    "go to inbox": Key("g, i"),
    "go to starred": Key("g, s"),
    "go to sent": Key("g, t"),
    "go to drafts": Key("g, d"),
    "expand all": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Expand all']"),
    "click to": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='To']"),
    "click cc": Key("cs-c"),
    "open chat": Key("q"),
    "send mail": Key("c-enter"),
}
gmail_terminal_action_map = {
    "chat with <text>": Key("q/50") + Text("%(text)s") + Pause("50") + Key("enter"),
}

gmail_environment = MyEnvironment(name="Gmail",
                                  parent=chrome_environment,
                                  context=(AppContext(title="Gmail") |
                                           AppContext(title="Google.com Mail") |
                                           AppContext(title="<mail.google.com>") |
                                           AppContext(title="<inbox.google.com>")),
                                  action_map=gmail_action_map,
                                  terminal_action_map=gmail_terminal_action_map)


### Chrome: docs

docs_action_map = {
    "select column": Key("c-space:2"),
    "select row": Key("s-space:2"),
    "row up": Key("a-e/15, k"),
    "row down": Key("a-e/15, j"),
    "column left": Key("a-e/15, m"),
    "column right": Key("a-e/15, m"),
    "add comment": Key("ca-m"),
    "(previous|preev) comment": Key("ctrl:down, alt:down, p, c, ctrl:up, alt:up"),
    "next comment": Key("ctrl:down, alt:down, n, c, ctrl:up, alt:up"),
    "enter comment": Key("ctrl:down, alt:down, e, c, ctrl:up, alt:up"), 
    "(new|insert) row above": Key("a-i/15, r"),
    "(new|insert) row [below]": Key("a-i/15, b"),
    "duplicate row": Key("s-space:2, c-c/15, a-i/15, b, c-v/30, up/30, down"),
    "delete row": Key("a-e/15, d"),
}
docs_environment = MyEnvironment(name="Docs",
                                 parent=chrome_environment,
                                 context=AppContext(title="<docs.google.com>"),
                                 action_map=docs_action_map)


### Chrome: Buganizer

buganizer_action_map = {}
run_local_hook("AddBuganizerCommands", buganizer_action_map)
buganizer_environment = MyEnvironment(name="Buganizer",
                                      parent=chrome_environment,
                                      context=(AppContext(title="Buganizer V2") |
                                               AppContext(title="<b.corp.google.com>") |
                                               AppContext(title="<buganizer.corp.google.com>") |
                                               AppContext(title="<b2.corp.google.com>")),
                                      action_map=buganizer_action_map)


### Chrome: Analog

analog_action_map = {
    "next": Key("n"),
    "preev": Key("p"),
}
analog_environment = MyEnvironment(name="Analog",
                                   parent=chrome_environment,
                                   context=AppContext(title="<analog.corp.google.com>"),
                                   action_map=analog_action_map)


### Notepad

notepad_action_map = {
    "dumbbell [<n>]": Key("shift:down, c-left/5:%(n)d, backspace, shift:up"),
    "transfer out": Key("c-a, c-x, a-f4") + utils.UniversalPaste(),
}

notepad_environment = MyEnvironment(name="Notepad",
                                    parent=global_environment,
                                    context=AppContext(executable = "notepad"),
                                    action_map=notepad_action_map)


### Linux

# TODO Figure out either how to integrate this with the repeating rule or move out.
linux_action_map = utils.combine_maps(
    {
        "create terminal": Key("ca-t"),
        "go to Emacs": linux.ActivateLinuxWindow("Emacs editor"),
        "go to terminal": linux.ActivateLinuxWindow(" - Terminal"),
        "go to Firefox": linux.ActivateLinuxWindow("Mozilla Firefox"),
    })
run_local_hook("AddLinuxCommands", linux_action_map)
linux_rule = utils.create_rule("LinuxRule", linux_action_map, {}, True,
                               (AppContext(title="Oracle VM VirtualBox") |
                                AppContext(title=" - Chrome Remote Desktop")))


#-------------------------------------------------------------------------------
# Populate and load the grammar.

grammar = Grammar("repeat")   # Create this module's grammar.
global_environment.install(grammar)
# TODO Figure out either how to integrate this with the repeating rule or move out.
grammar.add_rule(linux_rule)
grammar.load()


#-------------------------------------------------------------------------------
# Start a server which lets Emacs send us nearby text being edited, so we can
# use it for contextual recognition. Note that the server is only exposed to
# the local machine, except it is still potentially vulnerable to CSRF
# attacks. Consider this when adding new functionality.

# Register timer to run arbitrary callbacks added by the server.
callbacks = Queue.Queue()


def RunCallbacks():
    global callbacks
    while not callbacks.empty():
        callbacks.get_nowait()()


timer = get_engine().create_timer(RunCallbacks, 0.1)


# Update the context words and phrases.
def UpdateWords(words, phrases):
    context_word_list.set(words)
    context_phrase_list.set(phrases)


class TextRequestHandler(BaseHTTPServer.BaseHTTPRequestHandler):
    """HTTP handler for receiving a block of text. The file type can be provided
    with header My-File-Type.
    TODO: Use JSON instead of custom headers and dispatch based on path.
    """

    def do_POST(self):
    # Uncomment to enable profiling.
    #     cProfile.runctx("self.PostInternal()", globals(), locals())
    # def PostInternal(self):
        start_time = time.time()
        length = self.headers.getheader("content-length")
        file_type = self.headers.getheader("My-File-Type")
        request_text = self.rfile.read(int(length)) if length else ""
        # print("received text: %s" % request_text)
        words = text.extract_words(request_text, file_type)
        phrases = text.extract_phrases(request_text, file_type)
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
server_thread = threading.Thread(target=server.serve_forever)
server_thread.start()
print("started server")

# Connect to Chrome WebDriver if possible.
webdriver.create_driver()

# Connect to eye tracker if possible.
eye_tracker.connect()


#-------------------------------------------------------------------------------
# Unload function which will be called by NatLink.
def unload():
    global grammar, server, server_thread, timer
    if grammar:
        grammar.unload()
        grammar = None
    eye_tracker.disconnect()
    webdriver.quit_driver()
    timer.stop()
    server.shutdown()
    server_thread.join()
    print("shutdown server")
