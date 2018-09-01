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
import re
import socket
import threading
import time
import webbrowser
import win32clipboard

from dragonfly import (
    ActionBase,
    Alternative,
    AppContext,
    Choice,
    Compound,
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
    MappingRule,
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
from dragonfly import a11y
from dragonfly.a11y import utils as a11y_utils
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
# symbols_map = {
#     "plus": "+",
#     "plus twice": "++",
#     "minus": "-",
#     ",": ",",
#     "colon": ":",
#     "equals": "=",
#     "equals twice": "==",
#     "not equals": "!=",
#     "plus equals": "+=",
#     "greater than": ">",
#     "less than": "<",
#     "greater equals": ">=",
#     "less equals": "<=",
#     "dot": ".",
#     "leap": "(",
#     "reap": ")",
#     "lake": "{",
#     "rake": "}",
#     "lobe": "[",
#     "robe": "]",
#     "luke": "<",
#     "luke twice": "<<",
#     "ruke": ">",
#     "ruke twice": ">>",
#     "quote": "\"",
#     "dash": "-",
#     "semi": ";",
#     "bang": "!",
#     "percent": "%",
#     "star": "*",
#     "backslash": "\\",
#     "slash": "/",
#     "tilde": "~",
#     "backtick": "`",
#     "underscore": "_",
#     "single quote": "'",
#     "dollar": "$",
#     "carrot": "^",
#     "arrow": "->",
#     "fat arrow": "=>",
#     "colon twice": "::",
#     "amper": "&",
#     "amper twice": "&&",
#     "pipe": "|",
#     "pipe twice": "||",
#     "hash": "#",
#     "at sign": "@",
#     "question": "?",
# }

# symbols_map = {
#     "plus [sign]": "+",
#     "plus [sign] twice": "++",
#     "minus|hyphen|dash": "-",
#     ",": ",",
#     "colon": ":",
#     "equals [sign]": "=",
#     "equals [sign] twice": "==",
#     "not equals": "!=",
#     "plus equals": "+=",
#     "greater than|ruke|close angle": ">",
#     "less than|luke|open angle": "<",
#     "(greater than|ruke|close angle) twice": ">>",
#     "(less than|luke|open angle) twice": "<<",
#     "greater equals": ">=",
#     "less equals": "<=",
#     "dot|period": ".",
#     "leap|open paren": "(",
#     "reap|close paren": ")",
#     "lake|open brace": "{",
#     "rake|close brace": "}",
#     "lobe|open bracket": "[",
#     "robe|close bracket": "]",
#     "quote|open quote|close quote": "\"",
#     "semi|semicolon": ";",
#     "bang|exclamation mark": "!",
#     "percent [sign]": "%",
#     "star|asterisk": "*",
#     "backslash": "\\",
#     "slash": "/",
#     "tilde": "~",
#     "backtick": "`",
#     "underscore": "_",
#     "single quote|apostrophe": "'",
#     "dollar [sign]": "$",
#     "caret": "^",
#     "arrow": "->",
#     "fat arrow": "=>",
#     "colon twice": "::",
#     "amper|ampersand": "&",
#     "(amper|ampersand) twice": "&&",
#     "pipe": "|",
#     "pipe twice": "||",
#     "hash|number sign": "#",
#     "at sign": "@",
#     "question [mark]": "?",
# }

symbols_map = {
    "plus": "+",
    "plus sign": "+",
    "plus twice": "++",
    "plus sign twice": "++",
    "minus": "-",
    "hyphen": "-",
    "dash": "-",
    ",": ",",
    "colon": ":",
    "equals": "=",
    "equals sign": "=",
    "equals twice": "==",
    "equals sign twice": "==",
    "not equals": "!=",
    "plus equals": "+=",
    "minus equals": "-=",
    "greater than": ">",
    "ruke": ">",
    "close angle": ">",
    "less than": "<",
    "luke": "<",
    "open angle": "<",
    "greater than twice": ">>",
    "ruke twice": ">>",
    "close angle twice": ">>",
    "less than twice": "<<",
    "luke twice": "<<",
    "open angle twice": "<<",
    "greater equals": ">=",
    "less equals": "<=",
    "dot": ".",
    "period": ".",
    "leap": "(",
    "open paren": "(",
    "reap": ")",
    "close paren": ")",
    "lake": "{",
    "open brace": "{",
    "rake": "}",
    "close brace": "}",
    "lobe": "[",
    "open bracket": "[",
    "robe": "]",
    "close bracket": "]",
    "quote": "\"",
    "open quote": "\"",
    "close quote": "\"",
    "semi": ";",
    "semicolon": ";",
    "bang": "!",
    "exclamation mark": "!",
    "percent": "%",
    "percent sign": "%",
    "star": "*",
    "asterisk": "*",
    "backslash": "\\",
    "slash": "/",
    "tilde": "~",
    "backtick": "`",
    "underscore": "_",
    "underscore twice": "__",
    "dunder": "__",
    "single quote": "'",
    "apostrophe": "'",
    "dollar": "$",
    "dollar sign": "$",
    "caret": "^",
    "arrow": "->",
    "fat arrow": "=>",
    "colon twice": "::",
    "amper": "&",
    "ampersand": "&",
    "amper twice": "&&",
    "ampersand twice": "&&",
    "pipe": "|",
    "pipe twice": "||",
    "hash": "#",
    "number sign": "#",
    "at sign": "@",
    "question": "?",
    "question mark": "?",
}

symbol_keys_map = {
    "minus|dash": "-",
    ",": ",",
    "equals": "=",
    "dot|period": ".",
    "semi": ";",
    "backslash": "\\",
    "slash": "/",
    "backtick": "`",
    "single quote": "'",
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
    "colon": ":",
    ",": ",",
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

# Disabled for efficiency.
# long_letters_map = {
#     "alpha": "a",
#     "bravo": "b",
#     "charlie": "c",
#     "delta": "d",
#     "echo": "e",
#     "foxtrot": "f",
#     "golf": "g",
#     "hotel": "h",
#     "india": "i",
#     "juliet": "j",
#     "kilo": "k",
#     "lima": "l",
#     "mike": "m",
#     "november": "n",
#     "oscar": "o",
#     "poppa": "p",
#     "quebec": "q",
#     "romeo": "r",
#     "sierra": "s",
#     "tango": "t",
#     "uniform": "u",
#     "victor": "v",
#     "whiskey": "w",
#     "x-ray": "x",
#     "yankee": "y",
#     "zulu": "z",
#     "dot": ".",
# }

prefixes = [
    "num",
    "min",
]

suffixes = [
    "bytes",
]

letters_map = utils.combine_maps(quick_letters_map)

chars_map = utils.combine_maps(letters_map, numbers_map, symbols_map)

# Load commonly misrecognized words saved to a file.
# TODO: Revisit.
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

a11y_controller = a11y.GetA11yController()

dictation_key_action_map = {
    "enter|slap": Key("enter"),
    "space|spooce|spacebar": Key("space"),
    "tab-key": Key("tab"),
}

dictation_action_map = utils.combine_maps(dictation_key_action_map,
                                          utils.text_map_to_action_map(symbols_map))

standalone_key_action_map = utils.combine_maps(
    dictation_key_action_map,
    {
        "up": Key("up"),
        "down": Key("down"),
        "left": Key("left"),
        "right": Key("right"),
        "page up": Key("pgup"),
        "page down": Key("pgdown"),
        "apps key": Key("apps"),
        "escape": Key("escape"),
        "backspace": Key("backspace"),
        "delete key": Key("del"),
    })

full_key_action_map = utils.combine_maps(
    standalone_key_action_map,
    utils.text_map_to_action_map(utils.combine_maps(letters_map, numbers_map, symbol_keys_map)),
    {
        "home": Key("home"),
        "end": Key("end"),
        "tab": Key("tab"),
        "delete": Key("del"),
    })

repeatable_action_map = utils.combine_maps(
    standalone_key_action_map,
    {
        "after": Key("c-right"),
        "before": Key("c-left"),
        "afters": Key("shift:down, c-right, shift:up"),
        "befores": Key("shift:down, c-left, shift:up"),
        "ahead": Key("a-f"),
        "behind": Key("a-b"),
        "aheads": Key("shift:down, a-f, shift:up"),
        "behinds": Key("shift:down, a-b, shift:up"),
        "rights": Key("shift:down, right, shift:up"),
        "lefts": Key("shift:down, left, shift:up"),
        "kill": Key("c-k"),
        "screen up": Key("pgup"),
        "screen down": Key("pgdown"),
        "cancel": Key("escape"),
    })


def select_words(text):
    selection_points = a11y_utils.get_text_selection_points(a11y_controller, str(text))
    Mouse("[%d, %d], left:down, [%d, %d]/10, left:up" % (selection_points[0][0], selection_points[0][1],
                                                         selection_points[1][0], selection_points[1][1])).execute()


def replace_words(text, replacement):
    cursor_before = a11y_utils.get_cursor_offset(a11y_controller)
    select_words(text)
    # TODO Add escaping.
    Text(str(replacement)).execute()
    # TODO Use actual selection length, not search phrase length.
    if cursor_before:
        a11y_utils.set_cursor_offset(a11y_controller, cursor_before + len(str(replacement)) - len(str(text)))

# Actions of commonly used text navigation and mousing commands. These can be
# used anywhere except after commands which include arbitrary dictation.
# TODO: Better solution for holding shift during a single command. Think about whether this could enable a simpler grammar for other modifiers.
command_action_map = utils.combine_maps(
    utils.text_map_to_action_map(symbols_map),
    {
        "delete": Key("del"),
        "go home|[go] west": Key("home"),
        "go end|[go] east": Key("end"),
        "go top|[go] north": Key("c-home"),
        "go bottom|[go] south": Key("c-end"),
        # These work like the built-in commands and are available for any
        # application that supports IAccessible2.
        "my go before <text>": Function(lambda text: a11y_utils.move_cursor(a11y_controller, str(text), before=True)),
        "my go after <text>": Function(lambda text: a11y_utils.move_cursor(a11y_controller, str(text), before=False)),
        "my words <text>": Function(select_words),
        "volume [<n>] up": Key("volumeup/5:%(n)d"),
        "volume [<n>] down": Key("volumedown/5:%(n)d"),
        "volume (mute|unmute)": Key("volumemute"),
        "track next": Key("tracknext"),
        "track preev": Key("trackprev"),
        "track (pause|play)": Key("playpause"),

        "paste": Key("c-v"),
        "copy": Key("c-c"),
        "cut": Key("c-x"),
        "all select":                       Key("c-a"),
        "here edit": utils.RunApp("notepad"),
        "all edit": Key("c-a, c-x") + utils.RunApp("notepad") + Key("c-v"),
        "this edit": Key("c-x") + utils.RunApp("notepad") + Key("c-v"),
        "shift hold":                     Key("shift:down"),
        "shift release":                    Key("shift:up"),
        "control hold":                   Key("ctrl:down"),
        "control release":                  Key("ctrl:up"),
        "(meta|alt) hold":                   Key("alt:down"),
        "(meta|alt) release":                  Key("alt:up"),
        "all release":                    Key("shift:up, ctrl:up, alt:up"),

        "(I|eye) connect": Function(eye_tracker.connect),
        "(I|eye) disconnect": Function(eye_tracker.disconnect),
        "(I|eye) print position": Function(eye_tracker.print_position),
        "(I|eye) move": Function(eye_tracker.move_to_position),
        "(I|eye) act": Function(eye_tracker.activate_position),
        "(I|eye) pan": Function(eye_tracker.panning_step_position),
        "(I|eye) (touch|click)": Function(eye_tracker.move_to_position) + Mouse("left"),
        "(I|eye) (touch|click) right": Function(eye_tracker.move_to_position) + Mouse("right"),
        "(I|eye) (touch|click) middle": Function(eye_tracker.move_to_position) + Mouse("middle"),
        "(I|eye) (touch|click) [left] twice": Function(eye_tracker.move_to_position) + Mouse("left:2"),
        "(I|eye) (touch|click) hold": Function(eye_tracker.move_to_position) + Mouse("left:down"),
        "(I|eye) (touch|click) release": Function(eye_tracker.move_to_position) + Mouse("left:up"),
        "scroll up": Function(lambda: eye_tracker.move_to_position((0, 50))) + Mouse("scrollup:8"),
        "scroll up half": Function(lambda: eye_tracker.move_to_position((0, 50))) + Mouse("scrollup:4"),
        "scroll down": Function(lambda: eye_tracker.move_to_position((0, -50))) + Mouse("scrolldown:8"),
        "scroll down half": Function(lambda: eye_tracker.move_to_position((0, -50))) + Mouse("scrolldown:4"),
        "(touch|click) [left]": Mouse("left"),
        "(touch|click) right": Mouse("right"),
        "(touch|click) middle": Mouse("middle"),
        "(touch|click) [left] twice": Mouse("left:2"),
        "(touch|click) hold": Mouse("left:down"),
        "(touch|click) release": Mouse("left:up"),

        "webdriver open": Function(webdriver.create_driver),
        "webdriver close": Function(webdriver.quit_driver),
    })

# Actions for speaking out sequences of characters.
character_action_map = {
    "<chars> short": Text("%(chars)s"),
    "number <numerals>": Text("%(numerals)s"),
    "letter <letters>": Text("%(letters)s"),
    "upper letter <letters>": Function(lambda letters: Text(letters.upper()).execute()),
}

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

symbols_dict_list = DictList("symbols_dict_list", symbols_map)
symbol_element = DictListRef(None, symbols_dict_list)
# symbol_element = RuleWrap(None, Choice(None, symbols_map))
numbers_dict_list = DictList("numbers_dict_list", numbers_map)
number_element = DictListRef(None, numbers_dict_list)
# number_element = RuleWrap(None, Choice(None, numbers_map))
letters_dict_list = DictList("letters_dict_list", letters_map)
letter_element = DictListRef(None, letters_dict_list)
# letter_element = RuleWrap(None, Choice(None, letters_map))
chars_dict_list = DictList("chars_dict_list", chars_map)
char_element = DictListRef(None, chars_dict_list)
# char_element = RuleWrap(None, Choice(None, chars_map))
saved_word_list = List("saved_word_list", saved_words)
# Lists which will be populated later via RPC.
context_phrase_list = List("context_phrase_list", [])
context_word_list = List("context_word_list", [])
prefix_list = List("prefix_list", prefixes)
suffix_list = List("suffix_list", suffixes)

# Either arbitrary or custom dictation.
mixed_dictation = RuleWrap(None, utils.JoinedSequence(" ", [
    Optional(ListRef(None, prefix_list)),
    Alternative([
        Dictation(),
        letter_element,
        ListRef(None, saved_word_list),
    ]),
    Optional(ListRef(None, suffix_list)),
]))

# Same as above, except no arbitrary dictation allowed.
custom_dictation = RuleWrap(None, utils.JoinedSequence(" ", [
    Optional(ListRef(None, prefix_list)),
    Alternative([
        letter_element,
        ListRef(None, saved_word_list),
    ]),
    Optional(ListRef(None, suffix_list)),
]))

# A sequence of either short letters or long letters.
letters_element = RuleWrap(None, utils.JoinedRepetition(
    "", letter_element, min=1, max=10))

# A sequence of numbers.
numbers_element = RuleWrap(None, utils.JoinedRepetition(
    "", number_element, min=1, max=10))

# A sequence of characters.
chars_element = RuleWrap(None, utils.JoinedRepetition(
    "", char_element, min=1, max=5))

# Simple element map corresponding to command action maps from earlier.
command_element_map = {
    "n": (IntegerRef(None, 1, 21), 1),
    "text": Dictation(),
    "char": char_element,
    "custom_text": RuleWrap(None, Alternative([
        Dictation(),
        chars_element,
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

# Rule for formatting symbols.
symbol_format_rule = utils.create_rule(
    "SymbolFormatRule",
    {
        "padded <symbol>": Text(" %(symbol)s "),
    },
    {
        "symbol": symbol_element,
    }
)

# Rule for formatting pure dictation elements.
pure_format_rule = utils.create_rule(
    "PureFormatRule",
    dict([("pure (" + k + ")", v)
          for (k, v) in format_functions.items()]),
    {"dictation": Dictation()}
)

# Rule for formatting custom_dictation elements.
custom_format_rule = utils.create_rule(
    "CustomFormatRule",
    dict([("my (" + k + ")", v)
          for (k, v) in format_functions.items()]),
    {"dictation": custom_dictation}
)

# Rule for handling raw dictation.
# TODO: Improve grammar.
dictation_rule = utils.create_rule(
    "DictationRule",
    {
        "(mim|mimic) text <text>": Text("%(text)s"),
        "mim small <text>": utils.uncapitalize_text_action("%(text)s"),
        "mim big <text>": utils.capitalize_text_action("%(text)s"),
        "mimic <text>": Mimic(extra="text"),
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
        "numerals": number_element,
        "letters": letter_element,
        "chars": char_element,
    }
)

# Rule for spelling a word letter by letter and formatting it.
# Disabled for efficiency.
# spell_format_rule = utils.create_rule(
#     "SpellFormatRule",
#     dict([("spell (" + k + ")", v)
#           for (k, v) in format_functions.items()]),
#     {"dictation": letters_element}
# )

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

# Element matching dictation and commands allowed at the end of an utterance.
# For efficiency, this should not contain any repeating elements. For accuracy,
# few custom commands should be included to avoid clashes with dictation
# elements.
dictation_element = RuleWrap(None, Alternative([
    RuleRef(rule=dictation_rule),
    RuleRef(rule=format_rule),
    RuleRef(rule=symbol_format_rule),
    RuleRef(rule=pure_format_rule),
    # Disabled for efficiency.
    # RuleRef(rule=custom_format_rule),
    RuleRef(rule=utils.create_rule("DictationActionRule",
                                   dictation_action_map,
                                   command_element_map)),
    # Disabled for efficiency.
    # RuleRef(rule=single_character_rule),
]))


### Final commands that can be used once after everything else. These change the
### application context so it is important that nothing else gets run after
### them.

# Ordered list of pinned taskbar items. Sublists refer to windows within a specific application.
windows = [
    "explorer",
    ["dragonbar", "dragon [messages]", "dragonpad"],
    "[home] chrome",
    "[home] terminal",
    "[home] emacs",
]
json_windows = utils.load_json("windows.json")
if json_windows:
    windows = json_windows

windows_suffix = "(win|window)"
windows_mapping = {}
for i, window in enumerate(windows):
    if isinstance(window, str):
        window = [window]
    for j, words in enumerate(window):
        windows_mapping["(" + words + ") " + windows_suffix] = Key("win:down, %d:%d/20, win:up" % (i + 1, j + 1))

final_action_map = utils.combine_maps(windows_mapping, {
    "[<n>] swap": utils.SwitchWindows("%(n)d"),
})
final_element_map = {
    "n": (IntegerRef(None, 1, 20), 1)
}
final_rule = utils.create_rule("FinalRule",
                               final_action_map,
                               final_element_map)


#-------------------------------------------------------------------------------
# System for benchmarking other commands.

class CommandBenchmark:
    def __init__(self):
        self.remaining_count = 0

    def start(self, command, repeat_count):
        if self.remaining_count > 0:
            print("Benchmark already running!")
            return
        self.repeat_count = repeat_count
        self.remaining_count = repeat_count
        self.command = command
        self.start_time = time.time()
        Mimic(*self.command.split()).execute()

    def record_and_replay_recognition(self):
        if self.remaining_count == 0:
            print("Benchmark not running!")
            return
        self.remaining_count -= 1
        if self.remaining_count == 0:
            print("Average response for command %s: %.10f" % (self.command, (time.time() - self.start_time) / self.repeat_count))
        else:
            Mimic(*self.command.split()).execute()

    def is_active(self):
        return self.remaining_count > 0


def reset_benchmark():
    global command_benchmark
    command_benchmark = CommandBenchmark()

reset_benchmark()

#---------------------------------------------------------------------------
# Here we define the top-level rule which the user can say.

# This is the rule that actually handles recognitions.
#  When a recognition occurs, its _process_recognition()
#  method will be called.  It receives information about the
#  recognition in the "extras" argument: the sequence of
#  actions and the number of times to repeat them.
class RepeatRule(CompoundRule):

    def __init__(self, name, command, repeatable_command, terminal_command):
        # Here we define this rule's spoken-form and special elements. Note that
        # nested_repetitions is the only one that contains Repetitions, and it
        # is not itself repeated. This is for performance purposes. We also
        # include a special escape command "terminal <dictation>" in case
        # recognition problems occur with repeated dictation commands.
        spec = ("[<sequence>] "
                "[<nested_repetitions>] "
                "([<dictation_sequence>] [terminal <dictation>] | <terminal_command>) "
                "[<n> times] "
                "[<final_command>]")
        repeated_command = Compound(spec="[<n>] <repeatable_command>",
                                    extras=[IntegerRef("n", 1, 21, default=1),
                                            utils.renamed_element("repeatable_command", repeatable_command)],
                                    value_func=lambda node, extras: (extras["repeatable_command"] + Pause("5")) * Repeat(extras["n"]))
        full_key_element = RuleRef(rule=utils.create_rule("full_key_rule", full_key_action_map, {}), name="single_key")
        combo_key_element = Compound(spec="[<n>] <modifier> <single_key>",
                                     extras=[IntegerRef("n", 1, 21, default=1),
                                             RuleWrap("modifier", Choice(None, {
                                                 "control": lambda action: Key("ctrl:down") + action + Key("ctrl:up"),
                                                 "alt|meta|under": lambda action: Key("alt:down") + action + Key("alt:up"),
                                                 "shift": lambda action: Key("shift:down") + action + Key("shift:up"),
                                                 "control (alt|meta|under)": lambda action: Key("ctrl:down, alt:down") + action + Key("ctrl:up, alt:up"),
                                                 "control shift": lambda action: Key("ctrl:down, shift:down") + action + Key("ctrl:up, shift:up"),
                                                 "(alt|meta|under) shift": lambda action: Key("alt:down, shift:down") + action + Key("alt:up, shift:up"),
                                                 "control (alt|meta|under) shift": lambda action: Key("ctrl:down, alt:down, shift:down") + action + Key("ctrl:up, alt:up, shift:up"),
                                             })),
                                             full_key_element],
                                     value_func=lambda node, extras: (((extras["modifier"])(extras["single_key"]) + Pause("5")) * Repeat(extras["n"])))
        extras = [
            Repetition(RuleWrap(None, Alternative([command, repeated_command, combo_key_element])), min=1, max = 5, name="sequence"),
            Alternative([RuleRef(rule=character_rule)],
                        name="nested_repetitions"),
            Repetition(dictation_element, min=1, max=5, name="dictation_sequence"),
            utils.renamed_element("dictation", dictation_element),
            utils.renamed_element("terminal_command", terminal_command),
            IntegerRef("n", 1, 21),  # Times to repeat the sequence.
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
                              extras=extras, defaults=defaults, exported=True)

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
        if final_command:
            final_command.execute()
        global command_benchmark
        if command_benchmark.is_active():
            command_benchmark.record_and_replay_recognition()


#-------------------------------------------------------------------------------
# Define top-level rules for different contexts. Note that Dragon only allows
# top-level rules to be context-specific, but we want control over sub-rules. To
# work around this limitation, we compile a mutually exclusive top-level rule
# for each context.

class Environment(object):
    """Environment where voice commands can be spoken. Combines grammar and context
    and adds hierarchy. When installed, will produce a mutually-exclusive
    top-level grammar for each environment.

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

    def install(self, exported_rule_factory):
        grammars = []
        exclusive_context = self.context
        for child in self.children:
            grammars.extend(child.install(exported_rule_factory))
            exclusive_context = utils.combine_contexts(exclusive_context, ~child.context)
        rule_map = dict([(key, RuleRef(rule=utils.create_rule(self.name + "_" + key, action_map, element_map)) if action_map else Empty())
                         for (key, (action_map, element_map)) in self.environment_map.items()])
        grammar = Grammar(self.name, context=exclusive_context)
        grammar.add_rule(exported_rule_factory(self.name + "_exported", **rule_map))
        grammar.load()
        grammars.append(grammar)
        return grammars


class MyEnvironment(object):
    """Specialization of Environment for convenience with my exported rule factory
    (RepeatRule).
    """

    def __init__(self,
                 name,
                 parent=None,
                 context=None,
                 action_map=None,
                 repeatable_action_map=None,
                 terminal_action_map=None,
                 element_map=None):
        self.environment = Environment(
            name,
            {"command": (action_map or {}, element_map or {}),
             "repeatable_command": (repeatable_action_map or {}, element_map or {}),
             "terminal_command": (terminal_action_map or {}, element_map or {})},
            context,
            parent.environment if parent else None)

    def add_child(self, child):
        self.environment.add_child(child.environment)

    def install(self):
        def create_exported_rule(name, command, terminal_command, repeatable_command):
            return RepeatRule(name, command or Empty(), repeatable_command or Empty(), terminal_command or Empty())
        return self.environment.install(create_exported_rule)


### Global

global_environment = MyEnvironment(name="Global",
                                   action_map=command_action_map,
                                   repeatable_action_map=repeatable_action_map,
                                   element_map=command_element_map)


### Shell commands

shell_command_map = utils.combine_maps({
    "git commit": Text("git commit -am "),
    "git commit done": Text("git commit -am done "),
    "git checkout new": Text("git checkout -b "),
    "git reset hard head": Text("git reset --hard HEAD "),
    "fig XL": Text("hg xl "),
    "fig sync": Text("hg sync "),
    "fig checkout": Text("hg checkout "),
    "fig checkout P4 head": Text("hg checkout p4head "),
    "fig add": Text("hg add "),
    "fig commit": Text("hg commit -m "),
    "fig diff": Text("hg diff "),
    "fig P diff": Text("hg pdiff "),
    "fig amend": Text("hg amend "),
    "fig mail": Text("hg mail -m "),
    "fig upload": Text("hg uploadchain "),
    "fig submit": Text("hg submit "),
    "(soft|sym) link": Text("ln -s "),
    "list": Text("ls -l "),
    "make dear": Text("mkdir "),
    "ps (a UX|aux)": Text("ps aux "),
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

    def __init__(self, pre_action, post_action, tight=False, other_buffer=False):
        super(UseLinesAction, self).__init__()
        self.pre_action = pre_action
        self.post_action = post_action
        self.tight = tight
        self.other_buffer = other_buffer

    def _execute(self, data=None):
        if self.other_buffer:
            Key("c-x, o").execute()
        else:
            # Set mark without activating.
            Key("c-backslash").execute()
        MarkLinesAction(self.tight).execute(data)
        self.pre_action.execute(data)
        # Jump back to the beginning of the selection.
        Key("c-langle").execute()
        if self.other_buffer:
            Key("c-x, o").execute()
        else:
            # Jump back to the original position.
            Key("c-langle").execute()
        if not self.tight:
            Key("c-a").execute()
        self.post_action.execute(data)

emacs_repeatable_action_map = {
    # Overrides
    "afters": None,
    "befores": None,
    "aheads": None,
    "behinds": None,
    "rights": None,
    "lefts": None,

    # Movement
    "preev": Key("c-r"),
    "next": Key("c-s"),

    # Undo
    "cancel": Key("c-g"),
    "(shuck|undo)": Key("c-slash"),
    "redo": Key("c-question"),

    # Movement
    "layer forward": Key("ca-f"),
    "layer back": Key("ca-b"),
    "layer down": Key("ca-d"),
    "layer up": Key("ca-u"),
    "exper forward": Key("c-c, c, c-f"),
    "exper back": Key("c-c, c, c-b"),
    "word preev": Key("a-p"),
    "word next": Key("a-n"),
    "error preev": Key("f11"),
    "error next": Key("f12"),
}

emacs_action_map = {
    # Overrides
    "[<n>] up": Key("c-u") + Text("%(n)s") + Key("up"),
    "[<n>] down": Key("c-u") + Text("%(n)s") + Key("down"),
    "all select": Key("c-x, h"),
    "all edit": Key("c-x, h, c-w") + utils.RunApp("notepad") + Key("c-v"),
    "this edit": Key("c-w") + utils.RunApp("notepad") + Key("c-v"),

    # General
    "exec": Key("a-x"),
    "helm": Key("c-x, c"),
    "helm open": Key("c-x, c, b"),
    "prefix": Key("c-u"),
    "reload": Key("g"),
    "quit": Key("q"),
    "link open": Key("c-c, c, u/25") + OpenClipboardUrlAction(),

    # Emacs
    "help variable": Key("c-h, v"),
    "help function": Key("c-h, f"),
    "help key": Key("c-h, k"),
    "help mode": Key("c-h, m"),
    "help back": Key("c-c, c-b"),
    "customize open": Exec("customize-apropos"),

    # Window manipulation
    "buff open": Key("c-x, b"),
    "buff open split": Key("c-x, 3, c-x, o, c-x, b"),
    "buff switch": Key("c-x, b, enter"),
    "buff split": Key("c-x, 3"),
    "buff close": Key("c-x, 0"),
    "other close|buff close other": Key("c-x, 1"),
    "buff done": Key("c-x, hash"),
    "buff other": Key("c-x, o"),
    "buff kill": Key("c-x, k, enter"),
    "buff even": Key("c-x, plus"),
    "buff up": Exec("windmove-up"),
    "buff down": Exec("windmove-down"),
    "buff left": Exec("windmove-left"),
    "buff right": Exec("windmove-right"),

    # Filesystem
    "save": Key("c-x, c-s"),
    "save as": Key("c-x, c-w"),
    "save all": Key("c-x, s"),
    "save all now": Key("c-u, c-x, s"),
    "file open": Key("c-x, c-f"),
    "no ido": Key("c-f"),
    "directory open": Key("c-x, d"),
    "file open split": Key("c-x, 4, f"),
    "file open project": Key("c-c, p, f"),
    "file open simulator": Key("c-c, c, p, s"),
    "project open": Key("c-c, p, p"),
    "project switch": Key("c-c, s"),
    "result next": Key("a-comma"),
    "def open": Key("a-dot"),
    "ref open": Key("as-slash, enter"),
    "def close": Key("a-comma"),
    "R grep": Exec("rgrep"),
    "code search": Exec("cs-feeling-lucky"),
    "code search car": Exec("csc"),

    # Bookmarks
    "bookmark open": Key("c-x, r, b"),
    "bookmark save": Key("c-x, r, m"),
    "bookmark list": Key("c-x, r, l"),

    # Movement
    "start": Key("a-m"),
    "line <line> long": Key("a-g, a-g") + Text("%(line)s") + Key("enter"),
    "line <n1> [short]": jump_to_line("%(n1)s"),
    "here scroll": Key("c-l"),
    "mark set": Key("c-space"),
    "mark save": Key("c-backslash"),
    "go mark": Key("c-langle"),
    "go change": Key("c-c, c, c"),
    "go symbol": Key("a-i"),
    "go mark switch": Key("c-c, c-x"),
    "search edit": Key("a-e"),
    "search word": Key("a-s, w"),
    "search symbol": Key("a-s, underscore"),
    "regex preev": Key("ca-r"),
    "regex next": Key("ca-s"),
    "occur": Key("a-s, o"),
    "symbol preev": Key("a-s, dot, c-r, c-r"),
    "symbol next": Key("a-s, dot, c-s"),
    "go before [preev] <char>": Key("c-c, c, b") + Text("%(char)s"),
    "go after [next] <char>": Key("c-c, c, f") + Text("%(char)s"),
    "go before next <char>": Key("c-c, c, s") + Text("%(char)s"),
    "go after preev <char>": Key("c-c, c, e") + Text("%(char)s"),
    "other screen up": Key("c-minus, ca-v"),
    "other screen down": Key("ca-v"),
    "other <n1> enter": Key("c-x, o") + jump_to_line("%(n1)s") + Key("enter"),
    "go eye <char>": Key("c-c, c, j") + Text("%(char)s") + Function(lambda: eye_tracker.type_position("%d\n%d\n")),

    # Editing
    "delete": Key("c-c, c, c-w"),
    "[<n>] afters": Key("c-space, c-right/5:%(n)d"),
    "[<n>] befores": Key("c-space, c-left/5:%(n)d"),
    "[<n>] aheads": Key("c-space, a-f/5:%(n)d"),
    "[<n>] behinds": Key("c-space, a-b/5:%(n)d"),
    "[<n>] rights": Key("c-space, right/5:%(n)d"),
    "[<n>] lefts": Key("c-space, left/5:%(n)d"),
    "line open up": Key("a-enter"),
    "line open down": Key("c-enter"),
    "(this|line) move [<n>] up": Key("c-u") + Text("%(n)d") + Key("a-up"),
    "(this|line) move [<n>] down": Key("c-u") + Text("%(n)d") + Key("a-down"),
    "(this|line) copy [<n>] up": Key("c-u") + Text("%(n)d") + Key("as-up"),
    "(this|line) copy [<n>] down": Key("c-u") + Text("%(n)d") + Key("as-down"),
    "line clear": Key("c-a, c-c, c, k"),
    "(line|lines) join": Key("as-6"),
    "line <n1> open": jump_to_line("%(n1)s") + Key("a-enter"),
    "this select": Key("c-x, c-x"),
    "this indent": Key("ca-backslash"),
    "(this|here) comment": Key("a-semicolon"),
    "this format [clang]": Key("ca-q"),
    "this format comment": Key("a-q"),
    "replace": Key("as-5"),
    "regex replace": Key("cas-5"),
    "symbol replace": Key("a-apostrophe"),
    "cut": Key("c-w"),
    "copy": Key("a-w"),
    "paste": Key("c-y"),
    "paste other": Key("a-y"),
    "<n1> through [<n2>]": MarkLinesAction(),
    "<n1> through [<n2>] short": MarkLinesAction(True),
    "<n1> through [<n2>] copy here": UseLinesAction(Key("a-w"), Key("c-y")),
    "<n1> through [<n2>] short copy here": UseLinesAction(Key("a-w"), Key("c-y"), tight=True),
    "<n1> through [<n2>] move here": UseLinesAction(Key("c-w"), Key("c-y")),
    "<n1> through [<n2>] short move here": UseLinesAction(Key("c-w"), Key("c-y"), tight=True),
    "other <n1> through [<n2>] copy here": UseLinesAction(Key("a-w"), Key("c-y"), other_buffer=True),
    "other <n1> through [<n2>] short copy here": UseLinesAction(Key("a-w"), Key("c-y"), tight=True, other_buffer=True),
    "other <n1> through [<n2>] move here": UseLinesAction(Key("c-w"), Key("c-y"), other_buffer=True),
    "other <n1> through [<n2>] short move here": UseLinesAction(Key("c-w"), Key("c-y"), tight=True, other_buffer=True),
    "layer select": Key("cas-2"),
    "layer kill": Key("ca-k"),
    "select more": Key("c-equals"),
    "select less": Key("c-plus"),
    "this parens": Key("a-lparen"),
    "tag close": Key("c-c, c-e"),

    # Registers
    "(reg|rej) <char> here": Key("c-x, r, space, %(char)s"),
    "(reg|rej) <char> mark": Key("c-c, c, m, %(char)s"),
    "go (reg|rej) <char>": Key("c-x, r, j, %(char)s"),
    "(reg|rej) <char> this": Key("c-x, r, s, %(char)s"),
    "(reg|rej) <char> copy": Key("c-c, c, w, %(char)s"),
    "(reg|rej) <char> paste": Key("c-u, c-x, r, i, %(char)s"),

    # Templates
    "plate <template>": Key("c-c, ampersand, c-s") + Text("%(template)s") + Key("enter"),
    "(snippet|template) open": Key("c-c, ampersand, c-v"),
    "(snippet|template) new": Key("c-c, ampersand, c-n"),
    "(snippets|templates) reload": Exec("yas-reload-all"),

    # Compilation
    "file build": Key("c-c/10, c-g"),
    "file test": Key("c-c, c-t"),
    "recompile": Exec("recompile"),

    # Dired
    "toggle details": Exec("dired-hide-details-mode"),

    # Web editing
    "JavaScript mode": Exec("js-mode"),
    "HTML mode": Exec("html-mode"),

    # C++
    "header open": Key("c-x, c-h"),
    "header open split": Key("c-x, 3, c-x, o, c-x, c-h"),
    "import copy": Key("f5"),
    "import paste": Key("f6"),
    "this import": Exec("clang-include-fixer-at-point"),

    # Python
    "pie flakes": Key("c-c, c-v"),

    # Shell
    "shell open": Exec("shell"),
    "shell open directory": Key("c-c, c, dollar"),

    # Clojure
    "closure compile": Key("c-c, c-k"),
    "closure namespace": Key("c-c, a-n"),

    # Lisp
    "function eval": Key("ca-x"),
    "this eval": Exec("eval-region"),

    # Version control
    "magit open": Key("c-c, m"),
    "diff open": Key("c-x, v, equals"),
}

emacs_terminal_action_map = {
    "go before [preev] <custom_text>": Key("c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
    "go after preev <custom_text>": Key("left, c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("c-s, enter"),
    "go before next <custom_text>": Key("right, c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("c-r, enter"),
    "go after [next] <custom_text>": Key("c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
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
    "namespace": "namespace",
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
                                  repeatable_action_map=emacs_repeatable_action_map,
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
        "shell up": Key("a-p"),
        "shell down": Key("a-n"),
        "shell (preev|back)": Key("a-r"),
        "show output": Key("c-c, c-r"),
    })
emacs_shell_environment = MyEnvironment(name="EmacsShell",
                                        parent=emacs_environment,
                                        context=linux.UniversalAppContext(title="- Shell -"),
                                        action_map=emacs_shell_action_map)


### Shell

shell_repeatable_action_map = {
    "afters": None,
    "befores": None,
    "aheads": None,
    "behinds": None,
    "rights": None,
    "lefts": None,
    "afters delete": Key("a-d"),
    "befores delete": Key("a-backspace"),
    "aheads delete": Key("a-d"),
    "behinds delete": Key("a-backspace"),
    "rights delete": Key("del"),
    "lefts delete": Key("backspace"),
    "screen up": Key("s-pgup"),
    "screen down": Key("s-pgdown"),
    "tab left": Key("cs-left"),
    "tab right": Key("cs-right"),
    "preev": Key("c-r"),
    "next": Key("c-s"),
    "cancel": Key("c-g"),
    "tab close": Key("cs-w"),
}
shell_action_map = utils.combine_maps(
    shell_command_map,
    {
        "cut": Key("cs-x"),
        "copy": Key("cs-c"),
        "paste": Key("cs-v"),
        "tab move [<n>] left": Key("cs-pgup/5:%(n)d"),
        "tab move [<n>] right": Key("cs-pgdown/5:%(n)d"),
        "go tab <tab_n>": Key("a-%(tab_n)d"),
        "go tab last": Key("a-1, cs-left"),
        "tab new": Key("cs-t"),
    })

shell_element_map = {
    "tab_n": IntegerRef(None, 1, 10),
}

shell_environment = MyEnvironment(name="Shell",
                                  parent=global_environment,
                                  context=linux.UniversalAppContext(title=" - Terminal"),
                                  action_map=shell_action_map,
                                  repeatable_action_map=shell_repeatable_action_map,
                                  element_map=shell_element_map)


### Cmder

cmder_repeatable_action_map = {
    "tab left": Key("cs-tab"),
    "tab right": Key("c-tab"),
    "preev": Key("c-r"),
    "next": Key("c-s"),
    "cancel": Key("c-g"),
    "tab close": Key("c-w"),
}
cmder_action_map = utils.combine_maps(
    shell_command_map,
    {
        "tab (new|bash)": Key("as-5"),
        "tab dos": Key("as-2"),
    })

cmder_element_map = {
    "tab_n": IntegerRef(None, 1, 10),
}

cmder_environment = MyEnvironment(name="Cmder",
                                  parent=global_environment,
                                  context=AppContext(title="Cmder"),
                                  action_map=cmder_action_map,
                                  repeatable_action_map=cmder_repeatable_action_map,
                                  element_map=cmder_element_map)


### Chrome

chrome_repeatable_action_map = {
    "tab right":           Key("c-tab"),
    "tab left":           Key("cs-tab"),
    "tab close":          Key("c-w"),
}

chrome_action_map = {
    "go before <text>": Function(lambda text: a11y_utils.move_cursor(a11y_controller, str(text), before=True)),
    "go after <text>": Function(lambda text: a11y_utils.move_cursor(a11y_controller, str(text), before=False)),
    "words <text>": Function(select_words),
    "replace <text> with <replacement>": Function(replace_words),
    "link": Key("c-comma"),
    "link tab|tab [new] link": Key("c-dot"),
    "(link|links) background [tab]": Key("a-f"),
    "tab new":            Key("c-t"),
    "tab incognito":            Key("cs-n"),
    "window new": Key("c-n"),
    "go address":        Key("c-l"),
    "go [<n>] back":               Key("a-left/15:%(n)d"),
    "go [<n>] forward":            Key("a-right/15:%(n)d"),
    "reload": Key("c-r"),
    "go tab <tab_n>": Key("c-%(tab_n)d"),
    "go tab last": Key("c-9"),
    "go tab preev": Key("cs-1"),
    "tab move [<n>] left": Key("cs-pgup/5:%(n)d"),
    "tab move [<n>] right": Key("cs-pgdown/5:%(n)d"),
    "tab reopen":         Key("cs-t"),
    "tab dupe": Key("c-l/15, a-enter"),
    "find":               Key("c-f"),
    "<link> go":          Text("%(link)s"),
    "(caret|carrot) browsing": Key("f7"),
    "code search car": Key("c-l/15") + Text("csc") + Key("tab"),
    "code search simulator": Key("c-l/15") + Text("css") + Key("tab"),
    "code search": Key("c-l/15") + Text("cs") + Key("tab"),
    "calendar site": Key("c-l/15") + Text("calendar.google.com") + Key("enter"),
    "critique site": Key("c-l/15") + Text("cr/") + Key("enter"),
    "buganizer site": Key("c-l/15") + Text("b/") + Key("enter"),
    "drive site": Key("c-l/15") + Text("drive.google.com") + Key("enter"),
    "docs site": Key("c-l/15") + Text("docs.google.com") + Key("enter"),
    "slides site": Key("c-l/15") + Text("slides.google.com") + Key("enter"),
    "sheets site": Key("c-l/15") + Text("sheets.google.com") + Key("enter"),
    "new (docs|doc) site": Key("c-l/15") + Text("go/newdoc") + Key("enter"),
    "new (slides|slide) site": Key("c-l/15") + Text("go/newslide") + Key("enter"),
    "new (sheets|sheet) site": Key("c-l/15") + Text("go/newsheet") + Key("enter"),
    "new (scripts|script) site": Key("c-l/15") + Text("go/newscript") + Key("enter"),
    "amazon site": Key("c-l/15") + Text("smile.amazon.com") + Key("enter"),
    "this strikethrough": Key("as-5"),
    "this bullets": Key("cs-8"),
    "this bold": Key("c-b"),
    "this link": Key("c-k"),
    "insert text box": Key("a-i/15, t"),
    "paste raw": Key("cs-v"),
    "match next": Key("c-g"),
    "match preev": Key("cs-g"),
    "bookmark open": Key("c-semicolon"),
    "tab [new] bookmark": Key("c-apostrophe"),
    "bookmark save": Key("c-d"),
    "frame next": Key("c-lbracket"),
    "developer tools": Key("cs-j"),
    "webdriver test": Function(webdriver.test_driver),
    "go search": webdriver.ClickElementAction(By.NAME, "q"),
    "bill new": webdriver.ClickElementAction(By.LINK_TEXT, "Add a bill"),
}

chrome_terminal_action_map = {
    "search <text>":        Key("c-l/15") + Text("%(text)s") + Key("enter"),
    "history search <text>": Key("c-l/15") + Text("history") + Key("tab") + Text("%(text)s") + Key("enter"),
    "moma search <text>": Key("c-l/15") + Text("moma") + Key("tab") + Text("%(text)s") + Key("enter"),
    "<link>":          Text("%(link)s"),
}

link_chars_map = {
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
link_chars_dict_list  = DictList("link_chars_dict_list", link_chars_map)
chrome_element_map = {
    "tab_n": IntegerRef(None, 1, 9),
    "link": utils.JoinedRepetition(
        "", DictListRef(None, link_chars_dict_list), min=1, max=5),
    "replacement": Dictation(),
}

chrome_environment = MyEnvironment(name="Chrome",
                                   parent=global_environment,
                                   context=(AppContext(title=" - Google Chrome") | AppContext(executable="firefox.exe") | AppContext(title="Mozilla Firefox")),
                                   action_map=chrome_action_map,
                                   repeatable_action_map=chrome_repeatable_action_map,
                                   terminal_action_map=chrome_terminal_action_map,
                                   element_map=chrome_element_map)


### Chrome: Amazon

amazon_action_map = {
    "go search": webdriver.ClickElementAction(By.NAME, "field-keywords"),
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
    "today": Key("t"),
    "preev": Key("k"),
    "next": Key("j"),
    "day": Key("d"),
    "week": Key("w"),
    "month": Key("m"),
    "agenda": Key("a"),
}
calendar_environment = MyEnvironment(name="Calendar",
                                     parent=chrome_environment,
                                     context=(AppContext(title="Google Calendar") |
                                              AppContext(title="Google.com - Calendar")),
                                     action_map=calendar_action_map)


### Chrome: Code search

code_search_action_map = {
    "header open": Key("r/25, h"),
    "cc open": Key("r/25, c"),
    "directory open": Key("r/25, p"),
    "go search": Key("slash"),
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
    "snooze": Key("l/50") + Text("snooze") + Key("enter") + Text("["),
    "label candidates": Key("l/50") + Text("candidates") + Key("enter"),
    "check": Key("x"),
    "check next <n>": Key("x, j") * Repeat(extra="n"),
    "new messages": Key("N"),
    "go inbox|going box": Key("g, i"),
    "go starred": Key("g, s"),
    "go sent": Key("g, t"),
    "go drafts": Key("g, d"),
    "expand all": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Expand all']"),
    "collapse all": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Collapse all']"),
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
    "preev comment": Key("ctrl:down, alt:down, p, c, ctrl:up, alt:up"),
    "next comment": Key("ctrl:down, alt:down, n, c, ctrl:up, alt:up"),
    "enter comment": Key("ctrl:down, alt:down, e, c, ctrl:up, alt:up"),
    "(new|insert) row above": Key("a-i/15, r"),
    "(new|insert) row [below]": Key("a-i/15, b"),
    "dupe row": Key("s-space:2, c-c/15, a-i/15, b, c-v/30, up/30, down"),
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
# Populate and load the grammars.

grammars = global_environment.install()

# TODO Figure out either how to integrate this with the repeating rule or move out.
linux_grammar = Grammar("linux")   # Create this module's grammar.
linux_grammar.add_rule(linux_rule)
linux_grammar.load()
grammars.append(linux_grammar)

class BenchmarkRule(MappingRule):
    mapping = {
        "benchmark [<n>] command <command>": Function(lambda command, n: command_benchmark.start(str(command), n)),
        "benchmark reset": Function(reset_benchmark),
    }
    extras = [Dictation("command"), IntegerRef("n", 1, 10, default=1)]

benchmark_grammar = Grammar("benchmark")
benchmark_grammar.add_rule(BenchmarkRule())
benchmark_grammar.load()
grammars.append(benchmark_grammar)


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
        callback = callbacks.get_nowait()
        try:
            callback()
        except Exception as exception:
            traceback.print_exc()


timer = get_engine().create_timer(RunCallbacks, 0.1)


# Update the context words and phrases.
def UpdateWords(words, phrases):
    context_word_list.set(words)
    context_phrase_list.set(phrases)

def IsValidIp(ip):
    m = re.match(r"^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$", ip)
    return bool(m) and all(map(lambda n: 0 <= int(n) <= 255, m.groups()))

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
        # Check host in case of DNS rebinding attack.
        host = self.headers.getheader("Host")
        host = host.split(":")[0]
        if not (host == "localhost" or host == "localhost." or IsValidIp(host)):
            print "Host header rejected: " + host
            return
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

# Connect to Chrome WebDriver if possible.
webdriver.create_driver()

# Connect to eye tracker if possible.
eye_tracker.connect()

print("Loaded _repeat.py")


#-------------------------------------------------------------------------------
# Unload function which will be called by NatLink.
def unload():
    global grammars, server, server_thread, timer
    for grammar in grammars:
        grammar.unload()
    eye_tracker.disconnect()
    webdriver.quit_driver()
    timer.stop()
    server.shutdown()
    server_thread.join()
    server.server_close()
    print("Unloaded _repeat.py")
