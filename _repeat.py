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

from collections import OrderedDict
import os.path
import re
import socket
import sys
import threading
import time
import webbrowser
import win32clipboard
import yappi

from odictliteral import odict
from six.moves import BaseHTTPServer
from six.moves import queue
from six import string_types
from concurrent import futures

from dragonfly import (
    ActionBase,
    Alternative,
    AppContext,
    Choice,
    Compound,
    CompoundRule,
    Config,
    CursorPosition,
    DictList,
    DictListRef,
    Dictation,
    Empty,
    FocusWindow,
    Function,
    Grammar,
    IntegerRef,
    Key,
    List,
    ListRef,
    Literal,
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
    TextQuery,
    get_accessibility_controller,
    get_engine,
)
import dragonfly.log
from selenium.webdriver.common.by import By

import _dragonfly_local as local
import _dragonfly_utils as utils
import _eye_tracker_utils as eye_tracker
import _linux_utils as linux
import _ocr_utils as ocr
import _text_utils as text
import _webdriver_utils as webdriver

# Instantiate the tracker so we can refer to it (we will connect to it later).
tracker = eye_tracker.get_tracker()

# Start a single-threaded threadpool for running OCR.
ocr_executor = futures.ThreadPoolExecutor(max_workers=1)
ocr_future = None

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
    "minus twice": "--",
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


dictation_key_action_map = {
    "enter|slap": Key("enter"),
    "space|spooce|spacebar": Key("space"),
    "tab-key": Key("tab"),
}

dictation_action_map = utils.combine_maps(dictation_key_action_map,
                                          utils.text_map_to_action_map(utils.combine_maps(letters_map, symbols_map)))

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
        "home key": Key("home"),
        "end key": Key("end"),
    })

full_key_action_map = utils.combine_maps(
    standalone_key_action_map,
    utils.text_map_to_key_action_map(utils.combine_maps(letters_map, numbers_map, symbol_keys_map)),
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


accessibility = get_accessibility_controller()

accessibility_commands = {}
# = odict[
#     "go before <text_position_query>": Function(lambda text_position_query: accessibility.move_cursor(
#         text_position_query, CursorPosition.BEFORE)),
#     "go after <text_position_query>": Function(lambda text_position_query: accessibility.move_cursor(
#         text_position_query, CursorPosition.AFTER)),
#     # Note that the delete command is declared first so that it has higher
#     # priority than the selection variant.
#     "words <text_query> delete": Function(lambda text_query: accessibility.replace_text(text_query, "")),
#     "words <text_query>": Function(accessibility.select_text),
#     "replace <text_query> with <replacement>": Function(accessibility.replace_text),
# ]


#-------------------------------------------------------------------------------
# Action maps to be used in rules.


def start_cpu_profiling():
    yappi.set_clock_type("cpu")
    yappi.start()


def start_wall_profiling():
    yappi.set_clock_type("wall")
    yappi.start()


def stop_profiling():
    yappi.stop()
    yappi.get_func_stats().print_all()
    yappi.get_thread_stats().print_all()
    has_argv = hasattr(sys, "argv")
    if not has_argv:
        sys.argv = [""]
    profile_path = os.path.join(local.HOME, "yappi_{}.callgrind.out".format(time.time()))
    yappi.get_func_stats().save(profile_path, "callgrind")
    yappi.clear_stats()


def move_to_text(text, cursor_position=ocr.CursorPosition.MIDDLE):
    word = str(text)
    (nearby_words, image), gaze_point = ocr_future.result()
    click_position = ocr.find_nearest_word_position(word, gaze_point, nearby_words, cursor_position)
    if local.SAVE_OCR_DATA_DIR:
        file_name_prefix = "{}_{:.2f}".format("success" if click_position else "failure", time.time())
        file_path_prefix = os.path.join(local.SAVE_OCR_DATA_DIR, file_name_prefix)
        image.save(file_path_prefix + ".png")
        with open(file_path_prefix + ".txt", "w") as file:
            file.write(word)
    if not click_position:
        # Raise an exception so that the action returns False.
        raise RuntimeError("No matches found for word: {}".format(word))
    Mouse("[{}, {}]".format(int(click_position[0]), int(click_position[1]))).execute()


def select_text(text, text2=None):
    move_to_text(text, ocr.CursorPosition.BEFORE)
    Mouse("left:down").execute()
    if not text2:
        text2 = text
    move_to_text(text2, ocr.CursorPosition.AFTER)
    Pause("5").execute()
    Mouse("left:up").execute()


def replace_text(text, replacement):
    select_text(text)
    Text(replacement.replace("%", "%%")).execute()


# Actions of commonly used text navigation and mousing commands. These can be
# used anywhere except after commands which include arbitrary dictation.
# TODO: Better solution for holding shift during a single command. Think about whether this could enable a simpler grammar for other modifiers.
command_action_map = utils.combine_maps(
    # These work like the built-in commands and are available for any
    # application that supports IAccessible2. A "my" prefix is used to avoid
    # overwriting similarly-phrased commands built into Dragon, because in some
    # applications only those will work. These are primarily present to test
    # functionality; to add these commands to a specific application, just merge
    # in the map without a prefix.
    OrderedDict([("my " + k, v) for k, v in accessibility_commands.items()]),
    {
        "delete": Key("del"),
        "go home|[go] west": Key("home"),
        "go end|[go] east": Key("end"),
        "go top|[go] north": Key("c-home"),
        "go bottom|[go] south": Key("c-end"),
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

        "(I|eye) connect": Function(tracker.connect),
        "(I|eye) disconnect": Function(tracker.disconnect),
        "(I|eye) print position": Function(tracker.print_position),
        "(I|eye) move": Function(tracker.move_to_position),
        "(I|eye) (touch|click) [left]": Function(tracker.move_to_position) + Mouse("left"),
        "(I|eye) (touch|click) right": Function(tracker.move_to_position) + Mouse("right"),
        "(I|eye) (touch|click) middle": Function(tracker.move_to_position) + Mouse("middle"),
        "(I|eye) (touch|click) [left] twice": Function(tracker.move_to_position) + Mouse("left:2"),
        "(I|eye) (touch|click) hold": Function(tracker.move_to_position) + Mouse("left:down"),
        "(I|eye) (touch|click) release": Function(tracker.move_to_position) + Mouse("left:up"),
        "scroll up": Function(lambda: tracker.move_to_position((0, 40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheelup:8"),
        "scroll up half": Function(lambda: tracker.move_to_position((0, 40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheelup:4"),
        "scroll down": Function(lambda: tracker.move_to_position((0, -40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheeldown:8"),
        "scroll down half": Function(lambda: tracker.move_to_position((0, -40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheeldown:4"),
        "scroll left": Function(lambda: tracker.move_to_position((0, 40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheelleft:8"),
        "scroll right": Function(lambda: tracker.move_to_position((0, 40)) or Mouse("(0.5, 0.5)").execute()) + Mouse("wheelright:8"),
        "here (touch|click) [left]": Mouse("left"),
        "here (touch|click) right": Mouse("right"),
        "here (touch|click) middle": Mouse("middle"),
        "here (touch|click) [left] twice": Mouse("left:2"),
        "here (touch|click) hold": Mouse("left:down"),
        "here (touch|click) release": Mouse("left:up"),
        "(touch|click) <text>": Function(move_to_text) + Mouse("left"),
        "(touch|click) right <text>": Function(move_to_text) + Mouse("right"),
        "(touch|click) middle <text>": Function(move_to_text) + Mouse("middle"),
        "(touch|click) [left] twice <text>": Function(move_to_text) + Mouse("left:2"),
        "(touch|click) hold <text>": Function(move_to_text) + Mouse("left:down"),
        "(touch|click) release <text>": Function(move_to_text) + Mouse("left:up"),
        "control (touch|click) <text>": Function(move_to_text) + Key("ctrl:down") + Mouse("left") + Key("ctrl:up"),
        "go before <text>": Function(lambda text: move_to_text(text, ocr.CursorPosition.BEFORE)) + Mouse("left"),
        "go after <text>": Function(lambda text: move_to_text(text, ocr.CursorPosition.AFTER)) + Mouse("left"),
        # Note that the delete command is declared first so that it has higher
        # priority than the selection variant.
        "words <text> [through <text2>] delete": Function(select_text) + Key("backspace"),
        "words <text> [through <text2>]": Function(select_text),
        "replace <text> with <replacement>": Function(replace_text),

        "webdriver open": Function(webdriver.create_driver),
        "webdriver close": Function(webdriver.quit_driver),

        "(hey|OK) google <text>": Function(lambda text: None),

        "dragonfly CPU profiling start": Function(start_cpu_profiling),
        "dragonfly wall [time] profiling start": Function(start_wall_profiling),
        "dragonfly [(CPU|wall [time])] profiling stop": Function(stop_profiling),
    })

# Actions for speaking out sequences of characters.
character_action_map = {
    "<chars> short": Text(u"%(chars)s"),
    "number <numerals>": Text(u"%(numerals)s"),
    "letter <letters>": Text(u"%(letters)s"),
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

# Same as above, except no arbitrary dictation allowed and context phrases are
# included.
custom_dictation = RuleWrap(None, utils.JoinedSequence(" ", [
    Optional(ListRef(None, prefix_list)),
    Alternative([
        letter_element,
        ListRef(None, context_phrase_list),
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
    "text2": Dictation(),
    "char": char_element,
    "custom_text": RuleWrap(None, Alternative([
        Dictation(),
        chars_element,
        ListRef(None, prefix_list),
        ListRef(None, suffix_list),
        ListRef(None, saved_word_list),
    ])),
    # TODO Figure out why we can't reuse custom_text element.
    "custom_text2": RuleWrap(None, Alternative([
        Dictation(),
        chars_element,
        ListRef(None, prefix_list),
        ListRef(None, suffix_list),
        ListRef(None, saved_word_list),
    ])),
    "replacement": Dictation(),
    "text_query": Compound(
        spec=("[[([<start_phrase>] <start_relative_position> <start_relative_phrase>|<start_phrase>)] <through>] "
              "([<end_phrase>] <end_relative_position> <end_relative_phrase>|<end_phrase>)"),
        extras=[Dictation("start_phrase", default=""),
                Alternative([Literal("before"), Literal("after")], name="start_relative_position"),
                Dictation("start_relative_phrase", default=""),
                Literal("through", "through", value=True, default=False),
                Dictation("end_phrase", default=""),
                Alternative([Literal("before"), Literal("after")], name="end_relative_position"),
                Dictation("end_relative_phrase", default="")],
        value_func=lambda node, extras: TextQuery(
            start_phrase=str(extras["start_phrase"]),
            start_relative_position=CursorPosition[extras["start_relative_position"].upper()] if "start_relative_position" in extras else None,
            start_relative_phrase=str(extras["start_relative_phrase"]),
            through=extras["through"],
            end_phrase=str(extras["end_phrase"]),
            end_relative_position=CursorPosition[extras["end_relative_position"].upper()] if "end_relative_position" in extras else None,
            end_relative_phrase=str(extras["end_relative_phrase"]))),
    "text_position_query": Compound(
        spec="<phrase> [<relative_position> <relative_phrase>]",
        extras=[Dictation("phrase", default=""),
                Alternative([Literal("before"), Literal("after")], name="relative_position"),
                Dictation("relative_phrase", default="")],
        value_func=lambda node, extras: TextQuery(
            end_phrase=str(extras["phrase"]),
            end_relative_position=CursorPosition[extras["relative_position"].upper()] if "relative_position" in extras else None,
            end_relative_phrase=str(extras["relative_phrase"]))),
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
        "padded <symbol>": Text(u" %(symbol)s "),
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
        "(mim|mimic) text <text>": Text(u"%(text)s"),
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
    {
        "number <numeral>": Text(u"%(numeral)s"),
        "upper <letter>": Function(lambda letter: Text(letter.upper()).execute()),
    },
    {
        "numeral": number_element,
        "letter": letter_element,
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
    RuleRef(rule=single_character_rule),
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
    if isinstance(window, string_types):
        window = [window]
    for j, words in enumerate(window):
        windows_mapping["(" + words + ") " + windows_suffix] = Key("win:down, %d:%d/20, win:up" % (i + 1, j + 1))

final_action_map = utils.combine_maps(windows_mapping, {
    "[work] terminal win": FocusWindow(executable="nxclient.bin", title=" - Terminal"),
    "[work] emacs win": FocusWindow(executable="nxclient.bin", title=" - Emacs editor"),
    "[work] studio win": FocusWindow(executable="nxclient.bin", title=" - Android Studio"),
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

    def _process_begin(self):
        # Start OCR now so that results are ready when the command completes (if
        # it uses OCR). This also has the benefit of using the gaze from the
        # time the user starts speaking.
        global ocr_future
        gaze_point = tracker.get_gaze_point_or_default()
        # Don't enqueue multiple requests.
        if ocr_future and not ocr_future.done():
            canceled = ocr_future.cancel()
            if canceled:
                print("Canceled OCR future.")
            else:
                print("Unable to cancel OCR future.")
        ocr_future = ocr_executor.submit(lambda: (ocr.find_nearby_words(gaze_point), gaze_point))

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
    "fig commit": Text("hg commit -e "),
    "fig diff": Text("hg diff "),
    "fig diff stat": Text("hg diff --stat "),
    "fig diff no binary": Text("hg diff --no-binary "),
    "fig P diff": Text("hg pdiff "),
    "fig P diff stat": Text("hg pdiff --stat "),
    "fig P diff no binary": Text("hg pdiff --no-binary "),
    "fig amend": Text("hg amend "),
    "fig mail": Text("hg mail -m "),
    "fig upload": Text("hg uploadchain "),
    "fig submit": Text("hg submit "),
    "fig status": Text("hg status "),
    "fig fix": Text("hg fix "),
    "fig lint": Text("hg lint"),
    "fig evolve": Text("hg evolve "),
    "fig preev": Text("hg prev "),
    "fig next": Text("hg next "),
    "fig shelve": Text("hg shelve "),
    "fig unshelve": Text("hg unshelve "),
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
        print("Opening link: %s" % data)
        webbrowser.open(data)


class MarkLinesAction(ActionBase):
    """Mark several lines within a range."""

    def __init__(self, tight=False, tree=False):
        super(MarkLinesAction, self).__init__()
        self.tight = tight
        self.tree = tree

    def _execute(self, data=None):
        jump_to_line("%(n1)d" % data).execute()
        if self.tree:
            Key("a-h").execute()
            return
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

    def __init__(self, pre_action, post_action, tight=False, other_buffer=False, tree=False):
        super(UseLinesAction, self).__init__()
        self.pre_action = pre_action
        self.post_action = post_action
        self.tight = tight
        self.other_buffer = other_buffer
        self.tree = tree

    def _execute(self, data=None):
        if self.other_buffer:
            Key("c-x, o").execute()
        else:
            # Set mark without activating.
            Key("c-backslash").execute()
        MarkLinesAction(tight=self.tight, tree=self.tree).execute(data)
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
    "undo": Key("c-slash"),
    "redo": Key("c-question"),

    # Movement
    "layer preev": Key("ca-b"),
    "layer next": Key("ca-f"),
    "layer down": Key("ca-d"),
    "layer up": Key("ca-u"),
    "exper preev": Key("c-c, c, c-b"),
    "exper next": Key("c-c, c, c-f"),
    "word preev": Key("a-p"),
    "word next": Key("a-n"),
    "error preev": Key("f11"),
    "error next": Key("f12"),
}

emacs_action_map = odict[    
    # Overrides
    "go before <text>": None,
    "go after <text>": None,
    "words <text> [through <text2>] delete": None,
    "words <text> [through <text2>]": None,
    "replace <text> with <replacement>": None,
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
    "buff done": Key("c-x, hash"),
    "buff kill": Key("c-x, k, enter"),
    "buff even": Key("c-x, plus"),
    "go other": Key("c-x, o"),
    "other close|buff focus": Key("c-x, 1"),
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
    "ido close": Key("c-f"),
    "ido reload": Key("c-l"),
    "directory open": Key("c-x, d"),
    "file open recent": Key("c-x, c-r"),
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
    "go before [preev] <char>": Key("c-c, c, b") + Text(u"%(char)s"),
    "go after [next] <char>": Key("c-c, c, f") + Text(u"%(char)s"),
    "go before next <char>": Key("c-c, c, s") + Text(u"%(char)s"),
    "go after preev <char>": Key("c-c, c, e") + Text(u"%(char)s"),
    "other screen up": Key("c-minus, ca-v"),
    "other screen down": Key("ca-v"),
    "other <n1> enter": Key("c-x, o") + jump_to_line("%(n1)s") + Key("enter"),
    "go eye <char>": Key("c-c, c, j") + Text(u"%(char)s") + Function(lambda: tracker.type_position("%d\n%d\n")),

    # Editing
    "delete": Key("c-c, c, c-w"),
    # Note that the delete commands are declared first so that they have higher
    # priority than the selection variants.
    "[<n>] afters delete": Key("c-del/5:%(n)d"),
    "[<n>] befores delete": Key("c-backspace/5:%(n)d"),
    "[<n>] aheads delete": Key("a-d/5:%(n)d"),
    "[<n>] behinds delete": Key("a-backspace/5:%(n)d"),
    "[<n>] rights delete": Key("del/5:%(n)d"),
    "[<n>] lefts delete": Key("backspace/5:%(n)d"),
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
    "<n1> through [<n2>] short": MarkLinesAction(tight=True),
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
    "mark save (reg|rej) <char>": Key("c-x, r, space, %(char)s"),
    "go (reg|rej) <char>": Key("c-x, r, j, %(char)s"),
    "copy (reg|rej) <char>": Key("c-x, r, s, %(char)s"),
    "(reg|rej) <char> paste": Key("c-u, c-x, r, i, %(char)s"),

    # Templates
    "plate <template>": Key("c-c, ampersand, c-s") + Text(u"%(template)s") + Key("enter"),
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
    "function run": Key("ca-x"),
    "this run": Exec("eval-region"),

    # Version control
    "magit open": Key("c-c, m"),
    "diff open": Key("c-x, v, equals"),
    "VC open": Key("c-x, v, d, enter"),

    # GhostText
    "ghost close": Key("c-c, c-c"),
]

emacs_terminal_action_map = {
    "go before [preev] <custom_text>": Key("c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
    "go after preev <custom_text>": Key("left, c-r") + utils.lowercase_text_action("%(custom_text)s") + Key("c-s, enter"),
    "go before next <custom_text>": Key("right, c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("c-r, enter"),
    "go after [next] <custom_text>": Key("c-s") + utils.lowercase_text_action("%(custom_text)s") + Key("enter"),
    "words <custom_text>": (Key("c-c, c, c-r")
                            + utils.lowercase_text_action("%(custom_text)s") + Key("enter")),
    "words <custom_text> through <custom_text2>": (Key("c-c, c, c-t")
                                                   + utils.lowercase_text_action("%(custom_text)s") + Key("enter")
                                                   + utils.lowercase_text_action("%(custom_text2)s") + Key("enter")),
    "replace <custom_text> with <custom_text2>": (Key("c-c, c, as-5")
                                                 + utils.lowercase_text_action("%(custom_text)s") + Key("enter")
                                                 + utils.lowercase_text_action("%(custom_text2)s") + Key("enter")),
}

templates = {
    "beginend": "beginend",
    "car": "car",
    "catch": "catch",
    "doc": "doc",
    "field declaration": "field_declaration",
    "field definition": "field_definition",
    "field initialize": "field_initialize",
    "finally": "finally",
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
    "lambda": "lambda",
    "list": "list",
    "map": "map",
    "method": "method",
    "namespace": "namespace",
    "new": "new",
    "override": "override",
    "ref": "ref",
    "set": "set",
    "shared pointer": "shared_pointer",
    "test": "test",
    "ternary": "ternary",
    "text": "text",
    "to do": "todo",
    "try": "try",
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

emacs_org_repeatable_action_map = {
    "heading preev": Key("c-c, c-b"),
    "heading next": Key("c-c, c-f"),
    "heading up": Key("c-c, c-u"),
}

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
    "tree select": Key("a-h"),
    "<n1> tree": MarkLinesAction(tree=True),
    "<n1> tree copy here": UseLinesAction(Key("a-w"), Key("c-y"), tree=True),
    "<n1> tree move here": UseLinesAction(Key("c-w"), Key("c-y"), tree=True),
    "other <n1> tree copy here": UseLinesAction(Key("a-w"), Key("c-y"), other_buffer=True, tree=True),
    "other <n1> tree move here": UseLinesAction(Key("c-w"), Key("c-y"), other_buffer=True, tree=True),
    "open org link": Key("c-c, c-o"),
    "show to do's": Key("c-c, slash, t"),
    "archive": Key("c-c, c-x, c-a"),
    "org (West|start)": Key("c-c, c, c-a"),
    "tag <tag>": Key("c-c, c-q") + Text(u"%(tag)s") + Key("enter"),
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
                                      repeatable_action_map=emacs_org_repeatable_action_map,
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
    "screen up": Key("c-pgup"),
    "screen down": Key("c-pgdown"),
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
        "tab new [cygwin]": Key("as-6"),
        "tab new ubuntu": Key("as-5"),
        "tab new dos": Key("as-2"),
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
    "workspace open": Mouse("[0.8, 0.8]") + Key("c-1/15, a-s"),
    "workspace tab new": Key("as-f"),
    "workspace close": Key("a-w"),
    "workspace new": Key("a-n"),
    "workspace [tab] save": Key("a-d"),
    "find":               Key("c-f"),
    "<link> go":          Text("%(link)s"),
    "(caret|carrot) browsing": Key("f7"),
    "code search (voice access|VA)": Key("c-l/15") + Text("csva") + Key("tab"),
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
    "this numbers": Key("cs-7"),
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
    # "<text> touch": webdriver.SmartClickElementAction(By.XPATH,
    #                                                   ("//*[text()[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), " +
    #                                                    "translate('%(text)s', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'))]]"),
    #                                                   tracker),
    "ghost open": Key("ca-k"),
}

chrome_terminal_action_map = utils.combine_maps(
    accessibility_commands,
    {
        "search <text>":        Key("c-l/15") + Text(u"%(text)s") + Key("enter"),
        "history search <text>": Key("c-l/15") + Text("history") + Key("tab") + Text(u"%(text)s") + Key("enter"),
        "history search": Key("c-l/15") + Text("history") + Key("tab"),
        "moma search <text>": Key("c-l/15") + Text("moma") + Key("tab") + Text(u"%(text)s") + Key("enter"),
        "moma search": Key("c-l/15") + Text("moma") + Key("tab"),
        "<link>":          Text("%(link)s"),
    })

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
    "(touch|click) LGTM": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='LGTM']"),
    "(touch|click) action required": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Action required']"),
    "(touch|click) send": webdriver.ClickElementAction(By.XPATH, "//*[starts-with(@aria-label, 'Send')]"),
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

gmail_repeatable_action_map = {
    "preev": Key("plus, k"),
    "next": Key("plus, j"),
    "message preev": Key("plus, p"),
    "message next": Key("plus, n"),
    "section next": Key("plus, backtick"),
    "section preev": Key("plus, tilde"),
}

gmail_action_map = {
    "open": Key("plus, o"),
    "archive": Key("+, {"),
    "done": Key("+, ["),
    "this unread": Key("+, _"),
    "undo": Key("plus, z"),
    "list": Key("plus, u"),
    "compose": Key("plus, c"),
    "reply": Key("plus, r"),
    "reply all": Key("plus, a"),
    "forward": Key("plus, f"),
    "important": Key("plus, plus"),
    "this star": Key("plus, s"),
    "this important": Key("plus, plus"),
    "this not important": Key("plus, minus"),
    "label waiting": Pause("50") + Key("plus, l/50") + Text("waiting") + Pause("50") + Key("enter"),
    "label snooze": Pause("50") + Key("plus, l/50") + Text("snooze") + Pause("50") + Key("enter"),
    "snooze": Pause("50") + Key("plus, l/50") + Text("snooze") + Pause("50") + Key("enter") + Pause("50") + Key("+, ["),
    "label vacation": Pause("50") + Key("plus, l/50") + Text("vacation") + Pause("50") + Key("enter"),
    "label house": Pause("50") + Key("plus, l/50") + Text("house") + Pause("50") + Key("enter"),
    "this select": Key("plus, x"),
    "<n> select": Key("plus, x, plus, j") * Repeat(extra="n"),
    "(message|messages) reload": Key("plus, N"),
    "go inbox|going box": Key("plus, g, i"),
    "go starred": Key("plus, g, s"),
    "go sent": Key("plus, g, t"),
    "go drafts": Key("plus, g, d"),
    "expand all": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Expand all']"),
    "collapse all": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Collapse all']"),
    "go field to": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='To']"),
    "go field cc": Key("cs-c"),
    "chat open": Key("plus, q"),
    "this send": Key("c-enter"),
    "go search": Key("plus, slash"),
}

gmail_environment = MyEnvironment(name="Gmail",
                                  parent=chrome_environment,
                                  context=(AppContext(title="Gmail") |
                                           AppContext(title="Google.com Mail") |
                                           AppContext(title="<mail.google.com>") |
                                           AppContext(title="<inbox.google.com>")),
                                  action_map=gmail_action_map,
                                  repeatable_action_map=gmail_repeatable_action_map)


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
    "(click|touch) present": webdriver.ClickElementAction(By.XPATH, "//*[@aria-label='Start presentation (Ctrl+F5)']"),
    "file rename": Key("as-f/50, r"),
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


### Chrome: Colab

colab_repeatable_action_map = {
    "[cell] next": Key("c-m, n"),
    "[cell] preev": Key("c-m, p"),
}
colab_action_map = {
    "save": Key("c-s"),
    "cell run": Key("c-enter"),
    "all (cell|cells) run": Key("c-f9"),
    "cell (expand|collapse)": Key("c-apostrophe"),
    "cell open down": Key("c-m, b"),
    "cell open up": Key("c-m, a"),
    "cell delete": Key("c-m, d"),
    "this run": Key("cs-enter"),
    "this comment": Key("c-slash"),
}
colab_environment = MyEnvironment(name="Colab",
                                  parent=chrome_environment,
                                  context=(AppContext(title="<colab.sandbox.google.com>") |
                                           AppContext(title="<colab.research.google.com>")),
                                  action_map=colab_action_map,
                                  repeatable_action_map=colab_repeatable_action_map)


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
        "terminal new": Key("ca-t"),
        # "[work] terminal win": linux.ActivateLinuxWindow(" - Terminal"),
        # "[work] emacs win": linux.ActivateLinuxWindow(" - Emacs editor"),
        # "[work] studio win": linux.ActivateLinuxWindow(" - Android Studio"),
        "remote firefox win": linux.ActivateLinuxWindow("Mozilla Firefox"),
        "remote chrome win": linux.ActivateLinuxWindow("Google Chrome"),
    })
run_local_hook("AddLinuxCommands", linux_action_map)
linux_rule = utils.create_rule("LinuxRule", linux_action_map, {}, True,
                               (AppContext(title="Oracle VM VirtualBox") |
                                AppContext(title="<remotedesktop.corp.google.com>")))


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
callbacks = queue.Queue()


def RunCallbacks():
    while callbacks and not callbacks.empty():
        callback = callbacks.get_nowait()
        try:
            callback()
        except Exception as exception:
            traceback.print_exc()


timer = get_engine().create_timer(RunCallbacks, 0.1)

# Update the context phrases.
def UpdateWords(phrases):
    # Not currently used, and not working if enabled. Tests indicate that the
    # list of words is being updated, but Dragon is only aware of the words that
    # were initially loaded.
    # context_phrase_list.set(phrases)
    pass

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
            print("Host header rejected: " + host)
            return
        length = self.headers.getheader("content-length")
        file_type = self.headers.getheader("My-File-Type")
        request_text = self.rfile.read(int(length)) if length else ""
        # print("received text: %s" % request_text)
        phrases = text.extract_phrases(request_text, file_type)
        # Asynchronously update word lists available to Dragon.
        callbacks.put_nowait(lambda: UpdateWords(phrases))
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
tracker_thread = threading.Thread(target=tracker.connect)
tracker_thread.start()

# Force NatLink to schedule background threads frequently by regularly waking up
# a dummy thread.
shutdown_dummy_thread_event = threading.Event()
def run_dummy_thread():
    while not shutdown_dummy_thread_event.is_set():
        time.sleep(1)

dummy_thread = threading.Thread(target=run_dummy_thread)
dummy_thread.start()

# Initialise a dragonfly timer to manually yield control to the thread.
def wake_dummy_thread():
    dummy_thread.join(0.002)

wake_dummy_thread_timer = get_engine().create_timer(wake_dummy_thread, 0.02)

print("Loaded _repeat.py")


#-------------------------------------------------------------------------------
# Unload function which will be called by NatLink.
def unload():
    for grammar in grammars:
        grammar.unload()
    if tracker.is_available:
        tracker.disconnect()
    webdriver.quit_driver()
    timer.stop()
    server.shutdown()
    server_thread.join()
    server.server_close()
    wake_dummy_thread_timer.stop()
    shutdown_dummy_thread_event.set()
    dummy_thread.join()
    ocr_executor.shutdown(wait=False)
    print("Unloaded _repeat.py")
