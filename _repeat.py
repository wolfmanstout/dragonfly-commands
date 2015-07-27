#
# This file is a command-module for Dragonfly.
# (c) Copyright 2008 by Christo Butcher
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>
#

"""
This contains all commands which may be spoken continuously or repeated.

This is heavily modified from _multiedit.py, found here:
https://code.google.com/p/dragonfly-modules/

"""

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r99")
except ImportError:
    pass

import BaseHTTPServer
import Queue
import socket
import threading
import time
import urllib

from dragonfly import *
import dragonfly.log
from selenium.webdriver.common.by import By

from _dragonfly_utils import *
from _text_utils import *
from _eye_tracker_utils import *
from _webdriver_utils import *

# Load local hooks if defined.
try:
    import _dragonfly_local_hooks as local_hooks
    def RunLocalHook(name, *args, **kwargs):
        """Function to run local hook if defined."""
        hook = getattr(local_hooks, name)
        if hook:
            return hook(*args, **kwargs)
except:
    print("Local hooks not loaded.")
    def RunLocalHook(name, *args, **kwargs):
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

letters_map = combine_maps(quick_letters_map, long_letters_map)

char_map = dict((k, v.strip()) for (k, v) in combine_maps(letters_map, numbers_map, symbol_map).iteritems())

# Load commonly misrecognized words saved to a file.
saved_words = []
try:
    with open(WORDS_PATH) as file:
        for line in file:
            word = line.strip()
            if len(word) > 2 and word not in letters_map:
                saved_words.append(line.strip())
except:
    print("Unable to open: " + WORDS_PATH)

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

    "paste":                            release + Key("c-v"),
    "copy":                             release + Key("c-c"),
    "cut":                              release + Key("c-x"),
    "select everything":                       release + Key("c-a"),
    "edit text": RunApp("notepad"),
    "edit everything": Key("c-a, c-x") + RunApp("notepad") + Key("c-v"),
    "edit region": Key("c-x") + RunApp("notepad") + Key("c-v"),
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
    "do click": Mouse("left"),
    "do right click": Mouse("right"),
    "do middle click": Mouse("middle"),
    "do double click": Mouse("left:2"),
    "do triple click": Mouse("left:3"),
    "do start drag": Mouse("left:down"),
    "do stop drag": Mouse("left:up"),
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
char_dict_list = DictList("char_dict_list", char_map)
saved_word_list = List("saved_word_list", saved_words)
# Lists which will be populated later via RPC.
context_phrase_list = List("context_phrase_list", [])
context_word_list = List("context_word_list", [])
prefix_list = List("prefix_list", prefixes)
suffix_list = List("suffix_list", suffixes)

# Dictation consisting of sources of contextually likely words.
custom_dictation = Alternative([
    ListRef(None, saved_word_list),
    ListRef(None, context_phrase_list),
])

# Either arbitrary dictation or letters.
mixed_dictation = JoinedSequence(" ", [
    Optional(ListRef(None, prefix_list)),
    Alternative([
        Dictation(),
        DictListRef(None, letters_dict_list),
        ListRef(None, saved_word_list),
        ListRef(None, context_phrase_list),
    ]),
    Optional(ListRef(None, suffix_list))])

# A sequence of either short letters or long letters.
letters_element = JoinedRepetition("", DictListRef(None, letters_dict_list), min = 1, max = 10)

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

# Rule for formatting pure dictation elements.
pure_format_rule = create_rule(
    "PureFormatRule",
    dict([("pure " + k, v)
          for (k, v) in format_functions.items()]),
    {"dictation": Dictation()}
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
        "letters": DictListRef(None, letters_dict_list),
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
dictation_element = Alternative([
    RuleRef(rule=dictation_rule),
    RuleRef(rule=format_rule),
    RuleRef(rule=pure_format_rule),
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
    def __init__(self, name, command, terminal_command, context):
        # Here we define this rule's spoken-form and special elements. Note that
        # nested_repetitions is the only one that contains Repetitions, and it
        # is not itself repeated. This is for performance purposes. We also
        # include a special escape command "terminal <dictation>" in case
        # recognition problems occur with repeated dictation commands.
        spec     = "[<sequence>] [<nested_repetitions>] ([<dictation_sequence>] [terminal <dictation>] | <terminal_command>) [[[and] repeat [that]] <n> times]"
        extras   = [
            Repetition(command, min=1, max = 5, name="sequence"), 
            Alternative([RuleRef(rule=character_rule), RuleRef(rule=spell_format_rule)],
                        name="nested_repetitions"),
            Repetition(dictation_element, min = 1, max = 5, name = "dictation_sequence"), 
            ElementWrapper("dictation", dictation_element), 
            ElementWrapper("terminal_command", terminal_command),
            IntegerRef("n", 1, 100),  # Times to repeat the sequence.
        ]
        defaults = {
            "n": 1,                   # Default repeat count.
            "sequence": [], 
            "nested_repetitions": None, 
            "dictation_sequence": [],
            "dictation": None, 
            "terminal_command": None, 
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

#-------------------------------------------------------------------------------
# Define top-level rules for different contexts. Note that Dragon only allows
# top-level rules to be context-specific, but we want control over sub-rules. To
# work around this limitation, we compile a mutually exclusive top-level rule
# for each context.

class ContextHelper:
    """Helper to define a context hierarchy in terms of sub-rules but pass it to
    dragonfly as top-level rules."""

    def __init__(self, name, context, element, terminal_element=Empty()):
        """Associate the provided context with the element to be repeated."""
        self.name = name
        self.context = context
        self.element = element
        self.children = []
        self.terminal_element = terminal_element

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
                                    self.terminal_element,
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
    "five tap": Text("5tap "), 
    "git commit": Text("git commit -am "),
    "git commit done": Text("git commit -am done "),
    "git checkout new": Text("git checkout -b "),
    "git reset hard head": Text("git reset --hard HEAD "), 
    "(soft|sym) link": Text("ln -s "),
    "list": Text("ls -l "),
    "make dir": Text("mkdir "),
    "ps all": Text("ps aux "),
    "kill command": Text("kill "),
    "echo command": Text("echo "),
    "CH mod": Text("chmod "),
    "TK diff": Text("tkdiff "),
    "MV": Text("mv "),
    "CP": Text("cp "),
    "RM": Text("rm "),
    "CD": Text("cd "),
    "LS": Text("ls "),
    "PS": Text("ps "),
}, dict((command, Text(command + " ")) for command in [
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
]))


def Exec(command):
    return Key("c-c, a-x") + Text(command) + Key("enter")

# Work in progress.
def FastExec(command):
    return Function(lambda: urllib.urlopen("http://127.0.0.1:9091/" + command).close())

def jump_to_line(line_string):
    return Key("c-u") + Text(line_string) + Key("c-c, c, g")

class MarkLinesAction(ActionBase):
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

emacs_action_map = combine_maps(
    command_action_map,
    {
        "up [<n>]": Key("c-u") + Text("%(n)s") + Key("up"),
        "down [<n>]": Key("c-u") + Text("%(n)s") + Key("down"),
        "crack [<n>]": Key("c-u") + Text("%(n)s") + Key("c-d"),
        "kimble [<n>]": Key("c-u") + Text("%(n)s") + Key("as-d"),
        "select everything": Key("c-x, h"), 
        "edit everything": Key("c-x, h, c-w") + RunApp("notepad") + Key("c-v"),
        "edit region": Key("c-w") + RunApp("notepad") + Key("c-v"),
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
        "(save|set) bookmark": Key("c-x, r, m"),
        "list bookmarks": Key("c-x, r, l"),
        "indent region": Key("ca-backslash"), 
        "comment region": Key("a-semicolon"), 
        "project file": Key("c-c, p, f"),
        "switch project": Key("c-c, p, p"),
        "build file": Key("c-c/10, c-g"),
        "test file": Key("c-c, c-t"),
        "helm": Key("c-x, c"),
        "helm resume": Key("c-x, c, b"), 
        "full line <line>": Key("a-g, a-g") + Text("%(line)s") + Key("enter"),
        "line <n1>": jump_to_line("%(n1)s"),
        "open line <n1>": jump_to_line("%(n1)s") + Key("a-enter"),
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
        "no ido": Key("c-f"),
        "dear ed": Key("c-d"),
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
        "narrow region": Key("c-x, n, n"),
        "widen buffer": Key("c-x, n, w"),
        "(prev|preev) symbol": Key("a-s, dot, c-r, c-r"), 
        "(next|neck) symbol": Key("a-s, dot, c-s"),
        "jump before <context_word>": Key("c-r") + Text("%(context_word)s") + Key("enter"),
        "jump after <context_word>": Key("c-s") + Text("%(context_word)s") + Key("enter"),
        "next result": Key("a-comma"),
        "preev error": Key("f11"), 
        "next error": Key("f12"),
        "cut": Key("c-w"),
        "copy": Key("a-w"),
        "yank": Key("c-y"),
        "yank <n1> [(through|to) <n2>]": UseLinesAction(Key("a-w"), Key("c-y")),
        "yank tight <n1> [(through|to) <n2>]": UseLinesAction(Key("a-w"), Key("c-y"), True),
        "grab <n1> [(through|to) <n2>]": UseLinesAction(Key("c-w"), Key("c-y")),
        "grab tight <n1> [(through|to) <n2>]": UseLinesAction(Key("c-w"), Key("c-y"), True),
        "sank": Key("a-y"), 
        "Mark": Key("c-space"),
        "Mark <n1> [(through|to) <n2>]": MarkLinesAction(),
        "Mark tight <n1> [(through|to) <n2>]": MarkLinesAction(True),
        "nasper": Key("ca-f"),
        "pesper": Key("ca-b"),
        "moosper": Key("cas-2"),
        "kisper": Key("ca-k"),
        "dowsper": Key("ca-d"),
        "usper": Key("ca-u"),
        "fopper": Key("c-c, c, c-f"),
        "bapper": Key("c-c, c, c-b"), 
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
        "new template": Key("c-c, ampersand, c-n"),
        "reload templates": Exec("yas-reload-all"), 
        "prefix": Key("c-u"), 
        "quit": Key("q"),
        "save": Key("c-x, c-s"),
        "save as": Key("c-x, c-w"), 
        "open (definition|def)": Key("c-backtick, c-c, comma, d"),
        "toggle (definition|def)": Key("c-c, comma, D"),
        "open cross (references|ref)": Key("c-c, comma, x"),
        "open tag": Key("a-dot, enter"), 
        "clang format": Key("ca-q"),
        "format comment": Key("a-q"),
        "other top": Key("c-minus, ca-v"),
        "other pown": Key("ca-v"),
        "close tag": Key("c-c, c-e"),
        "I jump <char>": Key("c-c, c, j") + Text("%(char)s") + Function(lambda: type_position("%d\n%d\n")),
        "R grep": Exec("rgrep"),
        "code search": Exec("cs"),
        "code search car": Exec("csc"),
        "mark (reg|rej) <char>": Key("c-x, r, space, %(char)s"),
        "save mark <char>": Key("c-c, c, m, %(char)s"),
        "jump (reg|rej) <char>": Key("c-x, r, j, %(char)s"),
        "copy (reg|rej) <char>": Key("c-x, r, s, %(char)s"),
        "save copy <char>": Key("c-c, c, w, %(char)s"),
        "yank (reg|rej) <char>": Key("c-u, c-x, r, i, %(char)s"),
        "expand diff": Key("a-4"),
        "expand region": Key("c-equals"),
        "contract region": Key("c-plus"),
        "refresh": Key("g"),
        "JavaScript mode": Exec("js-mode"),
        "HTML mode": Exec("html-mode"),
        "toggle details": Exec("dired-hide-details-mode"),
        "up fub": Exec("windmove-up"), 
        "down fub": Exec("windmove-down"), 
        "left fub": Exec("windmove-left"), 
        "right fub": Exec("windmove-right"),
        "split header": Key("c-x, 3, c-x, o, c-x, c-h"),
        "header": Key("c-x, c-h"),
        "create shell": Exec("shell"),
        "hello world": FastExec("hello-world"),
        "kill emacs server": Exec("ws-stop-all"), 
        "closure compile": Key("c-c, c-k"),
        "closure namespace": Key("c-c, a-n"),
        "pie flakes": Key("c-c, c-v"),
        "help variable": Key("c-h, v"), 
        "help function": Key("c-h, f"), 
        "help key": Key("c-h, k"),
        "help mode": Key("c-h, m"), 
        "help back": Key("c-c, c-b"), 
        "eval defun": Key("ca-x"),
        "eval region": Exec("eval-region"),
        "magit status": Key("c-c, m"),
        "submit comment": Key("c-c, c-c"),
        "show diff": Key("c-x, v, equals"),
        "recompile": Exec("recompile"),
        "customize": Exec("customize-apropos"),
        "python indent": Key("c-c, rangle"),
        "python dedent": Key("c-c, langle"),
    })

templates = {
    "beginend": "beginend",
    "car": "car",
    "class": "class",
    "const ref": "const_ref", 
    "def": "function",
    "each": "each",
    "error": "error", 
    "eval": "eval",
    "for": "for",
    "fun declaration": "fun_declaration",
    "function": "function",
    "if": "if",
    "info": "info", 
    "method": "method", 
    "ref": "ref",
    "to do": "todo",
    "var": "vardef",
    "warning": "warning",
    "while": "while",
    "unique pointer": "unique_pointer",
    "inverse if": "inverse_if",
    "else": "else",
    "ternary": "ternary", 
    "key": "key",
    "text": "text",
    "entry": "entry",
}
template_dict_list = DictList("template_dict_list", templates)
emacs_element_map = combine_maps(
    keystroke_element_map,
    {
        "n1": IntegerRef(None, 0, 100),
        "n2": IntegerRef(None, 0, 100), 
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
        "shell (preev|back)": Key("a-r"),
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
        "frink": Key("apostrophe"),
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
        "go to drive": Key("c-l/15") + Text("drive.google.com") + Key("enter"),
        "go to amazon": Key("c-l/15") + Text("smile.amazon.com") + Key("enter"),
        "(new|insert) row": Key("a-i/15, r"),
        "delete row": Key("a-e/15, d"),
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
        "create driver": Function(create_driver),
        "quit driver": Function(quit_driver),
        "test driver": Function(test_driver),
        "search bar": ClickElementAction(By.NAME, "q"),
        "add bill": ClickElementAction(By.LINK_TEXT, "Add a bill"),
    })
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
chrome_element_map = combine_maps(
    keystroke_element_map,
    {
        "tab_n": IntegerRef(None, 1, 9),
        "link": JoinedRepetition("", DictListRef(None, link_char_dict_list), min = 0, max = 5),
    })

chrome_element = RuleRef(rule=create_rule("ChromeKeystrokeRule", chrome_action_map, chrome_element_map))
chrome_terminal_element = RuleRef(rule=create_rule("ChromeTerminalRule", chrome_terminal_action_map, chrome_element_map))
chrome_context_helper = ContextHelper("Chrome", AppContext(executable="chrome"),
                                      chrome_element, chrome_terminal_element)
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
        "resolve": Key("c-j"), 
        "done": Key("d"),
        "save": Key("c-s"),
        "expand|collapse": Key("e"),
        "reply": Key("r"),
        "comment <line_n>": DoubleClickElementAction(By.XPATH,
                                                     ("//span[contains(@class, 'stx-line') and "
                                                      "starts-with(@id, 'c') and "
                                                      "substring-after(@id, '_') = '%(line_n)s']")), 
    })
critique_element_map = combine_maps(
    chrome_element_map,
    {
        "line_n": IntegerRef(None, 1, 10000),
    })
critique_element = RuleRef(rule=create_rule("CritiqueKeystrokeRule", critique_action_map, critique_element_map))
critique_context_helper = ContextHelper("Critique", AppContext(title = "<critique.corp.google.com>"),
                                        critique_element, chrome_terminal_element)
chrome_context_helper.add_child(critique_context_helper)

calendar_action_map = combine_maps(
    chrome_action_map,
    {
        "click <name>": ClickElementAction(By.XPATH, "//*[@role='option' and contains(string(.), '%(name)s')]"),
        "today": Key("t"),
        "preev": Key("k"),
        "next": Key("j"),
        "day": Key("d"),
        "week": Key("w"),
        "month": Key("m"),
    })
names_dict_list = DictList(
    "name_dict_list",
    {
        "Sonica": "Sonica"
    })
calendar_element_map = combine_maps(
    chrome_element_map,
    {
        "name": DictListRef(None, names_dict_list),
    })
calendar_element = RuleRef(rule=create_rule("CalendarKeystrokeRule", calendar_action_map, calendar_element_map))
calendar_context_helper = ContextHelper("Calendar", AppContext(title = "Google Calendar") | AppContext(title = "Google.com - Calendar"),
                                        calendar_element, chrome_terminal_element)
chrome_context_helper.add_child(calendar_context_helper)

code_search_action_map = combine_maps(
    chrome_action_map,
    {
        "header": Key("r/25, h"),
        "source": Key("r/25, c"), 
    })
code_search_element = RuleRef(rule=create_rule("CodeSearchKeystrokeRule", code_search_action_map, chrome_element_map))
code_search_context_helper = ContextHelper("CodeSearch", AppContext(title = "<cs.corp.google.com>"),
                                           code_search_element, chrome_terminal_element)
chrome_context_helper.add_child(code_search_context_helper)

gmail_action_map = combine_maps(
    chrome_action_map,
    {
        "open": Key("o"),
        "(archive|done)": Text("{"),
        "mark unread": Text("_"),
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
        "label waiting": Key("l/25") + Text("waiting") + Key("enter"),
        "select": Key("x"),
        "select next <n>": Key("x, j") * Repeat(extra="n"), 
        "new messages": Key("N"),
        "go to inbox": Key("g, i"), 
        "go to starred": Key("g, s"), 
        "go to sent": Key("g, t"),
        "go to drafts": Key("g, d"),
        "expand all": ClickElementAction(By.XPATH, "//*[@aria-label='Expand all']"),
        "click to": ClickElementAction(By.XPATH, "//*[@aria-label='To']"),
        "click cc": Key("cs-c"),
        "open chat": Key("q"),
    })
gmail_terminal_action_map = combine_maps(
    chrome_terminal_action_map,
    {
        "chat with <text>": Key("q/50") + Text("%(text)s") + Pause("50") + Key("enter"),
    })

gmail_element = RuleRef(rule=create_rule("GmailKeystrokeRule", gmail_action_map, chrome_element_map))
gmail_terminal_element = RuleRef(rule=create_rule("GmailTerminalRule", gmail_terminal_action_map, chrome_element_map))
gmail_context_helper = ContextHelper("Gmail",
                                     (AppContext(title = "Gmail") |
                                      AppContext(title = "Google.com Mail") |
                                      AppContext(title = "<mail.google.com>") |
                                      AppContext(title = "<inbox.google.com>")),
                                     gmail_element, gmail_terminal_element)
chrome_context_helper.add_child(gmail_context_helper)

docs_action_map = combine_maps(
    chrome_action_map,
    {
        "select column": Key("c-space"), 
        "select row": Key("s-space"),
        "row up": Key("a-e/15, k"), 
        "row down": Key("a-e/15, j"),
        "column left": Key("a-e/15, m"), 
        "column right": Key("a-e/15, m"),
        "add comment": Key("ca-m"), 
    })
docs_element = RuleRef(rule=create_rule("DocsKeystrokeRule", docs_action_map, chrome_element_map))
docs_context_helper = ContextHelper("Docs",
                                    AppContext(title = "<docs.google.com>"),
                                    docs_element, chrome_terminal_element)
chrome_context_helper.add_child(docs_context_helper)

def go_to_hotlist(hotlist):
    return Key("h") + Text(hotlist) + Key("down, enter");

buganizer_action_map = combine_maps(
    chrome_action_map,
    {
        "next": Key("j"),
        "preev": Key("k"),
        "list": Key("u"),
        "assign": Key("a"),
        "duplicate": Key("d"),
        "bug ID": Key("i"),
        "comment|reply": Key("r"),
        "priority <n>": Key("p, %(n)d"),
        "open": Key("o"),
        "select": Key("x"),
        "hotlist OKRs": go_to_hotlist("okrs"),
        "hotlist task queue": go_to_hotlist("task queue"),
    })
buganizer_element = RuleRef(rule=create_rule("BuganizerKeystrokeRule", buganizer_action_map, chrome_element_map))
buganizer_context_helper = ContextHelper("Buganizer",
                                         AppContext(title = "<b.corp.google.com>") |
                                         AppContext(title = "<buganizer.corp.google.com>") |
                                         AppContext(title = "<b2.corp.google.com>"),
                                         buganizer_element, chrome_terminal_element)
chrome_context_helper.add_child(buganizer_context_helper)

analog_action_map = combine_maps(
    chrome_action_map,
    {
        "next": Key("n"),
        "preev": Key("p"),
    })
analog_element = RuleRef(rule=create_rule("AnalogKeystrokeRule", analog_action_map, chrome_element_map))
analog_context_helper = ContextHelper("Analog",
                                      AppContext(title = "<analog.corp.google.com>"),
                                      analog_element, chrome_terminal_element)
chrome_context_helper.add_child(analog_context_helper)

notepad_action_map = combine_maps(
    command_action_map,
    {
        "dumbbell": Key("shift:down, c-left, backspace, shift:up"),
        "transfer out": Key("c-a, c-x, a-f4") + UniversalPaste(),
    })

notepad_element = RuleRef(rule=create_rule("NotepadKeystrokeRule", notepad_action_map, keystroke_element_map))
notepad_context_helper = ContextHelper("Notepad", AppContext(executable = "notepad"), notepad_element)
global_context_helper.add_child(notepad_context_helper)

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
        words = ExtractWords(text, file_type)
        phrases = ExtractPhrases(text, file_type)
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
