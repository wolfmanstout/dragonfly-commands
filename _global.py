# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Global commands where we expressly want to avoid repetition. This is
especially useful for window switching commands, where can be disastrous if the
window is changed accidentally in the middle of command execution."""

import platform

from dragonfly import *
from _dragonfly_utils import *

#-------------------------------------------------------------------------------
# Create the main command rule.

# Work around security restrictions in Windows 8.
if platform.release() == "8":
    swap_action = Mimic("press", "alt", "tab")
else:
    swap_action = Key("alt:down, tab:%(n)d/25, alt:up")

class CommandRule(MappingRule):
    mapping = {
        "swap [<n>]": swap_action,
    }
    extras = [
        IntegerRef("n", 1, 20),
        ]
    defaults = {
        "n": 1,
        }

#-------------------------------------------------------------------------------
# Create commands to jump to a specific window.

# Ordered list of pinned taskbar items. Sublists refer to windows within a specific application.
windows = [
    "explorer",
    ["dragonbar", "dragon [messages]", "dragonpad"],
    "home chrome",
    "home terminal",
    "home emacs",
]
json_windows = load_json("windows.json")
if json_windows:
    windows = json_windows

windows_prefix = "go to"
windows_mapping = {}
for i, window in enumerate(windows):
    if isinstance(window, str):
        window = [window]
    for j, words in enumerate(window):
        windows_mapping[windows_prefix + " " + words] = Key("win:down, %d:%d/10, win:up" % (i + 1, j + 1))

class WindowsRule(MappingRule):
    mapping = windows_mapping

#-------------------------------------------------------------------------------
# Global grammar.

grammar = Grammar("global")
grammar.add_rule(CommandRule())
grammar.add_rule(WindowsRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
