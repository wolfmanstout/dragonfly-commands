# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")
except ImportError:
    pass

from dragonfly import (
    Dictation,
    Grammar,
    IntegerRef,
    Key,
    MappingRule,
    Text,
)
import _linux_utils as linux


def Exec(command):
    return Key("a-x") + Text(command) + Key("enter")


class CommandRule(MappingRule):
    mapping = {
        "dragonfly add buffer": Exec("dragonfly-add-buffer"),
        "dragonfly add word": Exec("dragonfly-add-word"),
        "dragonfly blacklist word": Exec("dragonfly-blacklist-word"),
        "Foreclosure next": Exec("4clojure-next-question"),
        "Foreclosure previous": Exec("4clojure-previous-question"),
        "Foreclosure check": Exec("4clojure-check-answers"),
        "confirm": Text("yes") + Key("enter"),
        "confirm short": Text("y"),
        "deny": Text("no") + Key("enter"),
        "deny short": Text("n"),
        "relative line numbers": Exec("linum-relative-toggle"),
        "buff revert": Exec("revert-buffer"),
        "emacs close now": Key("c-x, c-c"),
        }
    extras = [
        IntegerRef("n", 1, 20),
        IntegerRef("line", 1, 10000),
        Dictation("text"),
        ]
    defaults = {
        "n": 1,
        }


context = linux.UniversalAppContext(title = "Emacs editor")
grammar = Grammar("Emacs", context=context)
grammar.add_rule(CommandRule())
grammar.load()


# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
