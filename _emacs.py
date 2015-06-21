# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

try:
    import pkg_resources
    pkg_resources.require("dragonfly >= 0.6.5beta1.dev-r76")
except ImportError:
    pass

from dragonfly import *

def Exec(command):
    return Key("a-x") + Text(command) + Key("enter")

class CommandRule(MappingRule):
    mapping = {
        "save as": Key("c-x, c-w"), 
        "(save|set) bookmark": Key("c-x, r, m"),
        "list bookmarks": Key("c-x, r, l"),
        "help variable": Key("c-h, v"), 
        "help function": Key("c-h, f"), 
        "help key": Key("c-h, k"),
        "help mode": Key("c-h, m"), 
        "help back": Key("c-c, c-b"), 
        "dragonfly add buffer": Exec("dragonfly-add-buffer"),
        "dragonfly add word": Exec("dragonfly-add-word"),
        "dragonfly blacklist word": Exec("dragonfly-blacklist-word"), 
        "eval defun": Key("ca-x"),
        "eval region": Exec("eval-region"), 
        "Foreclosure next": Exec("4clojure-next-question"), 
        "Foreclosure previous": Exec("4clojure-previous-question"), 
        "Foreclosure check": Exec("4clojure-check-answers"),
        "submit comment": Key("c-c, c-c"),
        "new template": Key("c-c, ampersand, c-n"),
        "reload templates": Exec("yas-reload-all"), 
        "confirm": Text("yes") + Key("enter"),
        "deny": Text("no") + Key("enter"),
        "magit status": Exec("magit-status"),
        "relative line numbers": Exec("linum-relative-toggle"),
        "revert buffer": Exec("revert-buffer"),
        "split header": Key("c-x, 3, c-x, o, c-x, c-h"),
        "header": Key("c-x, c-h"),
        "create shell": Exec("shell"),
        "exit out of Emacs": Key("c-x, c-c"),
        "recompile": Exec("recompile"),
        "show diff": Key("c-x, v, equals"),
        "customize": Exec("customize-apropos"), 
        }
    extras = [
        IntegerRef("n", 1, 20),
        IntegerRef("line", 1, 10000),
        Dictation("text"),
        ]
    defaults = {
        "n": 1,
        }

context = AppContext(title = "Emacs editor")
grammar = Grammar("Emacs", context=context)
grammar.add_rule(CommandRule())
grammar.load()

# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
