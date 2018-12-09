# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Grammar for loading a macro directory backup. WARNING: Renaming sometimes
quietly fails for unknown reasons."""

from dragonfly import (Function, Grammar, MappingRule)
import os


#---------------------------------------------------------------------------
# Create the main command rule.

def swap_macros():
    os.rename("c:/NatLink/NatLink/MacroSystem", "c:/NatLink/NatLink/MacroSystem.temp")
    os.rename("c:/NatLink/NatLink/MacroSystem.other", "c:/NatLink/NatLink/MacroSystem")
    os.rename("c:/NatLink/NatLink/MacroSystem.temp", "c:/NatLink/NatLink/MacroSystem.other")


class CommandRule(MappingRule):
    mapping = {
        "swap macros directory": Function(swap_macros),
    }


grammar = Grammar("safety")
grammar.add_rule(CommandRule())
grammar.load()


# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
