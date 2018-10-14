from dragonfly import *


# grammar = Grammar("test_grammar")
# grammar.load()


# Unload function which will be called by natlink at unload time.
def unload():
    global grammar
    if grammar: grammar.unload()
    grammar = None
