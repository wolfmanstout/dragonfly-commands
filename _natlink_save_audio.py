# Adapted by Danesprite from natlink grammar example file
# "_samplehypothesis.py" and the original dragonfly-save-audio patch by dwks.
# For more info, see https://github.com/dwks/dragonfly-save-audio
#
# Natlink grammar module to save audio and words from natlink into a
# specified directory. This module generates .txt files with the words and
# recognition type. For example:
#
#   Words: hello world
#   Type: dictation
#
# 'SAVE_DIR' must be set to an existing absolute directory, otherwise this
# module will raise errors when loaded.
#
# There four recognition types:
#
#   1. "dictation" - directory for speech that used dictation words only.
#   2. "grammar" - directory for speech that used grammar words only.
#   3. "mixed" - directory for speech that used grammar and dictation words.
#   4. "rejects" - directory for speech data recognised as noise.
#
# Saving rejects is disabled by default because incomplete utterances for
# grammar rules will also be included, which may not be desirable.
#
# This grammar's rules can be used to start and stop recording audio or
# noise.
#
# TODO Remove Dragon's formatting from dictation output.
# E.g. "\cap\cap" -> "cap"

import os
import time

import natlink
from natlinkutils import GrammarBase

import _dragonfly_local as local


class SaveAudioGrammar(GrammarBase):

    gramSpec = """
        <RecordState> exported = (start|stop) saving audio;
        <NoiseState> exported = (start|stop) saving (noise|rejects);
    """

    def initialize(self):
        self.load(self.gramSpec, hypothesis=1, allResults=1)
        self.activateAll()
        self.enabled = True  # Start saving audio by default.
        self.saveRejects = False

    @classmethod
    def getResultType(cls, details, resObj):
        if details == "reject":
            return "reject"

        # Check the rules for each word.
        # Anything between 1 and 1000000 (exclusive) is a grammar word.
        # 0 is used as the rule number for free-form dictation.
        rules = [r for _, r in resObj.getResults(0)]
        grammar_words = len([r for r in rules if 0 < r < 1000000])
        if grammar_words > 0:
            return "mixed" if 1000000 in rules else "grammar"
        else:
            return "dictation"

    def handleSelfResults(self, resObj):
        # Handle results for this grammar using the first word and rule ID.
        word, ruleId = resObj.getResults(0)[0]

        # Start/stop recording audio/rejects.
        value = word == "start"
        if ruleId == 1:
            self.enabled = value
        elif ruleId == 2:
            self.saveRejects = value

        print(" ".join(resObj.getWords(0)))

    def gotResultsObject(self, details, resObj):
        # Get any words from the result object.
        try:
            if details == "reject":
                words = "<???>"
            else:
                words = " ".join(resObj.getWords(0))
        except (natlink.OutOfRange, IndexError):
            details = "reject"
            words = "<???>"

        # Get the audio from the result object if there is any.
        try:
            wav = resObj.getWave()
        except natlink.DataMissing:
            wav = []

        # Save audio if there is any and if SAVE_DIR exists.
        # Only save rejects if specified.
        isReject = words == "<???>"
        shouldSave = (
            len(wav) > 0 and self.enabled and os.path.isdir(SAVE_DIR)
            and (not isReject or isReject and self.saveRejects)
        )

        if shouldSave:
            name = "rec-%.03f" % time.time()
            path = os.path.join(SAVE_DIR, name + ".wav")
            f = open(path, "wb")
            f.write(wav)
            f.close()
            path = os.path.join(SAVE_DIR, name + ".txt")
            f = open(path, "w")
            f.write("Words: %s\n" % str(words))
            f.write("Type: %s\n" % self.getResultType(details, resObj))
            f.close()

        # Handle this grammar's control rules after recording.
        if details == "self":
            self.handleSelfResults(resObj)


# The directory to save .wav and .txt files into.
# Must exist and be an absolute path.
SAVE_DIR = local.SAVE_AUDIO_DIR
if not os.path.isabs(SAVE_DIR) or not os.path.isdir(SAVE_DIR):
    grammar = None
    print("Not saving audio.")
else:
    # Instantiate and load the grammar.
    grammar = SaveAudioGrammar()
    grammar.initialize()
    print("Saving audio.")


def unload():
    global grammar
    if grammar:
        grammar.unload()
    grammar = None
