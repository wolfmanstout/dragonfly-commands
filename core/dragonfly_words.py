#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Library for extracting words and phrases from text."""

import sys
import re
import os
import fileinput
from dragonfly_local import *

WORDS_PATH = HOME + "/dotfiles/words.txt"
BLACKLIST_PATH = HOME + "/dotfiles/blacklist.txt"

def SplitDictation(dictation):
    """Preprocess dictation to do a better job of word separation. Returns a list of
    words."""
    # Convert the input to a list of lowercase words, removing punctuation other
    # than "." within words.
    word_punctuation_pattern = r"[^\w. ]"
    raw_words = [word for word
                 in re.sub(word_punctuation_pattern, "", str(dictation)).lower().split(" ")
                 if len(word) > 0]

    # Merge contiguous letters into a single word, and merge words separated by
    # punctuation marks into a single word. This way we can dictate something
    # like "score test case dot start now" and only have the underscores applied
    # at word boundaries, to produce "test_case.start_now".
    words = []
    previous_letter = False
    previous_punctuation = False
    punctuation_pattern = r"\W"
    for word in raw_words:
        current_punctuation = re.match(punctuation_pattern, word)
        current_letter = len(word) == 1 and not re.match(punctuation_pattern, word)
        if len(words) == 0:
            words.append(word)
        else:
            if current_punctuation or previous_punctuation or (current_letter and previous_letter):
                words.append(words.pop() + word)
            else:
                words.append(word)
        previous_letter = current_letter
        previous_punctuation = current_punctuation
    return words

def ParseWords(path):
  words = set()
  with open(path) as words_file:
    for line in words_file:
      words.add(line.strip())
  return words

def SaveWords(path, words):
  with open(path, "w") as words_file:
    for word in sorted(words):
      words_file.write(word + "\n")

def RemovePlaintext(text, file_type = None):
  if file_type == "py":
    text = re.sub(re.compile(r"#.*$", re.MULTILINE), "", text)
  if file_type == "el":
    text = re.sub(re.compile(r";.*$", re.MULTILINE), "", text)
  if file_type == "cc" or file_type == "h":
    text = re.sub(re.compile(r"//.*$", re.MULTILINE), "", text)
  text = re.sub(re.compile(r"\".*?\"", re.MULTILINE), "", text)
  return text

def RemoveBlacklistWords(words):
  try:
    blacklist_words = ParseWords(BLACKLIST_PATH)
  except:
    print("Unable to open: " + BLACKLIST_PATH)
    blacklist_words = set()
  return words - blacklist_words

def GetWords(text):
  return [word.lower() for word in re.findall(r"([A-Z][a-z]*|[a-z]+)", text)]

def ExtractWords(text, file_type = None):
  text = RemovePlaintext(text, file_type)
  words = set(GetWords(text))
  return RemoveBlacklistWords(words)

def ExtractPhrases(text, file_type = None):
  text = RemovePlaintext(text, file_type)
  words = set([" ".join(GetWords(phrase)) for phrase in re.findall(r"[A-z_-]+", text)])
  return RemoveBlacklistWords(words)
