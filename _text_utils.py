#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Library for extracting words and phrases from text."""

import re
import _dragonfly_local as local

WORDS_PATH = local.HOME + "/dotfiles/words.txt"
BLACKLIST_PATH = local.HOME + "/dotfiles/blacklist.txt"


def split_dictation(dictation):
    """Preprocess dictation to do a better job of word separation. Returns a list of
    words."""
    # Make lowercase.
    clean_dictation = str(dictation).lower()
    # Strip apostrophe.
    clean_dictation = re.sub(r"'", "", clean_dictation)
    # Convert dashes into spaces.
    clean_dictation = re.sub(r"-", " ", clean_dictation)
    # Surround all other punctuation marks with spaces.
    clean_dictation = re.sub(r"(\W)", r" \1 ", clean_dictation)
    # Convert the input to a list of words and punctuation marks.
    raw_words = [word for word
                 in clean_dictation.split(" ")
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


def parse_words(path):
  words = set()
  with open(path) as words_file:
    for line in words_file:
      words.add(line.strip())
  return words


def save_words(path, words):
  with open(path, "w") as words_file:
    for word in sorted(words):
      words_file.write(word + "\n")


def remove_plaintext(text, file_type=None):
  if file_type == "py":
    text = re.sub(re.compile(r"#.*$", re.MULTILINE), "", text)
  if file_type == "el":
    text = re.sub(re.compile(r";.*$", re.MULTILINE), "", text)
  if file_type == "cc" or file_type == "h":
    text = re.sub(re.compile(r"//.*$", re.MULTILINE), "", text)
  text = re.sub(re.compile(r"\".*?\"", re.MULTILINE), "", text)
  return text


def remove_blacklist_words(words):
  try:
    blacklist_words = parse_words(BLACKLIST_PATH)
  except:
    print("Unable to open: " + BLACKLIST_PATH)
    blacklist_words = set()
  return words - blacklist_words


def get_words(text):
  # Discard "k" which can be a prefix for constants and rarely occurs elsewhere.
  return [word.lower() for word in re.findall(r"([A-Z][a-z]+|[a-z]+|[A-Z]+(?![a-z]))", text)
          if word != "k"]


def extract_words(text, file_type=None):
  text = remove_plaintext(text, file_type)
  words = set(get_words(text))
  return remove_blacklist_words(words)


def extract_phrases(text, file_type=None):
  text = remove_plaintext(text, file_type)
  words = set([" ".join(get_words(phrase)) for phrase in re.findall(r"[A-z_-]+", text)])
  return remove_blacklist_words(words)
