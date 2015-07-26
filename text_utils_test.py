#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

from _text_utils import *
import unittest

class TextUtilsTestCase(unittest.TestCase):

    def setUp(self):
        pass

    def test_extract_words(self):
        self.assertEqual(set(["test", "word"]), ExtractWords("testWord"))
        self.assertEqual(set(["test", "word"]), ExtractWords("test_word"))
        self.assertEqual(set(["q", "test", "word"]), ExtractWords("QTestWord"))
        # self.assertEqual(set(["test", "word"]), ExtractWords("TEST_WORD"))
        self.assertEqual(set(["test", "word", "another"]), ExtractWords("test_word another_word"))
        self.assertEqual(set(["test", "word"]), ExtractWords("test_word // another_word", "cc"))
        self.assertEqual(set(["return", "key", "text", "command"]), ExtractWords("""return Key("a-x") + Text(command) + Key("enter")""", "py"))

    def test_extract_phrases(self):
        self.assertEqual(set(["test word"]), ExtractPhrases("testWord"))
        self.assertEqual(set(["test word"]), ExtractPhrases("test_word"))
        self.assertEqual(set(["q test word"]), ExtractPhrases("QTestWord"))
        self.assertEqual(set(["test word", "another word"]), ExtractPhrases("test_word another_word"))
        self.assertEqual(set(["test word"]), ExtractPhrases("test_word // another_word", "cc"))
        self.assertEqual(set(["test word"]), ExtractPhrases("test-word ; another_word", "el"))

    def test_split_dictation(self):
        self.assertEqual(["test", "word"], SplitDictation("test word"))
        self.assertEqual(["test", "word", "ab"], SplitDictation("test word A B"))
        self.assertEqual(["test\word"], SplitDictation("test\word"))
        self.assertEqual(["test", "word"], SplitDictation("test-word"))
        self.assertEqual(["test/word"], SplitDictation("test/word"))
        self.assertEqual(["test/word"], SplitDictation("test / word"))
        self.assertEqual(["test/word"], SplitDictation("test/ word"))
        self.assertEqual(["test/word"], SplitDictation("test /word"))
        self.assertEqual(["joes"], SplitDictation("Joe's"))
        self.assertEqual(["test", "case.start", "now"], SplitDictation("test case.start now"))
    

if __name__ == "__main__":
    unittest.main()
