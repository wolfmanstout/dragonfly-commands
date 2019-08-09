#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

from _text_utils import *
import unittest


class TextUtilsTestCase(unittest.TestCase):

    def setUp(self):
        pass

    def test_extract_words(self):
        self.assertEqual(set(["test", "word"]), extract_words("testWord"))
        self.assertEqual(set(["test", "word"]), extract_words("test_word"))
        self.assertEqual(set(["q", "test", "word"]), extract_words("QTestWord"))
        # self.assertEqual(set(["test", "word"]), extract_words("TEST_WORD"))
        self.assertEqual(set(["test", "word", "another"]), extract_words("test_word another_word"))
        self.assertEqual(set(["test", "word"]), extract_words("test_word // another_word", "cc"))
        self.assertEqual(set(["return", "key", "text", "command"]), extract_words("""return Key("a-x") + Text(command) + Key("enter")""", "py"))

    def test_extract_phrases(self):
        self.assertEqual(set(["test word"]), extract_phrases("testWord"))
        self.assertEqual(set(["test word"]), extract_phrases("test_word"))
        self.assertEqual(set(["test word"]), extract_phrases("kTestWord"))
        self.assertEqual(set(["test word"]), extract_phrases("TEST_WORD"))
        self.assertEqual(set(["q test word"]), extract_phrases("QTestWord"))
        self.assertEqual(set(["test word", "another word"]), extract_phrases("test_word another_word"))
        self.assertEqual(set(["test word"]), extract_phrases("test_word // another_word", "cc"))
        self.assertEqual(set(["test word"]), extract_phrases("test-word ; another_word", "el"))

    def test_split_dictation(self):
        self.assertEqual(["test", "word"], split_dictation("test word"))
        self.assertEqual(["test", "word", "ab"], split_dictation("test word A B"))
        self.assertEqual(["test\word"], split_dictation("test\word"))
        self.assertEqual(["test", "word"], split_dictation("test-word"))
        self.assertEqual(["test/word"], split_dictation("test/word"))
        self.assertEqual(["test/word"], split_dictation("test / word"))
        self.assertEqual(["test/word"], split_dictation("test/ word"))
        self.assertEqual(["test/word"], split_dictation("test /word"))
        self.assertEqual(["joes"], split_dictation("Joe's"))
        self.assertEqual(["test", "case.start", "now"], split_dictation("test case.start now"))
    

if __name__ == "__main__":
    unittest.main()
