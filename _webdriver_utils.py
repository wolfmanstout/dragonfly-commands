#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Actions for manipulating Chrome via WebDriver."""

import json
import semantic_locators
from six.moves import urllib_request
from six.moves import urllib_error

from dragonfly import (DynStrActionBase)
import _dragonfly_local as local

# Needed for marionette_driver.
import sys
has_argv = hasattr(sys, "argv")
if not has_argv:
    sys.argv = [""]
# Native driver for Firefox; more reliable.
import marionette_driver

from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains


class MarionetteWrapper(object):
    """Wraps a Marionette driver instance to make it compatible with the WebDriver API."""

    def __init__(self, delegate):
        self.delegate = delegate

    def __getattr__(self, name):
        try:
            return super().__getattr__(self, name)
        except AttributeError:
            return getattr(self.delegate, name)

    def execute_script(self, script, *args):
        return self.delegate.execute_script(script, args, new_sandbox=False)


def create_driver():
    global driver, browser
    driver = None
    browser = local.DEFAULT_BROWSER
    if browser == "chrome":
        try:
            urllib_request.urlopen("http://127.0.0.1:9222/json")
        except urllib_error.URLError:
            print("Unable to start WebDriver, Chrome is not responding.")
            return
        chrome_options = webdriver.chrome.options.Options()
        chrome_options.experimental_options["debuggerAddress"] = "127.0.0.1:9222"
        driver = webdriver.Chrome(local.CHROME_DRIVER_PATH, chrome_options=chrome_options)
    elif browser == "firefox":
        driver = MarionetteWrapper(marionette_driver.marionette.Marionette())
    else:
        print("Unknown browser: " + browser)
        browser = None


def quit_driver():
    global driver, browser
    if driver:
        if browser == "chrome":
            driver.quit()
        elif browser == "firefox":
            driver.delete_session()
    driver = None
    browser = None


def switch_to_active_tab():
    if browser == "chrome":
        tabs = json.load(urllib_request.urlopen("http://127.0.0.1:9222/json"))
        # Chrome seems to order the tabs by when they were last updated, so we find
        # the first one that is not an extension.
        for tab in tabs:
            if not tab["url"].startswith("chrome-extension://"):
                active_tab = tab["id"]
                break
        for window in driver.window_handles:
            # ChromeDriver adds to the raw ID, so we just look for substring match.
            if active_tab in window:
                driver.switch_to_window(window);
                print("Switched Chrome to: " + driver.title.encode('ascii', 'backslashreplace').decode())
                return
        print("Did not find active tab in Chrome.")
    elif browser == "firefox":
        driver.delete_session()
        driver.start_session()


def test_driver():
    switch_to_active_tab()
    driver.get('http://www.google.com/xhtml');


def find_nearest_element(get_elements_function, tracker):
    # Get gaze location as early as possible.
    gaze_location = tracker.get_gaze_point_or_default()
    switch_to_active_tab()
    elements = get_elements_function()
    if not elements:
        print("No matching elements found")
        return
    # Assume there is equal amount of browser chrome on the left and right sides of the screen.
    canvas_left = driver.execute_script("return window.screenX + (window.outerWidth - window.innerWidth) / 2 - window.scrollX;")
    # Assume all the browser chrome is on the top of the screen and none on the bottom.
    canvas_top = driver.execute_script("return window.screenY + (window.outerHeight - window.innerHeight) - window.scrollY;")
    canvas_bottom = driver.execute_script("return window.screenY + window.outerHeight - window.scrollY;")
    nearest_element = None
    nearest_element_distance_squared = float("inf")
    for element in elements:
        element_location = (element.rect["x"] + canvas_left + element.rect["width"] / 2,
                            element.rect["y"] + canvas_top + element.rect["height"] / 2)
        if not element.is_displayed():
            # Element is not visible onscreen.
            continue
        offsets = (element_location[0] - gaze_location[0], element_location[1] - gaze_location[1])
        distance_squared = offsets[0] * offsets[0] + offsets[1] * offsets[1]
        if distance_squared < nearest_element_distance_squared:
            nearest_element = element
            nearest_element_distance_squared = distance_squared
    if not nearest_element:
        print("Matching elements were not visible")
        return
    return nearest_element


def find_clickable_elements_by_name(name):
    elements = []
    for role in ("button", "tab"):
        for name_candidate in (name, name.capitalize()):
            elements.extend(semantic_locators.find_elements_by_semantic_locator(
                driver, "{{{} '*{}*'}}".format(role, name_candidate)))
    return elements


def click_element(element):
    if browser == "chrome":
        try:
            element.click()
        except:
            # Move then click to avoid Chrome unwillingness to send a click
            # which reaches an overlapping element.
            ActionChains(driver).move_to_element(element).click().perform()
    elif browser == "firefox":
        element.click()


def double_click_element(element):
    if browser == "chrome":
        try:
            element.double_click()
        except:
            # Move then click to avoid Chrome unwillingness to send a click
            # which reaches an overlapping element.
            ActionChains(driver).move_to_element(element).double_click().perform()
    elif browser == "firefox":
        element.double_click()


class ElementAction(DynStrActionBase):

    def __init__(self, by, spec):
        DynStrActionBase.__init__(self, spec)
        self.by = by

    def _parse_spec(self, spec):
        return spec

    def _execute_events(self, events):
        switch_to_active_tab()
        element = driver.find_element(self.by, events)
        self._execute_on_element(element)


class ClickElementAction(ElementAction):

    def _execute_on_element(self, element):
        click_element(element)


class SmartElementAction(DynStrActionBase):

    def __init__(self, by, spec, tracker):
        DynStrActionBase.__init__(self, spec)
        self.by = by
        self.tracker = tracker

    def _parse_spec(self, spec):
        return spec

    def _execute_events(self, events):
        nearest_element = find_nearest_element(lambda: driver.find_elements(self.by, events), tracker)
        if nearest_element:
            self._execute_on_element(nearest_element)


class SmartClickElementAction(SmartElementAction):

    def _execute_on_element(self, element):
        click_element(element)


class DoubleClickElementAction(ElementAction):

    def _execute_on_element(self, element):
        double_click_element(element)
