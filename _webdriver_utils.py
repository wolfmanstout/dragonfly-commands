#!/usr/bin/env python
# (c) Copyright 2015 by James Stout
# Licensed under the LGPL, see <http://www.gnu.org/licenses/>

"""Actions for manipulating Chrome via WebDriver."""

import json
from six.moves import urllib_request
from six.moves import urllib_error

from dragonfly import (DynStrActionBase)
import _dragonfly_local as local

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains

def create_driver():
    global driver
    driver = None
    try:
        urllib_request.urlopen("http://127.0.0.1:9222/json")
    except urllib_error.URLError:
        print("Unable to start WebDriver, Chrome is not responding.")
        return
    chrome_options = Options()
    chrome_options.experimental_options["debuggerAddress"] = "127.0.0.1:9222"
    driver = webdriver.Chrome(local.CHROME_DRIVER_PATH, chrome_options=chrome_options)


def quit_driver():
    global driver
    if driver:
        driver.quit()
    driver = None


def switch_to_active_tab():
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
            print("Switched to: " + driver.title.encode('ascii', 'backslashreplace'))
            return


def test_driver():
    switch_to_active_tab()
    driver.get('http://www.google.com/xhtml');


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
        element.click()


class SmartElementAction(DynStrActionBase):

    def __init__(self, by, spec, tracker):
        DynStrActionBase.__init__(self, spec)
        self.by = by
        self.tracker = tracker

    def _parse_spec(self, spec):
        return spec

    def _execute_events(self, events):
        # Get gaze location as early as possible.
        gaze_location = self.tracker.get_gaze_point_or_default()
        switch_to_active_tab()
        elements = driver.find_elements(self.by, events)
        if not elements:
            print("No matching elements found")
            return
        nearest_element = None
        nearest_element_distance_squared = float("inf")
        for element in elements:
            # Assume there is equal amount of browser chrome on the left and right sides of the screen.
            canvas_x_offset = driver.execute_script("return window.screenX + (window.outerWidth - window.innerWidth) / 2 - window.scrollX;")
            # Assume all the browser chrome is on the top of the screen and none on the bottom.
            canvas_y_offset = driver.execute_script("return window.screenY + (window.outerHeight - window.innerHeight) - window.scrollY;")
            element_location = (element.rect["x"] + canvas_x_offset + element.rect["width"] / 2,
                                element.rect["y"] + canvas_y_offset + element.rect["height"] / 2)
            offsets = (element_location[0] - gaze_location[0], element_location[1] - gaze_location[1])
            distance_squared = offsets[0] * offsets[0] + offsets[1] * offsets[1]
            if distance_squared < nearest_element_distance_squared:
                nearest_element = element
                nearest_element_distance_squared = distance_squared
        self._execute_on_element(nearest_element)


class SmartClickElementAction(SmartElementAction):

    def _execute_on_element(self, element):
        element.click()


class DoubleClickElementAction(ElementAction):

    def _execute_on_element(self, element):
        ActionChains(driver).double_click(element).perform()


class ClickElementOffsetAction(ElementAction):

    def __init__(self, by, spec, xoffset, yoffset):
        super(ClickElementOffsetAction, self).__init__(by, spec)
        self.xoffset = xoffset
        self.yoffset = yoffset

    def _execute_on_element(self, element):
        ActionChains(driver).move_to_element_with_offset(element, self.xoffset, self.yoffset).click().perform()
