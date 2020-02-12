# TODO: BUGS: vk.com window
# TODO: youtube.com fullscreen bug below:


from __future__ import print_function
import time
import json
import datetime
import sys
import re
from os import path, mkdir  # system
from activity import *
from dateutil.parser import parser


active_window_name = ""
activity_name = ""
start_time = datetime.datetime.now()
activeList = AcitivyList([])
first_time = True


if sys.platform in ['Windows', 'win32', 'cygwin']:
    import win32gui
    import uiautomation as auto
elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
    from AppKit import NSWorkspace
    from Foundation import *
elif sys.platform in ['linux', 'linux2']:
    import linux as l


def url_to_name(url):
    """
    Retrieves url currently browsing on
    :param url: str
    :return: FQDN
    """
    try:
        string_list = url.split('/')
        return string_list[2]
    except IndexError:
        string_list = "youtube_fullscreen_mode"
        return string_list


def get_active_window():
    """
    Resolves title of opened window
    :return: str, window's name
    """
    _active_window_name = None
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        window = win32gui.GetForegroundWindow()
        _active_window_name = win32gui.GetWindowText(window)
    elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
        _active_window_name = (NSWorkspace.sharedWorkspace()
            .activeApplication()['NSApplicationName'])
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)
    return _active_window_name


def excluded_activities():
    """
    Excludes
    :return: True if exclusion pattern matches, False if not
    """
    window_name = get_active_window()
    exclusions = ['\d+%\s(complete)']
    #
    for exclusion in exclusions:
        if re.match(re.compile(exclusion), window_name):
            return True
    return False


def get_chrome_url():
    """
    Resolves url of the current chrome's active tab
    :return: str, active tab title
    """
    try:
        if sys.platform in ['Windows', 'win32', 'cygwin']:
            window = win32gui.GetForegroundWindow()
            chrome_control = auto.ControlFromHandle(window)
            edit = chrome_control.EditControl()
            return 'https://' + edit.GetValuePattern().Value
        elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
            text_of_my_script = """tell app "google chrome" to get the url of the active tab of window 1"""
            s = NSAppleScript.initWithSource_(
                NSAppleScript.alloc(), text_of_my_script)
            results, err = s.executeAndReturnError_(None)
            return results.stringValue()
        else:
            print("sys.platform={platform} is not supported."
                  .format(platform=sys.platform))
            print(sys.version)
        return _active_window_name
    except (LookupError, NameError):
        _active_window_name = "undefined"
        return _active_window_name
        # 2020-02-09 21:41:17.040 autotimer.py[80] get_chrome_url -> Find Control Timeout: {ControlType: EditControl}
        # Traceback (most recent call last):
        #   File ".\autotimer.py", line 107, in <module>
        #     new_window_name = url_to_name(get_chrome_url())
        #   File ".\autotimer.py", line 80, in get_chrome_url
        #     return 'https://' + edit.GetValuePattern().Value
        #   File "AutoTimer\env\lib\site-packages\uiautomation\uiautomation.py", line 6612, in GetValuePattern
        #     return self.GetPattern(PatternId.ValuePattern)
        #   File "AutoTimer\env\lib\site-packages\uiautomation\uiautomation.py", line 5598, in GetPattern
        #     pattern = self.Element.GetCurrentPattern(patternId)
        #   File "AutoTimer\env\lib\site-packages\uiautomation\uiautomation.py", line 5660, in Element
        #     self.Refind(maxSearchSeconds=TIME_OUT_SECOND, searchIntervalSeconds=self.searchInterval)
        #   File "AutoTimer\env\lib\site-packages\uiautomation\uiautomation.py", line 5908, in Refind
        #     raise LookupError('Find Control Timeout: ' + self.GetSearchPropertiesStr())
        # LookupError: Find Control Timeout: {ControlType: EditControl}
        #
        # pass
        # 2020 - 02 - 09 22: 24:50.020 autotimer.py[81] get_chrome_url -> Find ControlTimeout: {ControlType: EditControl}
        # Traceback (most recent call last):
        #   File ".\autotimer.py", line 111, in <module>
        #     new_window_name = url_to_name(get_chrome_url())
        #   File ".\autotimer.py", line 95, in get_chrome_url
        #     return _active_window_name
        # NameError: name '_active_window_name' is not defined
        #
        # pass
    except Exception as e:
        raise e


try:
    if not path.isdir('activities'):
        mkdir('activities')
    activeList.initialize_me()
except Exception:
    print('No json')

try:
    while True:
        previous_site = ""
        if sys.platform not in ['linux', 'linux2']:
            new_window_name = get_active_window()
            if 'Google Chrome' in new_window_name:
                new_window_name = url_to_name(get_chrome_url())
        if sys.platform in ['linux', 'linux2']:
            new_window_name = l.get_active_window_x()
            if 'Google Chrome' in new_window_name:
                new_window_name = l.get_chrome_url_x()
        #
        if (active_window_name != new_window_name) and not excluded_activities():
            print(active_window_name)
            activity_name = active_window_name
            #
            if not first_time:
                end_time = datetime.datetime.now()
                time_entry = TimeEntry(start_time, end_time, 0, 0, 0, 0)
                time_entry._get_specific_times()
                #
                exists = False
                for activity in activeList.activities:
                    if activity.name == activity_name:
                        exists = True
                        activity.time_entries.append(time_entry)
                #
                if not exists:
                    activity = Activity(activity_name, [time_entry])
                    activeList.activities.append(activity)
                with open('activities/activities.json', 'w') as json_file:
                    json.dump(activeList.serialize(), json_file,
                              indent=4, sort_keys=True)
                    start_time = datetime.datetime.now()
            first_time = False
            active_window_name = new_window_name
        #
        time.sleep(1)

except KeyboardInterrupt:
    with open('activities/activities.json', 'w') as json_file:
        json.dump(activeList.serialize(), json_file, indent=4, sort_keys=True)
