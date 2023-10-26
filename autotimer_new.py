from __future__ import print_function
import time
from os import system
from activity import *
import constants as cnst
import json
import datetime
import sys
if sys.platform in ['Windows', 'win32', 'cygwin']:
    import win32gui
    import uiautomation as auto
elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
    from AppKit import NSWorkspace
    from Foundation import *
elif sys.platform in ['linux', 'linux2']:
        import linux as l

# New libs
from win32process import GetWindowThreadProcessId
from pywinauto.application import Application

active_window_name = ""
activity_name = ""
start_time = datetime.datetime.now()
activeList = AcitivyList([])
first_time = True


def url_to_name(url):
    string_list = url.split('/')
    ret = string_list[2] if string_list[2] != '' else 'other'
    return string_list[2]


def get_active_window():
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

# Dodanie innych przeglądarek
def get_edge_url():
    """
    For chrome run in app mode.
    """
    if sys.platform in ['Windows', 'win32', 'cygwin']:
    #time.sleep(3)
        window = win32gui.GetForegroundWindow()
        tid, pid = GetWindowThreadProcessId(window)
        #app = Application(backend="uia").connect(title_re=".*Edge*", found_index=0)
        app = Application(backend="uia").connect(process=pid, time_out=10)
        dlg = app.top_window()
        title = "App bar"
        #url = dlg.child_window(title=title, control_type="Edit").get_value()
        wrapper = dlg.child_window(title=title, control_type="ToolBar")
        url = wrapper.descendants(control_type='Edit')[0].get_value()
        return url
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)

def get_chrome_url():
    """
    For chrome run in app mode.
    """
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        #time.sleep(3)
        window = win32gui.GetForegroundWindow()
        tid, pid = GetWindowThreadProcessId(window)
        app = Application(backend="uia").connect(process=pid, time_out=10)
        dlg = app.top_window()
        title = "Address and search bar"
        url = dlg.child_window(title=title, control_type="Edit").get_value()
        return url
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)

def get_brave_url():
    if sys.platform in ['Windows', 'win32', 'cygwin']:
        time.sleep(3)
        window = win32gui.GetForegroundWindow()
        chromeControl = auto.ControlFromHandle(window)
        edit = chromeControl.EditControl()
        ret = edit.GetValuePattern().Value
        ret = 'https://' + ret if 'https' not in ret else ret
        return ret
    elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
        textOfMyScript = """tell app "google chrome" to get the url of the active tab of window 1"""
        s = NSAppleScript.initWithSource_(
            NSAppleScript.alloc(), textOfMyScript)
        results, err = s.executeAndReturnError_(None)
        return results.stringValue()
    else:
        print("sys.platform={platform} is not supported."
              .format(platform=sys.platform))
        print(sys.version)
#    return _active_window_name

try:
    activeList.initialize_me()
except Exception:
    print('No json')


try:
    while True:
        previous_site = ""
        if sys.platform not in ['linux', 'linux2']:
            new_window_name = get_active_window()
            if 'google chrome' in new_window_name.lower():
                new_window_name = url_to_name(get_chrome_url())
            elif 'microsoft' in new_window_name.lower() and 'edge' in new_window_name.lower():
                new_window_name = url_to_name(get_edge_url())
            elif 'brave' in new_window_name.lower():
                new_window_name = url_to_name(get_brave_url())
            elif 'opera' in new_window_name.lower():
                new_window_name = 'Opera' # TODO
        if sys.platform in ['linux', 'linux2']:
            new_window_name = l.get_active_window_x()
            if 'Google Chrome' in new_window_name:
                new_window_name = l.get_chrome_url_x()

        
        if active_window_name != new_window_name:
            print(active_window_name)
            activity_name = active_window_name

            if not first_time:
                end_time = datetime.datetime.now()
                time_entry = TimeEntry(start_time, end_time, 0, 0, 0, 0)
                time_entry._get_specific_times()

                exists = False
                for activity in activeList.activities:
                    if activity.name == activity_name:
                        exists = True
                        activity.time_entries.append(time_entry)

                if not exists:
                    activity = Activity(activity_name, [time_entry])
                    activeList.activities.append(activity)
                with open(cnst.store_activities_path + 'activities.json', 'w') as json_file:
                    json.dump(activeList.serialize(), json_file,
                              indent=4, sort_keys=True)
                    start_time = datetime.datetime.now()
            first_time = False
            active_window_name = new_window_name

        time.sleep(1)
    
except KeyboardInterrupt:
    with open(cnst.store_activities_path + 'activities.json', 'w') as json_file:
        json.dump(activeList.serialize(), json_file, indent=4, sort_keys=True)
