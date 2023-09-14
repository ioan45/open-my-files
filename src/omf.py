import os
import sys
import winreg
import subprocess
import threading
import json
import time
import PySimpleGUI as sg
from watchdog.observers import Observer
from watchdog.events import PatternMatchingEventHandler
from copy import deepcopy

ENTRY_EXE_FILE = 'executable'
ENTRY_OTHER_FILE = 'other'
ENTRY_WEB_PAGE = 'web_page'

ICON_EXE_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAItJREFUOE/tksENg0AMBGfrCKUkBSAKQIg/PYU/6SAdIBqBPhYZiSiPiOguPOOPJd/teGVb/BiyXQF34JLIWoAuADNQS5pSALavwCMAlhT5GQXAQCspnB3Gpt0B3z5/en8BcsS75hwHMYMcF+fN4O9gO+VG0piyCds3YIg7KIEeKFIAQDTusvb/3mgFDG1uMt84gY0AAAAASUVORK5CYII='
ICON_OTHER_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAL9JREFUOE/lkz0KAjEQRt93DkWw0M4jWFiKBxBFsPMaeg1txTt4APEQCiJ6kE8C2SUbVt3GyjSTTCZvfjIj4rI9AbZAq9Al8g6MJAVZWUoAT2Aq6Zwb2d4AyzpICrCk8pxCImANPCLkVtw3BSyAfXxUSacpINjNgUGEhFS7Yd8IUFOTMt3fAGxfgF7m+SqpH3S2/ymCMu+ssX5cg5phqqje/UIYppmk0yeA7SFwkNTJO3EM7ID2lwiCo5WkY7B7ATUxkhEHIy87AAAAAElFTkSuQmCC'
ICON_WEB_PAGE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAa5JREFUOE+d08uLDmAUx/HP2bgtJn8CWQkhFhYyyiUsxUrKLWwkk5KNZmyk3BZqmolmZSNrNOOWKGUi5bIS/gEiuS6Oztvz6u0tNTmb09Nznu85v/OcE/osM5fhADZgQbt+jzu4EhEve59E95CZs3ER2zEXT/Gh3RdoNb7jBoYi4lfddQDt8S18xDWcx4oKbIACP8dx7MJ8bC1IFzCK3XiBeS3LBZxogLMNthNfG3wiIo5E03wfi7EZV/EWC7sVVpF4h0XYhym8wWABLlXpEXE6M6txpyJisEk709EZcbKdH2IkIu5m5jAGCvAKByPicWYewsqIONweVFABOj4zx/AsIsYycy1GC/AFpbc6vB5zcLtp39h8fWHZFvzAg/ZTQ13ARAOsaYAKKKssZY+aL2k/8aQB9vZLqNJLQkmpkvsljGM6IsZ7JVQTP0XEyAybOBwR93qbuLRp+p9vXNcdpMvYM4NB2oFvWF7zEhFHu4BZuInPbZTPVS9w7B+jPIBtEfG7d5kKUjtQWWqcp1FbWFZTuaplv147UY//LlML6rjMXIL92NS3zpOt7Ne98X8AKibG9Jc6ezQAAAAASUVORK5CYII='

KEY_GROUP_ID = 'group_id'
KEY_GROUP_NAME = 'group_name'
KEY_GROUP_LISTENING_DIR = 'group_listening_dir'
KEY_GROUP_ENTRIES = 'group_entries'
KEY_ENTRY_ID = 'entry_id'
KEY_ENTRY_PATH = 'entry_path'
KEY_ENTRY_TYPE = 'entry_type'
KEY_ENTRY_DETAILS = 'entry_details'
KEY_SETTING_START_WITH_WINDOWS = 'start_with_windows'
KEY_SETTING_AUTO_SAVE = 'auto_save'

KEY_LAYOUT_MAIN = '-MAIN_LAYOUT-'
KEY_LAYOUT_GROUP_EDIT = '-GROUP_EDIT_LAYOUT-'
KEY_LAYOUT_HELP = '-HELP_LAYOUT-'
KEY_TREE_MAIN = '-MAIN_TREE-'
KEY_TREE_GROUP_EDIT = '-GROUP_EDIT_TREE-'
KEY_BUTTON_OPEN_GROUP = '-OPEN_GROUP_BUTTON-'
KEY_BUTTON_NEW_GROUP = '-NEW_GROUP_BUTTON-'
KEY_BUTTON_EDIT_GROUP = '-EDIT_GROUP_BUTTON-'
KEY_BUTTON_DELETE_GROUP = '-DELETE_GROUP_BUTTON-'
KEY_BUTTON_SAVE_CHANGES = '-SAVE_CHANGES_BUTTON-'
KEY_BUTTON_REVERT_CHANGES = '-REVERT_CHANGES_BUTTON-'
KEY_BUTTON_HELP = '-HELP_BUTTON-'
KEY_BUTTON_EXIT = '-EXIT_BUTTON-'
KEY_BUTTON_ADD_FILES = '-ADD_FILES_BUTTON-'
KEY_BUTTON_ADD_WEB_PAGE = '-ADD_WEB_PAGE_BUTTON-'
KEY_BUTTON_DELETE_ENTRIES = '-DELETE_ENTRIES_BUTTON-'
KEY_BUTTON_EDIT_DETAILS = '-EDIT_DETAILS_BUTTON-'
KEY_BUTTON_BACK_GROUP_EDIT = '-GROUP_EDIT_BACK_BUTTON-'
KEY_BUTTON_BACK_HELP = '-HELP_BACK_BUTTON-'
KEY_BUTTON_LISTEN_TO_DIR = 'LISTEN_DIR_BUTTON'
KEY_STATUS_BAR_GROUPS = '-GROUPS_STATUS_BAR-'
KEY_STATUS_BAR_ENTRIES = '-ENTRIES_STATUS_BAR-'
KEY_LABEL_GROUP_NAME = '-GROUP_NAME_LABEL-'
KEY_HYPERLINK_HELP = '-HELP_HYPERLINK-'
KEY_CHECKBOX_START_WITH_WINDOWS = '-START_WITH_WINDOWS-'
KEY_CHECKBOX_AUTO_SAVE = '-AUTO_SAVE-'

EVENT_OPENING_COMPLETE = '-OPENING_COMPLETE-'

PATH_GROUPS_DATA = os.path.join(os.getenv('LOCALAPPDATA'), 'OpenMyFiles', 'groups.json')
PATH_SETTINGS_DATA = os.path.join(os.getenv('LOCALAPPDATA'), 'OpenMyFiles', 'settings.json')

def get_default_browser_path():
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.html\\UserChoice") as id_key:
        browser_id = winreg.QueryValueEx(id_key, 'ProgId')[0]
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{browser_id}\\shell\\open\\command") as path_key:
        browser_path = winreg.QueryValueEx(path_key, '')[0]
    return browser_path.split('"')[1]

DEFAULT_BROWSER_PATH = get_default_browser_path()

class AppData:  
    def __init__(self, groups_path, settings_path) -> None:
        self.groups_path = groups_path
        self.__saved_groups_data = []
        self.groups_data = []
        self.settings_path = settings_path
        self.__saved_settings = dict()
        self.settings = dict()
    
    def load_groups_data(self):
        if os.path.isfile(self.groups_path):
            with open(self.groups_path, 'r') as f:
                self.__saved_groups_data = json.load(f)
            self.groups_data = deepcopy(self.__saved_groups_data)

    def load_settings(self):
        if os.path.isfile(self.settings_path):
            with open(self.settings_path, 'r') as f:
                self.__saved_settings = json.load(f)
            self.settings = deepcopy(self.__saved_settings)

    def save_group_data(self):
        self.__saved_groups_data = deepcopy(self.groups_data)
        os.makedirs(os.path.split(self.groups_path)[0], exist_ok=True)
        with open(self.groups_path, 'w') as f:
            json.dump(self.__saved_groups_data, f, indent='\t')
    
    def save_settings(self):
        self.__saved_settings = deepcopy(self.settings)
        os.makedirs(os.path.split(self.settings_path)[0], exist_ok=True)
        with open(self.settings_path, 'w') as f:
            json.dump(self.__saved_settings, f, indent='\t')

    def revert_changes(self):
        self.groups_data = deepcopy(self.__saved_groups_data)
        self.settings = deepcopy(self.__saved_settings)
    
    def refresh_ids(self, for_groups=False, for_group_with_id=None):
        if for_groups:
            for index, group in enumerate(self.groups_data):
                group[KEY_GROUP_ID] = index
        if for_group_with_id is not None:
            for index, entry in enumerate(self.groups_data[for_group_with_id][KEY_GROUP_ENTRIES]):
                entry[KEY_ENTRY_ID] = index

def listening_on_created(event):
    file_path = event.src_path.replace(os.sep, '/')
    if file_path.split('.')[-1] == 'exe':
        entry_type = ENTRY_EXE_FILE
    else:
        entry_type = ENTRY_OTHER_FILE
    path_to_listen = os.path.split(file_path)[0]
    
    for group in app_data.groups_data:
        if KEY_GROUP_LISTENING_DIR in group and group[KEY_GROUP_LISTENING_DIR] == path_to_listen:
            group[KEY_GROUP_ENTRIES].append({
                KEY_ENTRY_ID: len(group[KEY_GROUP_ENTRIES]),
                KEY_ENTRY_PATH: file_path, 
                KEY_ENTRY_TYPE: entry_type, 
                KEY_ENTRY_DETAILS: ''
            })
            if group_to_edit is None:
                window[KEY_TREE_MAIN].update(key=group[KEY_GROUP_ID], value=[len(group[KEY_GROUP_ENTRIES])])
    
    if group_to_edit is not None and KEY_GROUP_LISTENING_DIR in group_to_edit and group_to_edit[KEY_GROUP_LISTENING_DIR] == path_to_listen:
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
    window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
    window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    window.refresh()

def listening_on_deleted(event):
    file_path = event.src_path.replace(os.sep, '/')
    path_to_listen = os.path.split(file_path)[0]

    for group in app_data.groups_data:
        if KEY_GROUP_LISTENING_DIR in group and group[KEY_GROUP_LISTENING_DIR] == path_to_listen:
            group[KEY_GROUP_ENTRIES][:] = [entry for entry in group[KEY_GROUP_ENTRIES] if entry[KEY_ENTRY_PATH] != file_path]
            app_data.refresh_ids(for_group_with_id=group[KEY_GROUP_ID])
            if group_to_edit is None:
                window[KEY_TREE_MAIN].update(key=group[KEY_GROUP_ID], value=[len(group[KEY_GROUP_ENTRIES])])
    
    if group_to_edit is not None and KEY_GROUP_LISTENING_DIR in group_to_edit and group_to_edit[KEY_GROUP_LISTENING_DIR] == path_to_listen:
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
    window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
    window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    window.refresh()

def listening_on_moved(event):
    old_file_path = event.src_path.replace(os.sep, '/')
    new_file_path = event.dest_path.replace(os.sep, '/')
    path_to_listen = os.path.split(old_file_path)[0]

    for group in app_data.groups_data:
        if KEY_GROUP_LISTENING_DIR in group and group[KEY_GROUP_LISTENING_DIR] == path_to_listen:
            for entry in group[KEY_GROUP_ENTRIES]:
                if entry[KEY_ENTRY_PATH] == old_file_path:
                    entry[KEY_ENTRY_PATH] = new_file_path
    
    if group_to_edit is not None and KEY_GROUP_LISTENING_DIR in group_to_edit and group_to_edit[KEY_GROUP_LISTENING_DIR] == path_to_listen:
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
    window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
    window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    window.refresh()

def create_entries_tree_data():
    new_tree_data = sg.TreeData()
    for entry in group_to_edit[KEY_GROUP_ENTRIES]:
        icon = ICON_OTHER_FILE
        if entry[KEY_ENTRY_TYPE] == ENTRY_EXE_FILE:
            icon = ICON_EXE_FILE
        elif entry[KEY_ENTRY_TYPE] == ENTRY_WEB_PAGE:
            icon = ICON_WEB_PAGE
        new_tree_data.insert('', entry[KEY_ENTRY_ID], f'  {entry[KEY_ENTRY_PATH]}', [entry[KEY_ENTRY_DETAILS]], icon)
    return new_tree_data

def create_groups_tree_data():
    new_tree_data = sg.TreeData()
    for group in app_data.groups_data:
        new_tree_data.insert('', group[KEY_GROUP_ID], f'  {group[KEY_GROUP_NAME]}', [len(group[KEY_GROUP_ENTRIES])], ICON_OTHER_FILE)
    return new_tree_data

def change_active_layout(to_layout_key):
    global active_layout_key
    window[active_layout_key].update(visible=False)
    window[to_layout_key].update(visible=True)
    active_layout_key = to_layout_key
    window.disappear()
    window.refresh()
    window.move_to_center()
    window.reappear()

def update_groups_status_bar(message):
    with groups_status_bar_lock:
        window[KEY_STATUS_BAR_GROUPS].update(value=message)

def save_changes():
    with check_if_saving_lock:
        if is_saving.is_set():
            return
        else:
            is_saving.set()

    if threading.current_thread() is threading.main_thread():
        update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Saving changes...")
    else:
        update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Auto save: Saving changes...")

    window[KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
    window[KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
    app_data.save_group_data()
    app_data.save_settings()
    set_start_with_windows(app_data.settings[KEY_SETTING_START_WITH_WINDOWS])

    if threading.current_thread() is threading.main_thread():
        update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been successfully saved!")
    else:
        update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Auto save: Changes have been successfully saved!")

    is_saving.clear()

def set_start_with_windows(enabled):
    app_key = 'OpenMyFiles'
    path_to_exe = None
    if enabled:
        # For using this application as an executable made with PyInstaller:
        # The PyInstaller way of finding out if this script is running or not as an executable made by PyInstaller.
        if getattr(sys, 'frozen', False):
            path_to_exe = f'"{sys.executable}"'
        else:
            path_to_exe = f'"{sys.executable}" "{os.path.realpath(__file__)}"'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run", access=winreg.KEY_SET_VALUE) as reg_key:
        if path_to_exe is None:
            try:
                winreg.DeleteValue(reg_key, app_key)
            except FileNotFoundError:
                pass
        else:
            winreg.SetValueEx(reg_key, app_key, 0, winreg.REG_SZ, path_to_exe)

def async_auto_saving():
    auto_save_delay_s = 60
    while True:
        time.sleep(auto_save_delay_s)
        if auto_save_enabled.is_set() and not window[KEY_BUTTON_SAVE_CHANGES].Disabled:
            save_changes()

def async_group_opening(group_name, entries_to_open):
    update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Opening \"{group_name}\" group...")
    tabs_opened = 0
    for entry in entries_to_open:
        if entry[KEY_ENTRY_TYPE] == ENTRY_WEB_PAGE:
            # For some browsers (e.g. Firefox), if they are closed when group opening starts, it may be required to wait a bit for 
            # them to initialize on opening the first tab before opening others.
            if tabs_opened == 1:
                time.sleep(1.5)
            # Using this approach because os.startfile(URL) will not work if the URL specified by user doesn't contain the protocol.
            # Also, in this case, webbrowser.open(URL) will use Microsoft Edge regardless of the default browser.
            subprocess.Popen([DEFAULT_BROWSER_PATH, entry[KEY_ENTRY_PATH]], creationflags=subprocess.DETACHED_PROCESS)
            tabs_opened += 1
        elif os.path.isfile(entry[KEY_ENTRY_PATH]):
            os.startfile(entry[KEY_ENTRY_PATH])
    update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Group \"{group_name}\" opened.")

sg.theme('DarkAmber')

app_data = AppData(PATH_GROUPS_DATA, PATH_SETTINGS_DATA)
app_data.load_groups_data()
app_data.load_settings()

active_layout_key = KEY_LAYOUT_MAIN
group_to_edit = None
groups_status_bar_lock = threading.Lock()

check_if_saving_lock = threading.Lock()
is_saving = threading.Event()
auto_save_enabled = threading.Event()
if KEY_SETTING_AUTO_SAVE in app_data.settings and app_data.settings[KEY_SETTING_AUTO_SAVE]:
    auto_save_enabled.set()
auto_save_thread = threading.Thread(target=async_auto_saving, daemon=True)
auto_save_thread.start()

listening_event_handler = PatternMatchingEventHandler('*', None, True, True)
listening_event_handler.on_created = listening_on_created
listening_event_handler.on_deleted = listening_on_deleted
listening_event_handler.on_moved = listening_on_moved
listening_observer = Observer()
listening_watches = []
for group in app_data.groups_data:
    if KEY_GROUP_LISTENING_DIR in group:
        if os.path.isdir(group[KEY_GROUP_LISTENING_DIR]):
            listening_watches.append(listening_observer.schedule(listening_event_handler, group[KEY_GROUP_LISTENING_DIR]))
        else:
            group.pop(KEY_GROUP_LISTENING_DIR)
listening_observer.start()

layout_main = [[
        sg.Tree(create_groups_tree_data(), 
                key=KEY_TREE_MAIN, 
                headings=['Entries'],
                header_font=sg.DEFAULT_FONT, 
                select_mode=sg.TABLE_SELECT_MODE_BROWSE, 
                num_rows=20,
                auto_size_columns=False,
                col_widths=[7],
                col0_width=35,
                col0_heading='Group',
                justification='center',
                expand_x=True,
                expand_y=True,
                font=('Verdana', 10, 'normal'),
                pad=((15, 15), (15, 15)),
                enable_events=True),
        sg.Column([
                [sg.Button('Open Group', key=KEY_BUTTON_OPEN_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                [sg.Button('New Group', key=KEY_BUTTON_NEW_GROUP, expand_x=True, pad=((0, 0), (7, 0)))],
                [sg.Button('Edit Group', key=KEY_BUTTON_EDIT_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                [sg.Button('Delete Group', key=KEY_BUTTON_DELETE_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                [sg.Button('Save Changes', key=KEY_BUTTON_SAVE_CHANGES, expand_x=True, pad=((0, 0), (35, 0)), disabled=True)],
                [sg.Button('Revert Changes', key=KEY_BUTTON_REVERT_CHANGES, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                [sg.Button('Help', key=KEY_BUTTON_HELP, expand_x=True, pad=((0, 0), (35, 0)))],
                [sg.Button('Exit', key=KEY_BUTTON_EXIT, expand_x=True, pad=((0, 0), (7, 7)))],
                ],
                vertical_alignment='top',
                pad=((15, 15), (15, 15)))
    ],[
        sg.StatusBar('', 
                key=KEY_STATUS_BAR_GROUPS, 
                size=50, 
                pad=((15, 15), (10, 15)), 
                auto_size_text=True,
                font=('Verdana', 10, 'normal'))
    ],[
        sg.Checkbox('Launch the application when Windows starts', 
                key=KEY_CHECKBOX_START_WITH_WINDOWS,
                default=app_data.settings[KEY_SETTING_START_WITH_WINDOWS] if KEY_SETTING_START_WITH_WINDOWS in app_data.settings else False,
                enable_events=True),
        sg.Checkbox('Auto save changes (once per minute)', 
                key=KEY_CHECKBOX_AUTO_SAVE,
                default=app_data.settings[KEY_SETTING_AUTO_SAVE] if KEY_SETTING_AUTO_SAVE in app_data.settings else False,
                enable_events=True)
    ]
]
layout_group_edit = [[
        sg.Text('<Name> Group', key=KEY_LABEL_GROUP_NAME, font=('Arial', 16, 'normal'), pad=15)
    ],[
        sg.Tree(sg.TreeData(), 
                key=KEY_TREE_GROUP_EDIT, 
                headings=['Details'],
                header_font=sg.DEFAULT_FONT, 
                select_mode=sg.TABLE_SELECT_MODE_EXTENDED, 
                num_rows=20,
                auto_size_columns=False,
                col_widths=[35],
                col0_width=55,
                col0_heading='Path',
                justification='center',
                expand_x=True,
                expand_y=True,
                font=('Verdana', 10, 'normal'),
                pad=((15, 15), (0, 15)),
                enable_events=True)
    ],[
        sg.Stretch(),
        sg.Column([[
                sg.Input(key=KEY_BUTTON_ADD_FILES, visible=False, enable_events=True), 
                sg.FilesBrowse('Add Files', KEY_BUTTON_ADD_FILES),
                sg.Button('Add Web Page', key=KEY_BUTTON_ADD_WEB_PAGE),
                sg.Button('Delete Selected Entries', key=KEY_BUTTON_DELETE_ENTRIES, disabled=True),
                sg.Button('Edit Selected Details', key=KEY_BUTTON_EDIT_DETAILS, disabled=True),
                sg.Button('Start Listening', key=KEY_BUTTON_LISTEN_TO_DIR),
                sg.Button('Back', key=KEY_BUTTON_BACK_GROUP_EDIT),
                ]],
                pad=((15, 15), (0, 7))),
        sg.Stretch()
    ],[
        sg.StatusBar('', 
                key=KEY_STATUS_BAR_ENTRIES, 
                size=80, 
                pad=((15, 15), (10, 15)),
                auto_size_text=True,
                font=('Verdana', 10, 'normal'))
    ]
]
layout_help = [[
        sg.VStretch(),
        sg.Column([
                [sg.Text('Check out the documentation in the following GitHub repository:', font='Verdana 10 normal')],
                [sg.Text('https://github.com/ioan45/open-my-files', key=KEY_HYPERLINK_HELP, enable_events=True, font='Verdana 10 underline', text_color='LightBlue')],
                [sg.Button('Back', key=KEY_BUTTON_BACK_HELP, pad=((0, 0), (20, 10)))]
                ],
                element_justification='center',
                justification='center'),
        sg.VStretch()
    ]
]
layout = [[sg.Column(layout_main, key=KEY_LAYOUT_MAIN, expand_x=True, expand_y=True), 
           sg.Column(layout_group_edit, key=KEY_LAYOUT_GROUP_EDIT, visible=False, expand_x=True, expand_y=True),
           sg.Column(layout_help, key=KEY_LAYOUT_HELP, visible=False, expand_x=True, expand_y=True)]]
window = sg.Window('Open My Files', layout, resizable=True, enable_close_attempted_event=True)

while True:
    event, values = window.read()

    if event in [sg.WIN_CLOSE_ATTEMPTED_EVENT, KEY_BUTTON_EXIT]:
        if window[KEY_BUTTON_SAVE_CHANGES].Disabled:
            break
        else:
            response = sg.popup_yes_no('\n== Save changes? ==\n')
            if response is not None:
                if response == 'Yes':
                    save_changes()
                break

    elif event == KEY_TREE_MAIN:
        should_disable = False if values[KEY_TREE_MAIN] else True
        window[KEY_BUTTON_OPEN_GROUP].update(disabled=should_disable)
        window[KEY_BUTTON_EDIT_GROUP].update(disabled=should_disable)
        window[KEY_BUTTON_DELETE_GROUP].update(disabled=should_disable)

    elif event == KEY_BUTTON_OPEN_GROUP:
        selected_group_id = values[KEY_TREE_MAIN][0]
        group_name = app_data.groups_data[selected_group_id][KEY_GROUP_NAME]
        entries_to_open = deepcopy(app_data.groups_data[selected_group_id][KEY_GROUP_ENTRIES])
        threading.Thread(target=async_group_opening, args=(group_name, entries_to_open)).start()

    elif event == KEY_BUTTON_NEW_GROUP:
        group_name = sg.popup_get_text('How should the group be named?', 'Input a name for the new group')
        if group_name is not None and group_name != '':
            app_data.groups_data.append({
                KEY_GROUP_ID: len(app_data.groups_data),
                KEY_GROUP_NAME: group_name, 
                KEY_GROUP_ENTRIES: []
            })
            window[KEY_TREE_MAIN].update(create_groups_tree_data())
            window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_BUTTON_EDIT_GROUP:
        selected_group_id = values[KEY_TREE_MAIN][0]
        group_to_edit = app_data.groups_data[selected_group_id]
        window[KEY_LABEL_GROUP_NAME].update(f'{group_to_edit[KEY_GROUP_NAME]} Group')
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
        if KEY_GROUP_LISTENING_DIR in group_to_edit:
            window[KEY_BUTTON_LISTEN_TO_DIR].update('Stop Listening')
            window[KEY_STATUS_BAR_ENTRIES].update(value=f"Listening on {group_to_edit[KEY_GROUP_LISTENING_DIR]}")
        else:
            window[KEY_BUTTON_LISTEN_TO_DIR].update('Start Listening')
            window[KEY_STATUS_BAR_ENTRIES].update(value='')
        change_active_layout(KEY_LAYOUT_GROUP_EDIT)
    
    elif event == KEY_BUTTON_DELETE_GROUP:
        tmp = []
        for group in app_data.groups_data:
            if group[KEY_GROUP_ID] not in values[KEY_TREE_MAIN]:
                tmp.append(group)
            elif KEY_GROUP_LISTENING_DIR in group:
                path_to_listen = group[KEY_GROUP_LISTENING_DIR]
                watch = [watch for watch in listening_watches if watch.path == path_to_listen][0]
                listening_observer.unschedule(watch)
                listening_watches.remove(watch)
        app_data.groups_data = tmp
        app_data.refresh_ids(for_groups=True)
        window[KEY_TREE_MAIN].update(create_groups_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_BUTTON_SAVE_CHANGES:
        save_changes()
    
    elif event == KEY_BUTTON_REVERT_CHANGES:
        app_data.revert_changes()
        window[KEY_CHECKBOX_START_WITH_WINDOWS].update(app_data.settings[KEY_SETTING_START_WITH_WINDOWS])
        window[KEY_TREE_MAIN].update(create_groups_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
        update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been reverted!")

    elif event == KEY_BUTTON_HELP:
        change_active_layout(KEY_LAYOUT_HELP)
    
    elif event == KEY_HYPERLINK_HELP:
        os.startfile('https://github.com/ioan45/open-my-files')
    
    elif event == KEY_CHECKBOX_START_WITH_WINDOWS:
        app_data.settings[KEY_SETTING_START_WITH_WINDOWS] = values[KEY_CHECKBOX_START_WITH_WINDOWS]
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_CHECKBOX_AUTO_SAVE:
        checked = values[KEY_CHECKBOX_AUTO_SAVE]
        if checked:
            auto_save_enabled.set()
        else:
            auto_save_enabled.clear()
        app_data.settings[KEY_SETTING_AUTO_SAVE] = checked
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_BUTTON_BACK_HELP:
        change_active_layout(KEY_LAYOUT_MAIN)

    elif event == KEY_TREE_GROUP_EDIT:
        should_disable = False if values[KEY_TREE_GROUP_EDIT] else True
        window[KEY_BUTTON_DELETE_ENTRIES].update(disabled=should_disable)
        window[KEY_BUTTON_EDIT_DETAILS].update(disabled=should_disable)
    
    elif event == KEY_BUTTON_ADD_FILES:
        new_entries = values[KEY_BUTTON_ADD_FILES].split(';')
        for index, entry in enumerate(new_entries, start=len(group_to_edit[KEY_GROUP_ENTRIES])):
            if entry.split('.')[-1] == 'exe':
                group_to_edit[KEY_GROUP_ENTRIES].append({
                    KEY_ENTRY_ID: index,
                    KEY_ENTRY_PATH: entry, 
                    KEY_ENTRY_TYPE: ENTRY_EXE_FILE, 
                    KEY_ENTRY_DETAILS: ''
                })
            else:
                group_to_edit[KEY_GROUP_ENTRIES].append({
                    KEY_ENTRY_ID: index,
                    KEY_ENTRY_PATH: entry, 
                    KEY_ENTRY_TYPE: ENTRY_OTHER_FILE, 
                    KEY_ENTRY_DETAILS: ''
                })
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    
    elif event == KEY_BUTTON_ADD_WEB_PAGE:
        added_url = sg.popup_get_text("What's the web address of your desired web page?", "Input URL")
        if added_url is not None and added_url != '':
            group_to_edit[KEY_GROUP_ENTRIES].append({
                KEY_ENTRY_ID: len(group_to_edit[KEY_GROUP_ENTRIES]),
                KEY_ENTRY_PATH: added_url,
                KEY_ENTRY_TYPE: ENTRY_WEB_PAGE, 
                KEY_ENTRY_DETAILS: ''
            })
            window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
            window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    
    elif event == KEY_BUTTON_DELETE_ENTRIES:
        group_to_edit[KEY_GROUP_ENTRIES][:] = [entry for entry in group_to_edit[KEY_GROUP_ENTRIES] if entry[KEY_ENTRY_ID] not in values[KEY_TREE_GROUP_EDIT]]
        app_data.refresh_ids(for_group_with_id=group_to_edit[KEY_GROUP_ID])
        window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    
    elif event == KEY_BUTTON_EDIT_DETAILS:
        selected_entries = [entry for entry in group_to_edit[KEY_GROUP_ENTRIES] if entry[KEY_ENTRY_ID] in values[KEY_TREE_GROUP_EDIT]]
        old_details = selected_entries[0][KEY_ENTRY_DETAILS] if len(selected_entries) == 1 else ''
        new_details = sg.popup_get_text('What are the details of those entries?', 'Input details for selected entries', old_details)
        if new_details is not None:
            for entry in selected_entries:
                entry[KEY_ENTRY_DETAILS] = new_details
            window[KEY_TREE_GROUP_EDIT].update(create_entries_tree_data())
            window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_BUTTON_LISTEN_TO_DIR:
        if window[KEY_BUTTON_LISTEN_TO_DIR].get_text().startswith('Start'):
            path_to_listen = sg.popup_get_folder('Select the directory to listen:', title='Select Directory')
            if path_to_listen is not None and path_to_listen != '':
                listening_watches.append(listening_observer.schedule(listening_event_handler, path_to_listen))
                group_to_edit[KEY_GROUP_LISTENING_DIR] = path_to_listen
                window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
                window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
                window[KEY_BUTTON_LISTEN_TO_DIR].update('Stop Listening')
                window[KEY_STATUS_BAR_ENTRIES].update(value=f"Listening on {path_to_listen}")
        else:
            path_to_listen = group_to_edit[KEY_GROUP_LISTENING_DIR]
            watch = [watch for watch in listening_watches if watch.path == path_to_listen][0]
            listening_observer.unschedule(watch)
            listening_watches.remove(watch)
            group_to_edit.pop(KEY_GROUP_LISTENING_DIR)
            window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
            window[KEY_BUTTON_LISTEN_TO_DIR].update('Start Listening')
            window[KEY_STATUS_BAR_ENTRIES].update(value='')


    elif event == KEY_BUTTON_BACK_GROUP_EDIT:
        window[KEY_TREE_MAIN].update(key=group_to_edit[KEY_GROUP_ID], value=[len(group_to_edit[KEY_GROUP_ENTRIES])])
        group_to_edit = None
        change_active_layout(KEY_LAYOUT_MAIN)

window.close()
listening_observer.stop()
listening_observer.join()
