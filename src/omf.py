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
from copy import copy, deepcopy


class AppData:
    """
    Class for managing the application data.

    Groups data format:
        The groups data consists of a list of groups. Each group is a key-value type of structure which contains the group id, name, 
        entries and the path to the directory to listen. The entries value in the group structure is a list of key-value type of
        structure elements, each one representing a group entry. Each group entry structure contains the entry id, path/address,
        type (executable/web address/other file) and details.

        JSON Format:
        ```
        [
            {
                "group_id": <type_int>,
                "group_name": <type_str>,
                "group_entries": [
                    {
                        "entry_id": <type_int>,
                        "entry_path": <type_str>,
                        "entry_type": <type_str>,
                        "entry_details": <type_str>
                    },
                    ...
                ],
                "group_listening_dir": <type_str>
            },
            ...
        ]
        ```

    Settings data format:
        Settings are stored simply in a key-value type of structure where each entry in the structure consists of the setting name (key)
        and the setting value.

        JSON Format:
        ```
        {
            <setting_name>: <setting_value>,
            ...
        }
        ```
    """

    ICON_EXE_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAItJREFUOE/tksENg0AMBGfrCKUkBSAKQIg/PYU/6SAdIBqBPhYZiSiPiOguPOOPJd/teGVb/BiyXQF34JLIWoAuADNQS5pSALavwCMAlhT5GQXAQCspnB3Gpt0B3z5/en8BcsS75hwHMYMcF+fN4O9gO+VG0piyCds3YIg7KIEeKFIAQDTusvb/3mgFDG1uMt84gY0AAAAASUVORK5CYII='
    ICON_OTHER_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAL9JREFUOE/lkz0KAjEQRt93DkWw0M4jWFiKBxBFsPMaeg1txTt4APEQCiJ6kE8C2SUbVt3GyjSTTCZvfjIj4rI9AbZAq9Al8g6MJAVZWUoAT2Aq6Zwb2d4AyzpICrCk8pxCImANPCLkVtw3BSyAfXxUSacpINjNgUGEhFS7Yd8IUFOTMt3fAGxfgF7m+SqpH3S2/ymCMu+ssX5cg5phqqje/UIYppmk0yeA7SFwkNTJO3EM7ID2lwiCo5WkY7B7ATUxkhEHIy87AAAAAElFTkSuQmCC'
    ICON_WEB_PAGE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAa5JREFUOE+d08uLDmAUx/HP2bgtJn8CWQkhFhYyyiUsxUrKLWwkk5KNZmyk3BZqmolmZSNrNOOWKGUi5bIS/gEiuS6Oztvz6u0tNTmb09Nznu85v/OcE/osM5fhADZgQbt+jzu4EhEve59E95CZs3ER2zEXT/Gh3RdoNb7jBoYi4lfddQDt8S18xDWcx4oKbIACP8dx7MJ8bC1IFzCK3XiBeS3LBZxogLMNthNfG3wiIo5E03wfi7EZV/EWC7sVVpF4h0XYhym8wWABLlXpEXE6M6txpyJisEk709EZcbKdH2IkIu5m5jAGCvAKByPicWYewsqIONweVFABOj4zx/AsIsYycy1GC/AFpbc6vB5zcLtp39h8fWHZFvzAg/ZTQ13ARAOsaYAKKKssZY+aL2k/8aQB9vZLqNJLQkmpkvsljGM6IsZ7JVQTP0XEyAybOBwR93qbuLRp+p9vXNcdpMvYM4NB2oFvWF7zEhFHu4BZuInPbZTPVS9w7B+jPIBtEfG7d5kKUjtQWWqcp1FbWFZTuaplv147UY//LlML6rjMXIL92NS3zpOt7Ne98X8AKibG9Jc6ezQAAAAASUVORK5CYII='

    ENTRY_EXE_FILE = 'executable'
    ENTRY_OTHER_FILE = 'other'
    ENTRY_WEB_PAGE = 'web_page'

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
                group[AppData.KEY_GROUP_ID] = index
        if for_group_with_id is not None:
            for index, entry in enumerate(self.groups_data[for_group_with_id][AppData.KEY_GROUP_ENTRIES]):
                entry[AppData.KEY_ENTRY_ID] = index


class AppInterface:
    """
    Base class for all UI interfaces in the application.

    Each interface must define its UI layout and UI events and assign them to the `self.win_layout` and 
    `self.win_events` properties (the first interface must do it in the constructor, not in the `start` method). 
    
    The `self.win_layout` property must be a PySimpleGUI layout.

    The `self.win_events` property must be a `dict` with the event name as key and a list of callables as value
    (each callable must receive the event data as an argument).

    After the `App` class instance completes its initialization (which involves instantiating the interfaces), 
    it gives a reference to itself (`self.app` property) to each registered interface in the class and calls their 
    `start` method. Therefore, an interface should do its basic initialization in the constructor, and do in the 
    `start` method the initialization which involves other systems/interfaces in the application (accessed via the 
    `self.app` property).

    The `on_show` method is called each time the interface becomes visible on the screen.
    """

    def __init__(self) -> None:
        self.app = None             
        self.win_layout = [[]]      
        self.win_events = dict()    

    def start(self) -> None:
        pass

    def on_show(self) -> None:
        pass


class MainInterface(AppInterface):
    """
    The main interface of the application. It is the first interface shown.
    """

    KEY_TREE_MAIN = '-MAIN_TREE-'
    KEY_BUTTON_OPEN_GROUP = '-OPEN_GROUP_BUTTON-'
    KEY_BUTTON_NEW_GROUP = '-NEW_GROUP_BUTTON-'
    KEY_BUTTON_EDIT_GROUP = '-EDIT_GROUP_BUTTON-'
    KEY_BUTTON_DELETE_GROUP = '-DELETE_GROUP_BUTTON-'
    KEY_BUTTON_SAVE_CHANGES = '-SAVE_CHANGES_BUTTON-'
    KEY_BUTTON_REVERT_CHANGES = '-REVERT_CHANGES_BUTTON-'
    KEY_BUTTON_HELP = '-HELP_BUTTON-'
    KEY_BUTTON_EXIT = '-EXIT_BUTTON-'
    KEY_STATUS_BAR_GROUPS = '-GROUPS_STATUS_BAR-'
    KEY_CHECKBOX_START_WITH_WINDOWS = '-START_WITH_WINDOWS-'
    KEY_CHECKBOX_AUTO_SAVE = '-AUTO_SAVE-'

    def __init__(self) -> None:
        super().__init__()
        self.win_layout = [[
                sg.Tree(data=sg.TreeData(), 
                        key=self.KEY_TREE_MAIN, 
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
                        [sg.Button('Open Group', key=self.KEY_BUTTON_OPEN_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                        [sg.Button('New Group', key=self.KEY_BUTTON_NEW_GROUP, expand_x=True, pad=((0, 0), (7, 0)))],
                        [sg.Button('Edit Group', key=self.KEY_BUTTON_EDIT_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                        [sg.Button('Delete Group', key=self.KEY_BUTTON_DELETE_GROUP, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                        [sg.Button('Save Changes', key=self.KEY_BUTTON_SAVE_CHANGES, expand_x=True, pad=((0, 0), (35, 0)), disabled=True)],
                        [sg.Button('Revert Changes', key=self.KEY_BUTTON_REVERT_CHANGES, expand_x=True, pad=((0, 0), (7, 0)), disabled=True)],
                        [sg.Button('Help', key=self.KEY_BUTTON_HELP, expand_x=True, pad=((0, 0), (35, 0)))],
                        [sg.Button('Exit', key=self.KEY_BUTTON_EXIT, expand_x=True, pad=((0, 0), (7, 7)))],
                        ],
                        vertical_alignment='top',
                        pad=((15, 15), (15, 15)))
            ],[
                sg.StatusBar('', 
                        key=self.KEY_STATUS_BAR_GROUPS, 
                        size=50, 
                        pad=((15, 15), (10, 15)), 
                        auto_size_text=True,
                        font=('Verdana', 10, 'normal'))
            ],[
                sg.Checkbox('Launch the application when Windows starts', key=self.KEY_CHECKBOX_START_WITH_WINDOWS, enable_events=True),
                sg.Checkbox('Auto save changes (once per minute)', key=self.KEY_CHECKBOX_AUTO_SAVE, enable_events=True)
            ]
        ]
        self.win_events = {
            self.KEY_TREE_MAIN: [self.on_tree_event],
            self.KEY_BUTTON_OPEN_GROUP: [self.on_button_open_group],
            self.KEY_BUTTON_NEW_GROUP: [self.on_button_new_group],
            self.KEY_BUTTON_EDIT_GROUP: [self.on_button_edit_group],
            self.KEY_BUTTON_DELETE_GROUP: [self.on_button_delete_group],
            self.KEY_BUTTON_SAVE_CHANGES: [self.on_button_save_changes],
            self.KEY_BUTTON_REVERT_CHANGES: [self.on_button_revert_changes],
            self.KEY_BUTTON_HELP: [self.on_button_help],
            self.KEY_BUTTON_EXIT: [self.on_button_exit],
            self.KEY_CHECKBOX_START_WITH_WINDOWS: [self.on_checkbox_start_with_windows],
            self.KEY_CHECKBOX_AUTO_SAVE: [self.on_checkbox_auto_save]
        }
        self.groups_status_bar_lock = threading.Lock()
        self.check_if_saving_lock = threading.Lock()
        self.is_saving = threading.Event()
        self.made_changes = False
        self.tree_dirty = False
    
    def start(self) -> None:
        super().start()
        settings = self.app.app_data.settings
        window = self.app.window

        window[self.KEY_TREE_MAIN].update(self.__create_groups_tree_data())
        if AppData.KEY_SETTING_START_WITH_WINDOWS in settings and settings[AppData.KEY_SETTING_START_WITH_WINDOWS]:
            window[self.KEY_CHECKBOX_START_WITH_WINDOWS].update(value=True)
        if AppData.KEY_SETTING_AUTO_SAVE in settings and settings[AppData.KEY_SETTING_AUTO_SAVE]:
            window[self.KEY_CHECKBOX_AUTO_SAVE].update(value=True)

        if sg.WIN_CLOSE_ATTEMPTED_EVENT in self.app.win_global_events:
            self.app.win_global_events[sg.WIN_CLOSE_ATTEMPTED_EVENT].append(self.__on_exit)
        else:
            self.app.win_global_events[sg.WIN_CLOSE_ATTEMPTED_EVENT] = [self.__on_exit]

        self.auto_save_enabled = threading.Event()
        if AppData.KEY_SETTING_AUTO_SAVE in settings and settings[AppData.KEY_SETTING_AUTO_SAVE]:
            self.auto_save_enabled.set()
        threading.Thread(target=self.__async_auto_saving, daemon=True).start()

    def on_show(self) -> None:
        super().on_show()
        window = self.app.window
        if self.tree_dirty:
            window[self.KEY_TREE_MAIN].update(self.__create_groups_tree_data())
            self.tree_dirty = False
        if self.made_changes:
            window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
            self.made_changes = False

    def update_groups_status_bar(self, message) -> None:
        with self.groups_status_bar_lock:
            self.app.window[self.KEY_STATUS_BAR_GROUPS].update(value=message)
    
    def on_tree_event(self, values) -> None:
        window = self.app.window
        should_disable = False if values[self.KEY_TREE_MAIN] else True
        window[self.KEY_BUTTON_OPEN_GROUP].update(disabled=should_disable)
        window[self.KEY_BUTTON_EDIT_GROUP].update(disabled=should_disable)
        window[self.KEY_BUTTON_DELETE_GROUP].update(disabled=should_disable)
    
    def on_button_open_group(self, values) -> None:
        app_data = self.app.app_data
        selected_group_id = values[self.KEY_TREE_MAIN][0]
        group_name = app_data.groups_data[selected_group_id][AppData.KEY_GROUP_NAME]
        entries_to_open = deepcopy(app_data.groups_data[selected_group_id][AppData.KEY_GROUP_ENTRIES])
        threading.Thread(target=self.__async_group_opening, args=(group_name, entries_to_open)).start()
    
    def on_button_new_group(self, _) -> None:
        window = self.app.window
        group_name = sg.popup_get_text('How should the group be named?', 'Input a name for the new group')
        if group_name is not None and group_name != '':
            self.app.app_data.groups_data.append({
                AppData.KEY_GROUP_ID: len(self.app.app_data.groups_data),
                AppData.KEY_GROUP_NAME: group_name, 
                AppData.KEY_GROUP_ENTRIES: []
            })
            window[self.KEY_TREE_MAIN].update(self.__create_groups_tree_data())
            window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
            window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
    
    def on_button_edit_group(self, values) -> None:
        selected_group_id = values[self.KEY_TREE_MAIN][0]
        group_to_edit = self.app.app_data.groups_data[selected_group_id]
        self.app.get_interface(App.KEY_INTERFACE_GROUP_EDIT).group = group_to_edit
        self.app.change_shown_interface(App.KEY_INTERFACE_GROUP_EDIT)

    def on_button_delete_group(self, values) -> None:
        app_data = self.app.app_data
        window = self.app.window
        tmp = []
        for group in app_data.groups_data:
            if group[AppData.KEY_GROUP_ID] not in values[self.KEY_TREE_MAIN]:
                tmp.append(group)
            elif AppData.KEY_GROUP_LISTENING_DIR in group:
                self.app.get_interface(App.KEY_INTERFACE_GROUP_EDIT).remove_listener(group)
        app_data.groups_data = tmp
        app_data.refresh_ids(for_groups=True)
        window[self.KEY_TREE_MAIN].update(self.__create_groups_tree_data())
        window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    def on_button_save_changes(self, _) -> None:
        self.__save_changes()

    def on_button_revert_changes(self, _)-> None:
        app_data = self.app.app_data
        window = self.app.window
        app_data.revert_changes()
        window[self.KEY_CHECKBOX_START_WITH_WINDOWS].update(app_data.settings.get(AppData.KEY_SETTING_START_WITH_WINDOWS, False))
        window[self.KEY_CHECKBOX_AUTO_SAVE].update(app_data.settings.get(AppData.KEY_SETTING_AUTO_SAVE, False))
        window[self.KEY_TREE_MAIN].update(self.__create_groups_tree_data())
        window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
        window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
        self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been reverted!")

    def on_button_help(self, _) -> None:
        self.app.change_shown_interface(App.KEY_INTERFACE_HELP)
    
    def on_button_exit(self, _) -> None:
        self.__on_exit(None)

    def on_checkbox_start_with_windows(self, values) -> None:
        self.app.app_data.settings[AppData.KEY_SETTING_START_WITH_WINDOWS] = values[self.KEY_CHECKBOX_START_WITH_WINDOWS]
        self.app.window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        self.app.window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    def on_checkbox_auto_save(self, values) -> None:
        checked = values[self.KEY_CHECKBOX_AUTO_SAVE]
        if checked:
            self.auto_save_enabled.set()
        else:
            self.auto_save_enabled.clear()
        self.app.app_data.settings[AppData.KEY_SETTING_AUTO_SAVE] = checked
        self.app.window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        self.app.window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    def __async_group_opening(self, group_name, entries_to_open) -> None:
        self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Opening \"{group_name}\" group...")
        tabs_opened = 0
        for entry in entries_to_open:
            if entry[AppData.KEY_ENTRY_TYPE] == AppData.ENTRY_WEB_PAGE:
                # For some browsers (e.g. Firefox), if they are closed when group opening starts, it may be required to wait a bit for 
                # them to initialize on opening the first tab before opening others.
                if tabs_opened == 1:
                    time.sleep(1.5)
                # Using this approach because os.startfile(URL) will not work if the URL specified by user doesn't contain the protocol.
                # Also, in this case, webbrowser.open(URL) will use Microsoft Edge regardless of the default browser.
                subprocess.Popen([self.app.default_browser_path, entry[AppData.KEY_ENTRY_PATH]], creationflags=subprocess.DETACHED_PROCESS)
                tabs_opened += 1
            elif os.path.isfile(entry[AppData.KEY_ENTRY_PATH]):
                os.startfile(entry[AppData.KEY_ENTRY_PATH])
        self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Group \"{group_name}\" opened.")
    
    def __async_auto_saving(self) -> None:
        auto_save_delay_s = 60
        while True:
            time.sleep(auto_save_delay_s)
            if self.auto_save_enabled.is_set() and not self.app.window[self.KEY_BUTTON_SAVE_CHANGES].Disabled:
                self.__save_changes()

    def __create_groups_tree_data(self) -> sg.TreeData:
        new_tree_data = sg.TreeData()
        for group in self.app.app_data.groups_data:
            new_tree_data.insert('', group[AppData.KEY_GROUP_ID], f'  {group[AppData.KEY_GROUP_NAME]}', [len(group[AppData.KEY_GROUP_ENTRIES])], AppData.ICON_OTHER_FILE)
        return new_tree_data
    
    def __save_changes(self) -> None:
        with self.check_if_saving_lock:
            if self.is_saving.is_set():
                return
            else:
                self.is_saving.set()

        if threading.current_thread() is threading.main_thread():
            self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Saving changes...")
        else:
            self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Auto save: Saving changes...")

        window = self.app.window
        app_data = self.app.app_data
        window[self.KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
        window[self.KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
        app_data.save_group_data()
        app_data.save_settings()
        self.__set_start_with_windows(app_data.settings.get(AppData.KEY_SETTING_START_WITH_WINDOWS, False))

        if threading.current_thread() is threading.main_thread():
            self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been successfully saved!")
        else:
            self.update_groups_status_bar(f"({time.strftime('%H:%M:%S', time.localtime())}) Auto save: Changes have been successfully saved!")

        self.is_saving.clear()
    
    def __set_start_with_windows(self, enabled: bool) -> None:
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
    
    def __on_exit(self, _) -> None:
        if self.app.window[self.KEY_BUTTON_SAVE_CHANGES].Disabled:
            self.app.signal_shutdown()
            return
        response = sg.popup_yes_no('\n== Save changes? ==\n')
        if response is not None:
            if response == 'Yes':
                self.__save_changes()
            self.app.signal_shutdown()


class GroupEditInterface(AppInterface):
    """The interface shown for group editing."""

    KEY_TREE_GROUP_EDIT = '-GROUP_EDIT_TREE-'
    KEY_BUTTON_ADD_FILES = '-ADD_FILES_BUTTON-'
    KEY_BUTTON_ADD_WEB_PAGE = '-ADD_WEB_PAGE_BUTTON-'
    KEY_BUTTON_DELETE_ENTRIES = '-DELETE_ENTRIES_BUTTON-'
    KEY_BUTTON_EDIT_DETAILS = '-EDIT_DETAILS_BUTTON-'
    KEY_BUTTON_BACK = '-GROUP_EDIT_BACK_BUTTON-'
    KEY_BUTTON_LISTEN_TO_DIR = 'LISTEN_DIR_BUTTON'
    KEY_STATUS_BAR_ENTRIES = '-ENTRIES_STATUS_BAR-'
    KEY_LABEL_GROUP_NAME = '-GROUP_NAME_LABEL-'

    def __init__(self) -> None:
        super().__init__()
        self.win_layout = [[
                sg.Text('<Name> Group', key=self.KEY_LABEL_GROUP_NAME, font=('Arial', 16, 'normal'), pad=15)
            ],[
                sg.Tree(data=sg.TreeData(), 
                        key=self.KEY_TREE_GROUP_EDIT, 
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
                        sg.Input(key=self.KEY_BUTTON_ADD_FILES, visible=False, enable_events=True), 
                        sg.FilesBrowse('Add Files', self.KEY_BUTTON_ADD_FILES),
                        sg.Button('Add Web Page', key=self.KEY_BUTTON_ADD_WEB_PAGE),
                        sg.Button('Delete Selected Entries', key=self.KEY_BUTTON_DELETE_ENTRIES, disabled=True),
                        sg.Button('Edit Selected Details', key=self.KEY_BUTTON_EDIT_DETAILS, disabled=True),
                        sg.Button('Start Listening', key=self.KEY_BUTTON_LISTEN_TO_DIR),
                        sg.Button('Back', key=self.KEY_BUTTON_BACK),
                        ]],
                        pad=((15, 15), (0, 7))),
                sg.Stretch()
            ],[
                sg.StatusBar('', 
                        key=self.KEY_STATUS_BAR_ENTRIES, 
                        size=80, 
                        pad=((15, 15), (10, 15)),
                        auto_size_text=True,
                        font=('Verdana', 10, 'normal'))
            ]
        ]
        self.win_events = {
            self.KEY_TREE_GROUP_EDIT: [self.on_tree_event],
            self.KEY_BUTTON_ADD_FILES: [self.on_button_add_files],
            self.KEY_BUTTON_ADD_WEB_PAGE: [self.on_button_add_web_page],
            self.KEY_BUTTON_DELETE_ENTRIES: [self.on_button_delete_entries],
            self.KEY_BUTTON_EDIT_DETAILS: [self.on_button_edit_details],
            self.KEY_BUTTON_LISTEN_TO_DIR: [self.on_button_listen_dir],
            self.KEY_BUTTON_BACK: [self.on_button_back]
        }
        self.group = None
    
    def start(self) -> None:
        super().start()
        self.listening_event_handler = PatternMatchingEventHandler('*', None, True, True)
        self.listening_event_handler.on_created = self.__listening_on_created
        self.listening_event_handler.on_deleted = self.__listening_on_deleted
        self.listening_event_handler.on_moved = self.__listening_on_moved
        self.listening_observer = Observer()
        self.listening_watches = []
        for group in self.app.app_data.groups_data:
            if AppData.KEY_GROUP_LISTENING_DIR in group:
                if os.path.isdir(group[AppData.KEY_GROUP_LISTENING_DIR]):
                    self.listening_watches.append(self.listening_observer.schedule(self.listening_event_handler, group[AppData.KEY_GROUP_LISTENING_DIR]))
                else:
                    group.pop(AppData.KEY_GROUP_LISTENING_DIR)
        self.listening_observer.start()
        self.app.on_shutdown_actions.append(self.__on_app_shutdown)

    def on_show(self) -> None:
        super().on_show()
        window = self.app.window
        window[self.KEY_LABEL_GROUP_NAME].update(f'{self.group[AppData.KEY_GROUP_NAME]} Group')
        window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
        if AppData.KEY_GROUP_LISTENING_DIR in self.group:
            window[self.KEY_BUTTON_LISTEN_TO_DIR].update('Stop Listening')
            window[self.KEY_STATUS_BAR_ENTRIES].update(value=f"Listening on {self.group[AppData.KEY_GROUP_LISTENING_DIR]}")
        else:
            window[self.KEY_BUTTON_LISTEN_TO_DIR].update('Start Listening')
            window[self.KEY_STATUS_BAR_ENTRIES].update(value='')
    
    def remove_listener(self, group) -> None:
        path_to_listen = group[AppData.KEY_GROUP_LISTENING_DIR]
        watch = [watch for watch in self.listening_watches if watch.path == path_to_listen][0]
        self.listening_observer.unschedule(watch)
        self.listening_watches.remove(watch)

    def on_tree_event(self, values) -> None:
        window = self.app.window
        should_disable = False if values[self.KEY_TREE_GROUP_EDIT] else True
        window[self.KEY_BUTTON_DELETE_ENTRIES].update(disabled=should_disable)
        window[self.KEY_BUTTON_EDIT_DETAILS].update(disabled=should_disable)

    def on_button_add_files(self, values) -> None:
        window = self.app.window
        new_entries = values[self.KEY_BUTTON_ADD_FILES].split(';')
        for index, entry in enumerate(new_entries, start=len(self.group[AppData.KEY_GROUP_ENTRIES])):
            if entry.split('.')[-1] == 'exe':
                self.group[AppData.KEY_GROUP_ENTRIES].append({
                    AppData.KEY_ENTRY_ID: index,
                    AppData.KEY_ENTRY_PATH: entry, 
                    AppData.KEY_ENTRY_TYPE: AppData.ENTRY_EXE_FILE, 
                    AppData.KEY_ENTRY_DETAILS: ''
                })
            else:
                self.group[AppData.KEY_GROUP_ENTRIES].append({
                    AppData.KEY_ENTRY_ID: index,
                    AppData.KEY_ENTRY_PATH: entry, 
                    AppData.KEY_ENTRY_TYPE: AppData.ENTRY_OTHER_FILE, 
                    AppData.KEY_ENTRY_DETAILS: ''
                })
        window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
        main_interface = self.app.get_interface(App.KEY_INTERFACE_MAIN)
        main_interface.made_changes = True
        main_interface.tree_dirty = True

    def on_button_add_web_page(self, values) -> None:
        window = self.app.window
        added_url = sg.popup_get_text("What's the web address of your desired web page?", "Input URL")
        if added_url is not None and added_url != '':
            self.group[AppData.KEY_GROUP_ENTRIES].append({
                AppData.KEY_ENTRY_ID: len(self.group[AppData.KEY_GROUP_ENTRIES]),
                AppData.KEY_ENTRY_PATH: added_url,
                AppData.KEY_ENTRY_TYPE: AppData.ENTRY_WEB_PAGE, 
                AppData.KEY_ENTRY_DETAILS: ''
            })
            window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
            main_interface = self.app.get_interface(App.KEY_INTERFACE_MAIN)
            main_interface.made_changes = True
            main_interface.tree_dirty = True

    def on_button_delete_entries(self, values) -> None:
        window = self.app.window
        self.group[AppData.KEY_GROUP_ENTRIES][:] = [entry for entry in self.group[AppData.KEY_GROUP_ENTRIES] if entry[AppData.KEY_ENTRY_ID] not in values[self.KEY_TREE_GROUP_EDIT]]
        self.app.app_data.refresh_ids(for_group_with_id=self.group[AppData.KEY_GROUP_ID])
        window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
        main_interface = self.app.get_interface(App.KEY_INTERFACE_MAIN)
        main_interface.made_changes = True
        main_interface.tree_dirty = True

    def on_button_edit_details(self, values) -> None:
        window = self.app.window
        selected_entries = [entry for entry in self.group[AppData.KEY_GROUP_ENTRIES] if entry[AppData.KEY_ENTRY_ID] in values[self.KEY_TREE_GROUP_EDIT]]
        old_details = selected_entries[0][AppData.KEY_ENTRY_DETAILS] if len(selected_entries) == 1 else ''
        new_details = sg.popup_get_text('What are the details of those entries?', 'Input details for selected entries', old_details)
        if new_details is not None:
            for entry in selected_entries:
                entry[AppData.KEY_ENTRY_DETAILS] = new_details
            window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
            self.app.get_interface(App.KEY_INTERFACE_MAIN).made_changes = True

    def on_button_listen_dir(self, _) -> None:
        window = self.app.window
        if window[self.KEY_BUTTON_LISTEN_TO_DIR].get_text().startswith('Start'):
            path_to_listen = sg.popup_get_folder('Select the directory to listen:', title='Select Directory')
            if path_to_listen is not None and path_to_listen != '':
                self.listening_watches.append(self.listening_observer.schedule(self.listening_event_handler, path_to_listen))
                self.group[AppData.KEY_GROUP_LISTENING_DIR] = path_to_listen
                window[self.KEY_BUTTON_LISTEN_TO_DIR].update('Stop Listening')
                window[self.KEY_STATUS_BAR_ENTRIES].update(value=f"Listening on {path_to_listen}")
                self.app.get_interface(App.KEY_INTERFACE_MAIN).made_changes = True
        else:
            self.remove_listener(self.group)
            self.group.pop(AppData.KEY_GROUP_LISTENING_DIR)
            window[self.KEY_BUTTON_LISTEN_TO_DIR].update('Start Listening')
            window[self.KEY_STATUS_BAR_ENTRIES].update(value='')
            self.app.get_interface(App.KEY_INTERFACE_MAIN).made_changes = True

    def on_button_back(self, _) -> None:
        self.group = None
        self.app.change_shown_interface(App.KEY_INTERFACE_MAIN)

    def __create_entries_tree_data(self) -> sg.TreeData:
        new_tree_data = sg.TreeData()
        for entry in self.group[AppData.KEY_GROUP_ENTRIES]:
            icon = AppData.ICON_OTHER_FILE
            if entry[AppData.KEY_ENTRY_TYPE] == AppData.ENTRY_EXE_FILE:
                icon = AppData.ICON_EXE_FILE
            elif entry[AppData.KEY_ENTRY_TYPE] == AppData.ENTRY_WEB_PAGE:
                icon = AppData.ICON_WEB_PAGE
            new_tree_data.insert('', entry[AppData.KEY_ENTRY_ID], f'  {entry[AppData.KEY_ENTRY_PATH]}', [entry[AppData.KEY_ENTRY_DETAILS]], icon)
        return new_tree_data

    def __listening_on_created(self, event):
        file_path = event.src_path.replace(os.sep, '/')
        if file_path.split('.')[-1] == 'exe':
            entry_type = AppData.ENTRY_EXE_FILE
        else:
            entry_type = AppData.ENTRY_OTHER_FILE
        path_to_listen = os.path.split(file_path)[0]
        
        for group in self.app.app_data.groups_data:
            if AppData.KEY_GROUP_LISTENING_DIR in group and group[AppData.KEY_GROUP_LISTENING_DIR] == path_to_listen:
                group[AppData.KEY_GROUP_ENTRIES].append({
                    AppData.KEY_ENTRY_ID: len(group[AppData.KEY_GROUP_ENTRIES]),
                    AppData.KEY_ENTRY_PATH: file_path, 
                    AppData.KEY_ENTRY_TYPE: entry_type, 
                    AppData.KEY_ENTRY_DETAILS: ''
                })
                self.app.window[MainInterface.KEY_TREE_MAIN].update(key=group[AppData.KEY_GROUP_ID], value=[len(group[AppData.KEY_GROUP_ENTRIES])])
        
        self.__on_listening_event(path_to_listen)

    def __listening_on_deleted(self, event) -> None:
        file_path = event.src_path.replace(os.sep, '/')
        path_to_listen = os.path.split(file_path)[0]

        for group in self.app.app_data.groups_data:
            if AppData.KEY_GROUP_LISTENING_DIR in group and group[AppData.KEY_GROUP_LISTENING_DIR] == path_to_listen:
                group[AppData.KEY_GROUP_ENTRIES][:] = [entry for entry in group[AppData.KEY_GROUP_ENTRIES] if entry[AppData.KEY_ENTRY_PATH] != file_path]
                self.app.app_data.refresh_ids(for_group_with_id=group[AppData.KEY_GROUP_ID])
                self.app.window[MainInterface.KEY_TREE_MAIN].update(key=group[AppData.KEY_GROUP_ID], value=[len(group[AppData.KEY_GROUP_ENTRIES])])
        
        self.__on_listening_event(path_to_listen)

    def __listening_on_moved(self, event) -> None:
        old_file_path = event.src_path.replace(os.sep, '/')
        new_file_path = event.dest_path.replace(os.sep, '/')
        path_to_listen = os.path.split(old_file_path)[0]

        for group in self.app.app_data.groups_data:
            if AppData.KEY_GROUP_LISTENING_DIR in group and group[AppData.KEY_GROUP_LISTENING_DIR] == path_to_listen:
                for entry in group[AppData.KEY_GROUP_ENTRIES]:
                    if entry[AppData.KEY_ENTRY_PATH] == old_file_path:
                        entry[AppData.KEY_ENTRY_PATH] = new_file_path
        
        self.__on_listening_event(path_to_listen)

    def __on_listening_event(self, path_to_listen) -> None:
        window = self.app.window
        if self.group is not None and AppData.KEY_GROUP_LISTENING_DIR in self.group and self.group[AppData.KEY_GROUP_LISTENING_DIR] == path_to_listen:
            window[self.KEY_TREE_GROUP_EDIT].update(self.__create_entries_tree_data())
        window[MainInterface.KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[MainInterface.KEY_BUTTON_REVERT_CHANGES].update(disabled=False)
        window.refresh()

    def __on_app_shutdown(self) -> None:
        self.listening_observer.stop()
        self.listening_observer.join()


class HelpInterface(AppInterface):
    """The interface shown when pressing the "Help" button from the main interface."""

    KEY_HYPERLINK_HELP = '-HELP_HYPERLINK-'
    KEY_BUTTON_BACK = '-HELP_BACK_BUTTON-'

    def __init__(self) -> None:
        super().__init__()
        self.win_layout = [[
                sg.VStretch(),
                sg.Column([
                        [sg.Text('Check out the documentation in the following GitHub repository:', font='Verdana 10 normal')],
                        [sg.Text('https://github.com/ioan45/open-my-files', key=self.KEY_HYPERLINK_HELP, enable_events=True, font='Verdana 10 underline', text_color='LightBlue')],
                        [sg.Button('Back', key=self.KEY_BUTTON_BACK, pad=((0, 0), (20, 10)))]
                        ],
                        element_justification='center',
                        justification='center'),
                sg.VStretch()
            ]
        ]
        self.win_events = {
            self.KEY_HYPERLINK_HELP: [self.on_hyperlink_click],
            self.KEY_BUTTON_BACK: [self.on_button_back]
        }
    
    def on_hyperlink_click(self, _) -> None:
        os.startfile('https://github.com/ioan45/open-my-files')
    
    def on_button_back(self, _) -> None:
        self.app.change_shown_interface(App.KEY_INTERFACE_MAIN)
    

class App:
    """
    The core part of the application. 
    
    It manages the interfaces initialization and switching, the events handling, and runs the PySimpleGUI window.
    It also connects all parts of the application (e.g. interfaces can communicate with each other and all of them 
    can access the application data manager and the PySimpleGUI window).
    """

    KEY_INTERFACE_MAIN = '-MAIN_INTERFACE-'
    KEY_INTERFACE_GROUP_EDIT = '-GROUP_EDIT_INTERFACE-'
    KEY_INTERFACE_HELP = '-HELP_INTERFACE-'
    
    def __init__(self) -> None:
        groups_data_path = os.path.join(os.getenv('LOCALAPPDATA'), 'OpenMyFiles', 'groups.json')
        settings_data_path = os.path.join(os.getenv('LOCALAPPDATA'), 'OpenMyFiles', 'settings.json')
        self.app_data = AppData(groups_data_path, settings_data_path)
        self.app_data.load_groups_data()
        self.app_data.load_settings()

        sg.theme('DarkAmber')

        self.__interfaces = {
            App.KEY_INTERFACE_MAIN: MainInterface(),
            App.KEY_INTERFACE_GROUP_EDIT: GroupEditInterface(),
            App.KEY_INTERFACE_HELP: HelpInterface()
        }
        self.__current_interface_key = App.KEY_INTERFACE_MAIN
        self.default_browser_path = get_default_browser_path()
        self.__app_running = False

        win_layout = [[sg.Column(
                        layout=interface.win_layout, 
                        key=interface_key, 
                        visible=False if interface_key != self.__current_interface_key else True, 
                        expand_x=True, 
                        expand_y=True)
                        for interface_key, interface in self.__interfaces.items()
        ]]
        self.window = sg.Window('Open My Files', win_layout, resizable=True, enable_close_attempted_event=True, finalize=True)

        # Events which will be checked always. Key:event_key, Value:list[action(event_values)]
        self.win_global_events = dict()                                                     
        # Interface events which will be checked if the interface is shown. Key:event_key, Value:list[action(event_values)]    
        self.win_interface_events = self.__interfaces[self.__current_interface_key].win_events 
        # List of actions which will be performed on app shutdown. list[action()]
        self.on_shutdown_actions = []

        for interface in self.__interfaces.values():
            interface.app = self
            interface.start()
        self.window.refresh()

    def run(self) -> None:
        self.__app_running = True
        while self.__app_running:
            event, values = self.window.read()
            if event in self.win_global_events:
                actions = copy(self.win_global_events[event])
                for action in actions:
                    action(values)
            if event in self.win_interface_events:
                actions = copy(self.win_interface_events[event])
                for action in actions:
                    action(values)
            if event == sg.WIN_CLOSE_ATTEMPTED_EVENT and event not in self.win_global_events and event not in self.win_interface_events:
                self.signal_shutdown()
        self.__shutdown()
    
    def signal_shutdown(self) -> None:
        self.__app_running = False

    def change_shown_interface(self, new_interface_key: str) -> None:
        self.window[self.__current_interface_key].update(visible=False)
        self.window[new_interface_key].update(visible=True)
        self.win_interface_events = self.__interfaces[new_interface_key].win_events
        self.__interfaces[new_interface_key].on_show()
        self.__current_interface_key = new_interface_key
        self.window.disappear()
        self.window.refresh()
        self.window.move_to_center()
        self.window.reappear()

    def get_interface(self, interface_key: str) -> AppInterface:
        return self.__interfaces[interface_key]
    
    def get_current_interface_key(self) -> str:
        return self.__current_interface_key
    
    def __shutdown(self) -> None:
        for action in self.on_shutdown_actions:
            action()
        self.window.close()


def get_default_browser_path() -> str:
    """Returns the path to the default browser which is set in the Windows settings."""
    
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.html\\UserChoice") as id_key:
        browser_id = winreg.QueryValueEx(id_key, 'ProgId')[0]
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{browser_id}\\shell\\open\\command") as path_key:
        browser_path = winreg.QueryValueEx(path_key, '')[0]
    return browser_path.split('"')[1]

def main():
    app = App()
    app.run()

if __name__ == "__main__":
    main()
