import os
import winreg
import subprocess
import threading
import json
import time
import PySimpleGUI as sg
from copy import deepcopy

ENTRY_EXE_FILE = 'executable'
ENTRY_OTHER_FILE = 'other'
ENTRY_WEB_PAGE = 'web_page'

ICON_EXE_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAItJREFUOE/tksENg0AMBGfrCKUkBSAKQIg/PYU/6SAdIBqBPhYZiSiPiOguPOOPJd/teGVb/BiyXQF34JLIWoAuADNQS5pSALavwCMAlhT5GQXAQCspnB3Gpt0B3z5/en8BcsS75hwHMYMcF+fN4O9gO+VG0piyCds3YIg7KIEeKFIAQDTusvb/3mgFDG1uMt84gY0AAAAASUVORK5CYII='
ICON_OTHER_FILE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAL9JREFUOE/lkz0KAjEQRt93DkWw0M4jWFiKBxBFsPMaeg1txTt4APEQCiJ6kE8C2SUbVt3GyjSTTCZvfjIj4rI9AbZAq9Al8g6MJAVZWUoAT2Aq6Zwb2d4AyzpICrCk8pxCImANPCLkVtw3BSyAfXxUSacpINjNgUGEhFS7Yd8IUFOTMt3fAGxfgF7m+SqpH3S2/ymCMu+ssX5cg5phqqje/UIYppmk0yeA7SFwkNTJO3EM7ID2lwiCo5WkY7B7ATUxkhEHIy87AAAAAElFTkSuQmCC'
ICON_WEB_PAGE = b'iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAAXNSR0IArs4c6QAAAa5JREFUOE+d08uLDmAUx/HP2bgtJn8CWQkhFhYyyiUsxUrKLWwkk5KNZmyk3BZqmolmZSNrNOOWKGUi5bIS/gEiuS6Oztvz6u0tNTmb09Nznu85v/OcE/osM5fhADZgQbt+jzu4EhEve59E95CZs3ER2zEXT/Gh3RdoNb7jBoYi4lfddQDt8S18xDWcx4oKbIACP8dx7MJ8bC1IFzCK3XiBeS3LBZxogLMNthNfG3wiIo5E03wfi7EZV/EWC7sVVpF4h0XYhym8wWABLlXpEXE6M6txpyJisEk709EZcbKdH2IkIu5m5jAGCvAKByPicWYewsqIONweVFABOj4zx/AsIsYycy1GC/AFpbc6vB5zcLtp39h8fWHZFvzAg/ZTQ13ARAOsaYAKKKssZY+aL2k/8aQB9vZLqNJLQkmpkvsljGM6IsZ7JVQTP0XEyAybOBwR93qbuLRp+p9vXNcdpMvYM4NB2oFvWF7zEhFHu4BZuInPbZTPVS9w7B+jPIBtEfG7d5kKUjtQWWqcp1FbWFZTuaplv147UY//LlML6rjMXIL92NS3zpOt7Ne98X8AKibG9Jc6ezQAAAAASUVORK5CYII='

KEY_GROUP_ID = 'group_id'
KEY_GROUP_NAME = 'group_name'
KEY_GROUP_ENTRIES = 'group_entries'
KEY_ENTRY_ID = 'entry_id'
KEY_ENTRY_PATH = 'entry_path'
KEY_ENTRY_TYPE = 'entry_type'
KEY_ENTRY_DETAILS = 'entry_details'

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
KEY_STATUS_BAR = '-STATUS_BAR-'
KEY_LABEL_GROUP_NAME = '-GROUP_NAME_LABEL-'
KEY_HYPERLINK_HELP = '-HELP_HYPERLINK-'

EVENT_OPENING_COMPLETE = '-OPENING_COMPLETE-'

GROUPS_DATA_PATH = os.path.join(os.getenv('LOCALAPPDATA'), 'OpenMyFiles', 'groups.json')

def get_default_browser_path():
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\FileExts\\.html\\UserChoice") as id_key:
        browser_id = winreg.QueryValueEx(id_key, 'ProgId')[0]
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, f"{browser_id}\\shell\\open\\command") as path_key:
        browser_path = winreg.QueryValueEx(path_key, '')[0]
    return browser_path.split('"')[1]

DEFAULT_BROWSER_PATH = get_default_browser_path()

class AppData:  
    def __init__(self, groups_data_path: str) -> None:
        self.__saved_groups_data = []
        self.groups_path = groups_data_path
        self.groups_data = []
    
    def load_data(self):
        if os.path.isfile(self.groups_path):
            with open(self.groups_path, 'r') as f:
                self.__saved_groups_data = json.load(f)
            self.groups_data = deepcopy(self.__saved_groups_data)
    
    def save_data(self):
        self.__saved_groups_data = deepcopy(self.groups_data)
        os.makedirs(os.path.split(self.groups_path)[0], exist_ok=True)
        with open(self.groups_path, 'w') as f:
            json.dump(self.__saved_groups_data, f, indent='\t')
    
    def revert_changes(self):
        self.groups_data = deepcopy(self.__saved_groups_data)
    
    def refresh_ids(self, for_groups=False, for_group_with_id=None):
        if for_groups:
            for index, group in enumerate(self.groups_data):
                group[KEY_GROUP_ID] = index
        if for_group_with_id is not None:
            for index, entry in enumerate(self.groups_data[for_group_with_id][KEY_GROUP_ENTRIES]):
                entry[KEY_ENTRY_ID] = index

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

def async_group_opening(group_to_open):
    tabs_opened = 0
    for entry in group_to_open[KEY_GROUP_ENTRIES]:
        if entry[KEY_ENTRY_TYPE] == ENTRY_WEB_PAGE:
            # For some browsers (e.g. Firefox), if they are closed when group opening starts, it may be required to wait a bit for 
            # them to initialize on opening the first tab before opening others.
            if tabs_opened == 1:
                time.sleep(1.5)
            # Using this approach because os.startfile(URL) will not work if the URL specified by user doesn't contain the protocol.
            # Also, in this case, webbrowser.open(URL) will use Microsoft Edge regardless of the default browser.
            subprocess.Popen([DEFAULT_BROWSER_PATH, entry[KEY_ENTRY_PATH]], creationflags=subprocess.DETACHED_PROCESS)
            tabs_opened += 1
        else:
            os.startfile(entry[KEY_ENTRY_PATH])
    window.write_event_value(EVENT_OPENING_COMPLETE, None)

sg.theme('DarkAmber')

app_data = AppData(GROUPS_DATA_PATH)
app_data.load_data()

group_to_edit = None
active_layout_key = KEY_LAYOUT_MAIN

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
                key=KEY_STATUS_BAR, 
                size=50, 
                pad=((15, 15), (10, 15)), 
                auto_size_text=True,
                font=('Verdana', 10, 'normal'))
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
                sg.Button('Back', key=KEY_BUTTON_BACK_GROUP_EDIT),
                ]],
                pad=((15, 15), (0, 7))),
        sg.Stretch()
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
                    app_data.save_data()
                break

    elif event == KEY_TREE_MAIN:
        should_disable = False if values[KEY_TREE_MAIN] else True
        window[KEY_BUTTON_OPEN_GROUP].update(disabled=should_disable)
        window[KEY_BUTTON_EDIT_GROUP].update(disabled=should_disable)
        window[KEY_BUTTON_DELETE_GROUP].update(disabled=should_disable)

    elif event == KEY_BUTTON_OPEN_GROUP:
        selected_group_id = values[KEY_TREE_MAIN][0]
        group_to_open = app_data.groups_data[selected_group_id]
        window[KEY_STATUS_BAR].update(value=f"({time.strftime('%H:%M:%S', time.localtime())}) Opening \"{group_to_open[KEY_GROUP_NAME]}\" group...")
        threading.Thread(target=async_group_opening, args=(group_to_open,)).start()
    
    elif event == EVENT_OPENING_COMPLETE:
        window[KEY_STATUS_BAR].update(value=f"({time.strftime('%H:%M:%S', time.localtime())}) Group \"{group_to_open[KEY_GROUP_NAME]}\" opened.")

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
        change_active_layout(KEY_LAYOUT_GROUP_EDIT)
    
    elif event == KEY_BUTTON_DELETE_GROUP:
        app_data.groups_data[:] = [group for group in app_data.groups_data if group[KEY_GROUP_ID] not in values[KEY_TREE_MAIN]]
        app_data.refresh_ids(for_groups=True)
        window[KEY_TREE_MAIN].update(create_groups_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=False)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=False)

    elif event == KEY_BUTTON_SAVE_CHANGES:
        app_data.save_data()
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
        window[KEY_STATUS_BAR].update(value=f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been successfully saved!")
    
    elif event == KEY_BUTTON_REVERT_CHANGES:
        app_data.revert_changes()
        window[KEY_TREE_MAIN].update(create_groups_tree_data())
        window[KEY_BUTTON_SAVE_CHANGES].update(disabled=True)
        window[KEY_BUTTON_REVERT_CHANGES].update(disabled=True)
        window[KEY_STATUS_BAR].update(value=f"({time.strftime('%H:%M:%S', time.localtime())}) Changes have been reverted!")

    elif event == KEY_BUTTON_HELP:
        change_active_layout(KEY_LAYOUT_HELP)
    
    elif event == KEY_HYPERLINK_HELP:
        os.startfile('https://github.com/ioan45/open-my-files')
    
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

    elif event == KEY_BUTTON_BACK_GROUP_EDIT:
        window[KEY_TREE_MAIN].update(key=group_to_edit[KEY_GROUP_ID], value=[len(group_to_edit[KEY_GROUP_ENTRIES])])
        group_to_edit = None
        change_active_layout(KEY_LAYOUT_MAIN)

window.close()
