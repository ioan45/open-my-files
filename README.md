# Open My Files

This application helps Windows users to open their usual "setup" faster. Here, setup means groups of applications, documents and other files in the system, and even web pages. 

If you are like me, and by that I mean if you like to keep your desktop clean and organize your things who knows where in the system, then this application may prove to be useful for you.  

## Third Party

This application uses:

* [PySimpleGUI](https://github.com/PySimpleGUI/PySimpleGUI)
* [Bootstrap Icons](https://github.com/twbs/icons)
* [watchdog](https://github.com/gorakhargosh/watchdog)

## Installation

Here is how you can bring this application to life in your system:

1. Get the contents of this repository into your system. You can do it by cloning the repository using the [Git](https://git-scm.com/download/win) command `git clone https://github.com/ioan45/open-my-files.git` or, if you don't have Git installed, you can press the green "Code" button (which is at the top of this web page) and then the "Download ZIP" button to get an archive which should be extracted after the download completes.
2. Download and install [Python](https://www.python.org/downloads/windows/). This project was made using the 3.11.4 version, but if you already have another version then try using the one you have.
3. After installing Python, open a console window and run the command `pip install -r requirements.txt` in the root directory (the directory which contains the `requirements.txt` file).

## Usage

To launch the application you have to run the command `python omf.py` in the directory which contains the `omf.py` file. Make sure you have Python in your PATH system variable, otherwise you would have to specify the full path in the command.

In short, you have to create groups (just some containers) and add to them the paths to your files and/or web addresses, then you can open a group anytime just by selecting it in the list and pressing a button.

The application is structured as follows:

### Main interface

The main interface contains the list of groups, checkboxes for setting the application to start with Windows and for enabling the auto saving feature, and a status bar for application messages. It also contains some buttons:
* **Open Group** - Opens all entries added to the selected group. Each non-executable file will be opened using the default application for its type which is specified in the Windows settings. When it comes to web pages, those will be accessed using the default web browser which is set in the Windows settings.
* **New Group** - Creates an empty group.
* **Edit Group** - Shows the editing interface for the selected group.
* **Delete Group** - Removes the selected group from the list.
* **Save Changes** & **Revert Changes** - Saves/Reverts any changes you've made in the application.
* **Help** - Opens a little interface from where you can access this repository.
* **Exit** - Closes the application. It will pop up a window for saving changes (if any).

<p align="center"><img src="/res/main_interface.PNG?raw=true" width=70% height=70%/></p>

### Group editing interface

This is the interface shown on pressing the **Edit Group** button mentioned above. Any changes made here can be saved using the **Save Changes** button from the main interface. This interface shows the list of entries added to the group as well as a few buttons:
* **Add Files** - Adds files to the list by opening a window for you to select them from your system.
* **Add Web Page** - It shows an input field where you can specify the URL of the web page that you want to add.
* **Delete Selected Entries** - Removes the selected entries from the list.
* **Edit Selected Details** - Each list entry has a details field which can be used for extra information about the entry. This button lets you edit the field for the selected entries.
* **Start/Stop Listening** - You can make the group listen to changes made in a directory. Any file adding/deleting/renaming operation will be reflected in the group. Still, you have to save any changes occurred in the group. 
* **Back** - Shows the main interface.

<p align="center"><img src="/res/group_editing_interface.PNG?raw=true" align="center" width=100% height=100%/></p>
