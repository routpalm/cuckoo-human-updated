# Copyright (C) 2012-2013 Claudio Guarnieri.
# Copyright (C) 2014-2018 Cuckoo Foundation.
# This file is part of Cuckoo Sandbox - http://www.cuckoosandbox.org
# See the file 'docs/LICENSE' for copying permission.

# Modified by Nicholas Anthony, 2021

import random
import re
import logging
import threading
from datetime import time

from pywinauto.application import Application
from pywinauto import Desktop, ElementNotFoundError, WindowNotFoundError, mouse
import pyautogui, subprocess, random, os

from lib.common.abstracts import Auxiliary
from lib.common.defines import (
    KERNEL32, USER32, WM_GETTEXT, WM_GETTEXTLENGTH, WM_CLOSE, BM_CLICK,
    EnumWindowsProc, EnumChildProc, create_unicode_buffer
)

# Cuckoo stuff
log = logging.getLogger(__name__)

# Cuckoo stuff
RESOLUTION = {
    "x": USER32.GetSystemMetrics(0),
    "y": USER32.GetSystemMetrics(1)
}


# Cuckoo Module
def click(hwnd):
    USER32.SetForegroundWindow(hwnd)
    KERNEL32.Sleep(1000)
    USER32.SendMessageW(hwnd, BM_CLICK, 0, 0)


# Cuckoo module
def foreach_child(hwnd, lparam):
    # List of partial buttons labels to click.
    buttons = [
        "yes", "oui",
        "ok",
        "i accept",
        "next", "suivant",
        "new", "nouveau",
        "install", "installer",
        "file", "fichier",
        "run", "start", "marrer", "cuter",
        "extract",
        "i agree", "accepte",
        "enable", "activer", "accord", "valider",
        "don't send", "ne pas envoyer",
        "don't save",
        "continue", "continuer",
        "personal", "personnel",
        "scan", "scanner",
        "unzip", "dezip",
        "open", "ouvrir",
        "close the program",
        "execute", "executer",
        "launch", "lancer",
        "save", "sauvegarder",
        "download", "load", "charger",
        "end", "fin", "terminer",
        "later",
        "finish",
        "end",
        "allow access",
        "remind me later",
        "save", "sauvegarder"
    ]

    # List of complete button texts to click. These take precedence.
    buttons_complete = [
        "&Ja",  # E.g., Dutch Office Word 2013.
    ]

    # List of buttons labels to not click.
    dontclick = [
        "don't run",
        "i do not accept"
    ]

    classname = create_unicode_buffer(50)
    USER32.GetClassNameW(hwnd, classname, 50)

    # Check if the class of the child is button.
    if "button" in classname.value.lower():
        # Get the text of the button.
        length = USER32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0)
        text = create_unicode_buffer(length + 1)
        USER32.SendMessageW(hwnd, WM_GETTEXT, length + 1, text)

        if text.value in buttons_complete:
            log.info("Found button %r, clicking it" % text.value)
            click(hwnd)
            return True

        # Check if the button is set as "clickable" and click it.
        textval = text.value.replace("&", "").lower()
        for button in buttons:
            if button in textval:
                for btn in dontclick:
                    if btn in textval:
                        break
                else:
                    log.info("Found button %r, clicking it" % text.value)
                    click(hwnd)

    # Recursively search for childs (USER32.EnumChildWindows).
    return True


# Cuckoo module
# Callback procedure invoked for every enumerated window.
# Purpose is to close any office window
def get_office_window(hwnd, lparam):
    if USER32.IsWindowVisible(hwnd):
        text = create_unicode_buffer(1024)
        USER32.GetWindowTextW(hwnd, text, 1024)
        # TODO Would " - Microsoft (Word|Excel|PowerPoint)$" be better?
        if re.search("- (Microsoft|Word|Excel|PowerPoint)", text.value):
            USER32.SendNotifyMessageW(hwnd, WM_CLOSE, None, None)
            log.info("Closed Office window.")
    return True


# Cuckoo method
def move_mouse():
    x = random.randint(0, RESOLUTION["x"])
    y = random.randint(0, RESOLUTION["y"])

    # Originally was:
    # USER32.mouse_event(0x8000, x, y, 0, None)
    # Changed to SetCurorPos, since using GetCursorPos would not detect
    # the mouse events. This actually moves the cursor around which might
    # cause some unintended activity on the desktop. We might want to make
    # this featur optional.
    USER32.SetCursorPos(x, y)


# Cuckoo method
def click_mouse():
    # Move mouse to top-middle position.
    USER32.SetCursorPos(RESOLUTION["x"] / 2, 0)
    # Mouse down.
    USER32.mouse_event(2, 0, 0, 0, None)
    KERNEL32.Sleep(50)
    # Mouse up.
    USER32.mouse_event(4, 0, 0, 0, None)


# Cuckoo method
# Callback procedure invoked for every enumerated window.
def foreach_window(hwnd, lparam):
    # If the window is visible, enumerate its child objects, looking
    # for buttons.
    if USER32.IsWindowVisible(hwnd):
        USER32.EnumChildWindows(hwnd, EnumChildProc(foreach_child), 0)
    return True


# -------- START MODIFICATIONS --------
#
#
#
#
#
#
# -------- GENERAL NOTES --------
# | Script function |
#
# This collection of methods utilizes pywinauto and pyautogui to simulate user interaction across multiple applications:
# Notepad, Adobe Acrobat PDF Reader, Microsoft Word 2007, Paint, VLC, Internet Explorer, & Calculator.
# In general, pywinauto is used for most operations since the goal is to directly interact with GUI elements.
# However, pyautogui is used often for moving and clicking the mouse as well as pressing hotkeys.
#
# | Recommendations for future development |
#
# TODO: Move some pywinauto calls to their own respective flexible methods so we can use always_wait_until
# TODO: General changes to make script less dependent on coordinate-based interactions
# TODO: Optimize CPU usage
# TODO: More error catching
#


# Check to see if process exists, if so we can connect to the existing session
# Conveniently calls Win32 APIs.
def process_exists(process_name):
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    output = subprocess.check_output(call).decode()
    # check in last line for process name
    last_line = output.strip().split('\r\n')[-1]
    # because Fail message could be translated
    return last_line.lower().startswith(process_name.lower())


# Open notepad, type some lines, save the file.
def notepad_interaction():
    # If process exists, connect to process - link Application to Notepad essentially
    if process_exists("notepad.exe"):
        print("Notepad session already exists. Connecting...")

        # We are using the UIA backend here, which is the cornerstone of modern pywinauto
        # and makes some things easier for developers.
        # Some Windows applications can be entered as an argument without entering the full path.
        app = Application(backend="uia").connect(title_re=".*Notepad*")

        # Here we will create a dialog instance based off of Notepad's foremost window.
        app_dialog = app.top_window()

        # Make sure that it has focus
        app_dialog.minimize()
        Desktop(backend="uia").window(title='Untitled - Notepad', visible_only=False).restore()
        print("Connected to existing Notepad session.")

    # Else, start a new Notepad process. This general framework will be the same for most methods.
    else:
        print("Notepad process does not exist. Creating a new one...")

        # For some reason, starting a new Notepad session did not work with ".*Notepad*." so here we are.
        app = Application(backend="uia").start(r"notepad.exe", timeout=20)
        print("Launched new Notepad session.")

        # This time, we'll create our dialog with more specific instructions since it's a new session.
        app_dialog = app.window(title_re='Untitled - Notepad', visible_only=False)
        app_dialog.minimize()
        Desktop(backend="uia").window(title_re='Untitled - Notepad', visible_only=False).restore()

    # defining main dialog
    dlg = app.UntitledNotepad

    # This is used a lot throughout the course of the code, so to clarify:
    # wait_cpu_usage_lower() will force the script to wait until its cpu usage is below a certain amount.
    # This is especially good for VMs because they more often than not will have less than optimal processing power.
    app.wait_cpu_usage_lower(threshold=20)

    # type in the box with .1 second delay (Edit is the specific TextArea that we are interacting with)
    dlg.Edit.type_keys("Hello! This program is typing.\n\n Lorem ipsum dolor sit amet latin latin latin",
                       with_spaces=True,
                       with_newlines=True,
                       pause=.1,
                       with_tabs=True)
    n = 0

    # Scroll
    while n < 30:
        pyautogui.press('enter')
        n += 1
    pyautogui.scroll(1000)

    # Since pywinauto supports it, we can use an easier way of navigating to the Save As window (in Notepad only)
    dlg['File'].select()
    submenu = app['']
    submenu['Save As'].click_input()
    save_as = dlg.child_window(title_re="Save As", class_name="#32770")

    # We will save this file as TestFile.txt on the Desktop.
    # Originally, it was to be saved here so we could click on it later with pyautogui but alas, no dice with that.
    save_as.FileNameCombo.type_keys(os.environ['USERPROFILE'] + "\Desktop\TestFile.txt")
    save_as.Save.click()

    dlg.close()


# WORK IN PROGRESS
# Opens Acrobat and creates a new PDF.
def acrobat_interaction():
    # Start new process - link Application to Acrobat
    if process_exists("AcroRD32.exe"):
        print("Acrobat session already exists. Connecting...")
        app = Application(backend="uia").connect(
            path=r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRD32.exe")
        app.wait_cpu_usage_lower(threshold=20)
        app_dialog = app.top_window()
        app_dialog.minimize()
        Desktop(backend="uia").window(title='Adobe Acrobat Reader DC (32-bit)', visible_only=False).restore()
        print("Connected to existing Acrobat session.")

    else:
        print("Acrobat process does not exist. Creating a new one...")
        app = Application(backend="uia").start(r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRD32.exe",
                                               timeout=20)
        time.sleep(2)
        app.wait_cpu_usage_lower(threshold=20)
        app.connect(title='Adobe Acrobat Reader DC (32-bit)')
        print("Launched new Acrobat session.")

    app.wait_cpu_usage_lower(threshold=20)
    adobe = app.window(class_name='AcrobatSDIWindow')

    # Navigate to "Select Files to Convert to PDF" page
    # we will use click() and set_edit_text() when possible because it allows for
    # better performance when running multiple analysis machines
    app_menu = adobe.child_window(title="Application", control_type="MenuBar")
    app_menu.child_window(title="File").expand()
    file_menu = adobe.child_window(title="File", control_type="Menu", found_index=0)
    file_menu.child_window(title="Create PDF").click_input()
    app.wait_cpu_usage_lower(threshold=20)

    # Adobe doesn't list button as a control identifier, so we have to use pyautogui
    pyautogui.moveTo(453, 372, 2)
    pyautogui.click()
    time.sleep(2)

    app.wait_cpu_usage_lower(threshold=30)

    # Select notepad file
    file_dlg = app.AdobeAcrobatReaderDC.child_window(title_re="Select Files")
    file_dlg.FileNameEdit.set_edit_text(os.environ['USERPROFILE'] + '\Desktop\TestFile.txt')
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.click(x=360, y=436)

    time.sleep(4)

    # This next section is done mostly in pyautogui, because the pywinauto control hierarchy was an absolute mess.
    # Future work on this program should include more pywinauto code because coordinate based control is shaky at best.
    pyautogui.click(x=616, y=19)
    pyautogui.click(x=414, y=227)  # email field
    pyautogui.write('')  # insert email here
    pyautogui.click(x=414, y=271)  # pw field
    pyautogui.write('')  # insert password here
    pyautogui.click(x=356, y=335)  # sign in
    pyautogui.click(x=571, y=208)  # open pdf

    # on first open, adobe creates a window that blocks opening pdf.
    # we'll click on it and then let it navigate to ie, then refocus the window
    app_dialog.minimize()
    app_dialog.restore()
    pyautogui.click(x=571,
                    y=208)  # on first open, adobe creates a window that blocks opening pdf. we'll click on it and then let it navigate to ie, then refocus the window

    app_dialog.close()


# Open word, navigate through the setup, type some lines, scroll, save the file.
def word_interaction():
    if process_exists("WINWORD.exe"):
        print("Word session already exists, connecting...")
        app = Application(backend="uia").connect(path=r"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.exe",
                                                 timeout=20)
        print("Connected")
        app.wait_cpu_usage_lower(threshold=30)
        app_dialog = app.top_window()
        app_dialog.minimize()
        Desktop(backend="uia").window(title="Document 1 - Microsoft Word non-commercial use",
                                      visible_only=False).restore()
    else:
        print("Word session does not exist. Starting a new one...")
        app = Application(backend="uia").start(r"C:\Program Files (x86)\Microsoft Office\Office12\WINWORD.exe",
                                               timeout=20)
        app.wait_cpu_usage_lower(threshold=30)
        app_dialog = app.top_window()
        app_dialog.minimize()
        Desktop(backend="uia").window(title="Document 1 - Microsoft Word non-commercial use",
                                      visible_only=False).restore()

        print("Connected.")

    # If this is the first run of Office, it will generate a setup wizard which we can ignore by pressing the cancel button
    try:
        setup_dlg = app.Document1MicrosoftWord.child_window(title_re="Microsoft Office Activation Wizard",
                                                            found_index=0)
        setup_dlg.child_window(title="Cancel", control_type="Button").click_input()
    except ElementNotFoundError:
        pass
    except WindowNotFoundError:
        pass

    # Here we can write!
    app_dialog.type_keys("This is a test.",
                         with_spaces=True,
                         with_newlines=True,
                         pause=.1,
                         with_tabs=True)
    pyautogui.press('enter')
    app_dialog.type_keys("PyAutoGui is extremely buggy.",
                         with_spaces=True,
                         with_newlines=True,
                         pause=.1,
                         with_tabs=True)
    pyautogui.press('enter')
    pyautogui.press('tab')
    app_dialog.type_keys("Anyways, this is the third line. Bye!",
                         with_spaces=True,
                         with_newlines=True,
                         pause=.1,
                         with_tabs=True)
    pyautogui.doubleClick(x=397, y=466)
    app_dialog.type_keys("SIKE! I'm still typing but I double clicked beforehand",
                         with_spaces=True,
                         with_newlines=True,
                         pause=.1,
                         with_tabs=True)
    # Pressing enter a bunch of times to simulate going to a new page and also for scrolling purposes
    n = 0
    while n < 30:
        pyautogui.press('enter')
        n += 1

    # Much like the pywinauto scroll module, the pyautogui one similarly works at random. Swap to a different libraryrecommended.
    pyautogui.scroll(1000)
    pyautogui.scroll(-1000)
    pyautogui.scroll(1000)

    # In order: Select all, copy, paste, save
    pyautogui.hotkey('ctrl', 'a')
    pyautogui.hotkey('ctrl', 'c')
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.hotkey('ctrl', 's')

    app.wait_cpu_usage_lower(threshold=15)

    # Navigate through the Save As dialog
    save_dlg = app.Document1MicrosoftWord.child_window(title_re="Save As", found_index=0)
    save_dlg.child_window(title="Save", control_type="Button").click_input()

    app_dialog.close()


# Open Calculator, switch to scientific view, do 7 random operations, toggle history.
def calculator_interaction():
    if process_exists("calc.exe"):
        print("Calculator session already exists, connecting...")
        app = Application(backend="uia").connect(path=r"C:\Windows\System32\calc.exe", timeout=20)
        print("Connected.")
        app.wait_cpu_usage_lower(threshold=16)
        app_dialog = app.top_window()
        app_dialog.minimize()
        Desktop(backend="uia").window(title='Calculator', visible_only=False).restore()
    else:
        print("Calculator session does not exist. Starting a new one...")
        app = Application(backend="uia").start(r"C:\Windows\System32\calc.exe", timeout=20)
        app.wait_cpu_usage_lower(threshold=16)
        app_dialog = app.top_window()
        print("Connected.")

    # We will have the program execute a series of mathematical problems.
    operations_list = ['+', '-', '*', '/']
    n = 0
    pyautogui.hotkey('alt', '2')
    while n < 14:
        operation = random.choice(operations_list)
        rnum1 = random.randint(1, 99)
        rnum2 = random.randint(1, 99)
        app_dialog.type_keys("" + str(rnum1) + operation + str(rnum2))
        pyautogui.press('enter')
        time.sleep(2)
        n += 1

    # Show history
    pyautogui.hotkey('ctrl', 'h')

    # Clicking just for the hell of it
    pyautogui.doubleClick()

    app_dialog.close()


# Open paint, open koala.jpg, change image attributes, save and exit
def paint_interaction():
    print("MS Paint session does not exist. Starting...")
    app = Application(backend="uia").start(r"C:\Windows\System32\mspaint.exe", timeout=20)
    app_dialog = app.window(title_re='.* - Paint', visible_only=False)
    app_dialog.minimize()
    Desktop(backend="uia").window(title_re='.* - Paint', visible_only=False).restore()
    print("Connected.")

    app.wait_cpu_usage_lower(threshold=20)

    # Connecting to the Paint window, and navigating to the Open MenuItem/dialog
    dlg = app.window(title_re='.* - Paint')
    dlg.Applicationmenu.click_input()
    dlg.child_window(title='Open', control_type='MenuItem', found_index=0).invoke()
    file_dlg = app.UntitledPaint.child_window(title_re="Open", found_index=0)
    file_dlg.FileNameEdit.set_edit_text('Koala.jpg')
    pyautogui.press('enter')

    app.wait_cpu_usage_lower(threshold=20)

    # Changing image properties to 350x350
    pyautogui.hotkey('ctrl', 'e')
    attribute_dlg = app.KoalaPaint.child_window(title_re="Image Properties")
    attribute_dlg.child_window(title="Width:", auto_id="264", control_type="Edit").set_edit_text("350")
    attribute_dlg.child_window(title="Height:", auto_id="266", control_type="Edit").set_edit_text("350")
    attribute_dlg.child_window(title="OK", auto_id="1", control_type="Button").click_input()
    pyautogui.hotkey('ctrl', 's')

    app_dialog.close()


# Open Internet Explorer,
def ie_interaction():
    # We are directing this application to start by connecting to Google
    app = Application(backend="uia").start(
        r"C:\Program Files (x86)\Internet Explorer\iexplore.exe {}".format("https://google.com"), timeout=100)
    app.wait_cpu_usage_lower(threshold=35)
    ie_dialog = app.window(title_re="Google - Windows Internet Explorer")
    ie_dialog.minimize()
    Desktop(backend="uia").window(title_re="Google - Windows Internet Explorer", visible_only=False).restore()
    time.sleep(4)
    ie_dialog.set_focus()

    # Pywinauto cannot detect elements of an HTML page, so we have to use coordinate based interaction here.
    # Future development of pywinauto/some genius developer may change this fact.
    pyautogui.click(x=305, y=334)
    ie_dialog.type_keys("cat videos", with_spaces=True)
    pyautogui.press('enter')
    app.wait_cpu_usage_lower(threshold=35)
    newpage_dialog = app.window(title_re="cat videos - Google Search - Windows Internet Explorer")

    # Following pywinauto docs, this should scroll the mouse down the search page.
    # However, being inconsistent in testing, future versions should include a different library for scrolling
    ie_rect = newpage_dialog.rectangle()
    coords = (random.randint(ie_rect.left, ie_rect.right), random.randint(ie_rect.top, ie_rect.bottom))
    mouse.scroll(coords=coords, wheel_dist=-100)

    newpage_dialog.close()


# Open VLC, open a video from the Sample Videos folder, play the video
def vlc_interaction():
    if process_exists("vlc.exe"):
        print("VLC Media Player session already exists, connecting...")
        app = Application(backend="uia").connect(path=r"C:\Program Files\VideoLAN\VLC\vlc.exe", timeout=20)
        app.wait_cpu_usage_lower(threshold=25)
        app_dialog = app.top_window()
        app_dialog.minimize()
        app_dialog.restore()
        print("Connected.")
    else:
        print("VLC Media Player session does not exist. Starting a new one...")
        app = Application(backend="uia").start(r"C:\Program Files\VideoLAN\VLC\vlc.exe", timeout=20)
        app.wait_cpu_usage_lower(threshold=25)
        app_dialog = app.top_window()
        print("Connected.")

    # On the first run, privacy dialog will appear. Since (ideally) the vm will be unmodified/unopened applications,
    # we will assume it's there
    try:
        privacy_dlg = app.VLCMediaPlayer.child_window(title_re="Privacy and Network Access Policy", found_index=0)
        privacy_dlg.print_control_identifiers()
        # invoke because sometimes the window is generated with the bottom cut off
        privacy_dlg.child_window(title="Continue Enter", control_type="Button", found_index=0).invoke()
    except Exception as e:
        print("Not first run. Privacy dialog does not exist.")

    # Open dialog hotkey
    pyautogui.hotkey('ctrl', 'o')

    # Swap focus to Open dialog
    # VLC opens the default video folder which contains wmv files, we can open this
    open_dlg = app_dialog.child_window(title_re="Select one or more files to open")
    app.wait_cpu_usage_lower(threshold=10)  # Again, if VM was not allotted enough cores, may be script-killing
    open_dlg.FileNameEdit.set_edit_text("C:\Users\Public\Videos\Sample Videos\Wildlife.wmv")
    pyautogui.press('enter')  # Load video
    time.sleep(3)
    pyautogui.doubleClick(x=300, y=300)  # Another coord-based input, this should press the Play button
    time.sleep(40)
    app_dialog.close()


# Half cuckoo method, half my method
class Human(threading.Thread, Auxiliary):
    """Human after all"""

    def __init__(self, options={}, analyzer=None):
        threading.Thread.__init__(self)
        Auxiliary.__init__(self, options, analyzer)
        self.do_run = True

    def stop(self):
        self.do_run = False

    def run(self):
        seconds = 0

        # Global disable flag.
        if "human" in self.options:
            self.do_move_mouse = int(self.options["human"])
            self.do_click_mouse = int(self.options["human"])
            self.do_click_buttons = int(self.options["human"])
            self.do_notepad_interaction = False
            self.do_paint_interaction = False
            self.do_acrobat_interaction = False  # WIP
            self.do_word_interaction = False
            self.do_ie_interaction = False
            self.do_calculator_interaction = False
        else:
            # We want to disable the move and click mouse Cuckoo operations, because they will disrupt the script
            self.do_move_mouse = False
            self.do_click_mouse = False
            self.do_click_buttons = False # not sure about this one, it might still be able to stay
            self.do_notepad_interaction = True
            self.do_paint_interaction = True
            self.do_acrobat_interaction = True  # WIP
            self.do_word_interaction = True
            self.do_ie_interaction = True
            self.do_calculator_interaction = True

        # Per-feature enable or disable flag.
        if "human.move_mouse" in self.options:
            self.do_move_mouse = int(self.options["human.move_mouse"])

        if "human.click_mouse" in self.options:
            self.do_click_mouse = int(self.options["human.click_mouse"])

        if "human.click_buttons" in self.options:
            self.do_click_buttons = int(self.options["human.click_buttons"])

        while self.do_run:
            if seconds and not seconds % 60:
                USER32.EnumWindows(EnumWindowsProc(get_office_window), 0)

            if self.do_click_mouse:
                click_mouse()

            if self.do_move_mouse:
                move_mouse()

            if self.do_click_buttons:
                USER32.EnumWindows(EnumWindowsProc(foreach_window), 0)

            if self.do_notepad_interaction:
                notepad_interaction()
                self.do_notepad_interaction = False

            if self.do_paint_interaction:
                paint_interaction()
                self.do_paint_interaction = False

            if self.do_word_interaction:
                word_interaction()
                self.do_word_interaction = False

            if self.do_adobe_interaction:
                acrobat_interaction()
                self.do_acrobat_interaction = False

            if self.do_ie_interaction:
                ie_interaction()
                self.do_ie_interaction = False

            if self.do_calculator_interaction:
                calculator_interaction()
                self.do_calculator_interaction = False

            KERNEL32.Sleep(1000)
            seconds += 1
