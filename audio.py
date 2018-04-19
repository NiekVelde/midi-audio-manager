# Imports
import sys, os, time
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
from ctypes import cast, POINTER, wintypes
import pygame, pygame.midi
from pygame.locals import *
import tkinter, win32api, win32con, pywintypes
from multiprocessing import Pool
import datetime
import math
import ctypes

# Global variables and constants
ignored_audio_sessions = ["ShellExperienceHost.exe"]
excluded_sycle_programs = ["Master volume", "Microphone"]
text_duration = 1

# Label variables
label_visible = False
label_created = datetime.datetime.now()
label = None
label_timeout = 2

# Define controls and there rows
volume_sliders = (0, 1, 2, 3, 4, 5, 6, 7, 8)
sycle_buttons = (32, 33, 34, 35, 36, 37, 38, 39)
mute_buttons = (48, 49, 50, 51, 52, 53, 54, 55)
show_buttons = (64, 65, 66, 67, 68, 69, 70, 71)
knobs = (16, 17, 18, 19, 20, 21, 22, 23)
row_programs = ["Master volume", None, None, None, None, None, None, None]
play_pause_button = 41
next_track_button = 44
prev_track_button = 43

# Define user 32
user32 = ctypes.WinDLL('user32', use_last_error=True)
wintypes.ULONG_PTR = wintypes.WPARAM
INPUT_MOUSE    = 0
INPUT_KEYBOARD = 1
INPUT_HARDWARE = 2
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP       = 0x0002
KEYEVENTF_UNICODE     = 0x0004
KEYEVENTF_SCANCODE    = 0x0008
MAPVK_VK_TO_VSC = 0

# Define main function
def main():
    # Load audio programs
    load_audio_programs()

    # Setup label
    setup_label()

    # Setup midi input listener
    setup_midi_listener()

# Load saved audio programs
def load_audio_programs():
    if os.path.isfile("audioPrograms.data"):
        # Open file
        file = open("audioPrograms.data", "r")

        # Get programs array
        programs = file.readline().rstrip()
        programs = programs.split(",")

        # Loop trough array and set programs
        index = 1
        for program in programs:
            row_programs[index] = program
            index += 1

        # Close file
        file.close()

# Save current audio programs to save file
def save_audio_programs():
    # Open/Create file
    file = open("audioPrograms.data", "w")

    # Loop trough row programs
    program_string = ""
    for program in row_programs:
        if program != "Master volume":
            program_string += program + ","

    # Write string to file
    file.write(program_string[:-1])

    # Close file
    file.close()

# Midi event trigger
def trigger_midi_event(event):
    # Define midi event variables
    event_key = event.data1
    event_value = event.data2

    # Get right function from midi event
    if event_key in volume_sliders:
        change_program_volume(volume_sliders.index(event_key), event_value)
    elif event_key in sycle_buttons and event_value is not 0:
        sycle_programs(sycle_buttons.index(event_key))
    elif event_key in mute_buttons:
        # Check if mute needs to be added or removed
        if event_value is 0:
            unmute_program_volume(mute_buttons.index(event_key))
        else:
            mute_program_volume(mute_buttons.index(event_key))
    elif event_key in show_buttons and event_value is not 0:
        show_current_program(show_buttons.index(event_key))
    elif event_key == play_pause_button:
        press_key(0xb3)
    elif event_key == next_track_button:
        press_key(0xb0)
    elif event_key == prev_track_button:
        press_key(0xb1)
    else:
        print("No event handler found for button: %s" % event_key)

# Change program volume
def change_program_volume(row, value):
    # Get program
    program = get_row_program(row)

    # Check program
    if program not in excluded_sycle_programs:
        # Get sessions
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            # Check process name
            if session.Process and session.Process.name() == "%s.exe" % program:
                # Define program volume
                volume = session.SimpleAudioVolume
                # Mute
                volume.SetMasterVolume(get_volume_percentage(value), None)
    elif program == "Master volume":
        # Get master volume
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))

            # Set volume
            volume.SetMasterVolumeLevel(get_master_volume_value(value), None)
        except:
            pass

# Mute program volume
def mute_program_volume(row):
    toggle_mute_on_program_volume(row, 1)

# Unmute program volume
def unmute_program_volume(row):
    toggle_mute_on_program_volume(row, 0)

def toggle_mute_on_program_volume(row, value):
    # Get program
    program = get_row_program(row)

    # Check program
    if program not in excluded_sycle_programs:
        # Get sessions
        sessions = AudioUtilities.GetAllSessions()
        for session in sessions:
            # Check process name
            if session.Process and session.Process.name() == "%s.exe" % program:
                # Define program volume
                volume = session.SimpleAudioVolume
                # Mute
                volume.SetMute(value, None)
    elif program == "Master volume":
        # Get master volume
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))

            # Set volume
            volume.SetMute(value, None)
        except:
            pass

# Sycle between programs
def sycle_programs(row):
    # Define globals
    global row_programs

    # Get helper variables
    program = get_row_program(row)
    sessions = get_current_audio_sessions()

    # Check program
    if program not in excluded_sycle_programs:
        # Check current row
        if program is None or program not in sessions:
            index = 0
        else:
            index = get_current_audio_sessions().index(program)

        # Check sessions length
        if len(sessions) - 1 >= index + 1:
            index = index + 1
        else:
            index = 0

        # Set new program
        row_programs[row] = sessions[index]
        # Show new program
        show_label("New program: %s" % sessions[index])

        # Save new data file
        save_audio_programs()
    else:
        show_label("Can't sycle for program: %s" % program)

# Show current program
def show_current_program(row):
    show_label("Current: %s" % get_row_program(row))

# Get volume percentage
def get_volume_percentage(value):
    return (value / 127)

# Get master volume value
def get_master_volume_value(value):
    return (-64 / 100) * (100 - (get_volume_percentage(value) * 100))

# Get program by row
def get_row_program(row):
    return row_programs[row]

# Get current audio sessions
def get_current_audio_sessions():
    # Define return list
    found_sessions = []

    # Get sessions
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        # Check process
        if session.Process and session.Process.name() not in ignored_audio_sessions:
            # Create readable process name
            session_name = session.Process.name().replace(".exe", "")
            # Check if not already in found sessions
            if session_name not in found_sessions:
                found_sessions.append(session_name)

    # Return filled list
    return found_sessions

# Setup midi input listener
def setup_midi_listener():
    # Setup pygame
    pygame.init()
    pygame.fastevent.init()
    event_get = pygame.fastevent.get
    event_post = pygame.fastevent.post
    pygame.display.set_mode((200, 1))
    pygame.display.iconify()
    # Start midi
    pygame.midi.init()

    # Get midi input
    midi_input = pygame.midi.Input(pygame.midi.get_default_input_id())
    listening = True

    # Listening loop
    while listening:
        # Check for new midi events
        if midi_input.poll():
            midi_events = midi_input.read(10)
            # convert them into pygame events.
            midi_events = pygame.midi.midis2events(midi_events, midi_input.device_id)

            for midi_event in midi_events:
                event_post(midi_event)

        # Get events
        events = event_get()
        for event in events:
            # On quit
            if event.type in [QUIT]:
                listening = False
            # Midi event
            if event.type in [pygame.midi.MIDIIN]:
                trigger_midi_event(event)

        # Check for label visibility
        if label_visible and (datetime.datetime.now() - label_created).total_seconds() >= label_timeout:
            # Hide label
            hide_label()

        # Make sure this loop wont just steal all of the systems power,
        # cause he will if you let him
        time.sleep(0.02)

    # Delete midi input
    del midi_input
    # End midi
    pygame.midi.quit()

# Setup label
def setup_label():
    global label
    # Crate text label
    label = tkinter.Label(font=("Times New Roman","20"), fg="black", bg="white")
    label.master.overrideredirect(True)
    label.master.geometry("+0+-1000")
    label.master.wm_attributes("-topmost", True)
    label.master.wm_attributes("-disabled", True)

    hWindow = pywintypes.HANDLE(int(label.master.frame(), 16))
    # http://msdn.microsoft.com/en-us/library/windows/desktop/ff700543(v=vs.85).aspx
    # The WS_EX_TRANSPARENT flag makes events (like mouse clicks) fall through the window.
    exStyle = win32con.WS_EX_COMPOSITED | win32con.WS_EX_LAYERED | win32con.WS_EX_NOACTIVATE | win32con.WS_EX_TOPMOST | win32con.WS_EX_TRANSPARENT
    win32api.SetWindowLong(hWindow, win32con.GWL_EXSTYLE, exStyle)

    # Update screen
    label.pack()
    label.update()

def callback():
    print("callback")

# Show label with text
def show_label(text):
    # Define globals
    global label
    global label_visible
    global label_created

    # Set text
    label.config(text=text)
    # Set label placement in screen
    label.master.geometry("+0+0")
    # Update screen
    label.update()

    # Set label vars
    label_visible = True
    label_created = datetime.datetime.now()

# Hide label
def hide_label():
    # Define globals
    global label_visible
    global label

    # Set label placement of screen
    label.master.geometry("+0+-1000")
    # Update screen
    label.update()

    # Set visibility
    label_visible = False

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (("dx",          wintypes.LONG),
                ("dy",          wintypes.LONG),
                ("mouseData",   wintypes.DWORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (("wVk",         wintypes.WORD),
                ("wScan",       wintypes.WORD),
                ("dwFlags",     wintypes.DWORD),
                ("time",        wintypes.DWORD),
                ("dwExtraInfo", wintypes.ULONG_PTR))

    def __init__(self, *args, **kwds):
        super(KEYBDINPUT, self).__init__(*args, **kwds)
        # some programs use the scan code even if KEYEVENTF_SCANCODE
        # isn't set in dwFflags, so attempt to map the correct code.
        if not self.dwFlags & KEYEVENTF_UNICODE:
            self.wScan = user32.MapVirtualKeyExW(self.wVk,
                                                 MAPVK_VK_TO_VSC, 0)

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = (("uMsg",    wintypes.DWORD),
                ("wParamL", wintypes.WORD),
                ("wParamH", wintypes.WORD))

class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = (("ki", KEYBDINPUT),
                    ("mi", MOUSEINPUT),
                    ("hi", HARDWAREINPUT))
    _anonymous_ = ("_input",)
    _fields_ = (("type",   wintypes.DWORD),
                ("_input", _INPUT))

# Press keyboard key
def press_key(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

# Release keyboard key
def release_key(hexKeyCode):
    x = INPUT(type=INPUT_KEYBOARD,
              ki=KEYBDINPUT(wVk=hexKeyCode,
                            dwFlags=KEYEVENTF_KEYUP))
    user32.SendInput(1, ctypes.byref(x), ctypes.sizeof(x))

# Call main init function
main()
