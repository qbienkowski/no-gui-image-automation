# script2_collect_executables.py

import os
import sys
import win32com.client
import tkinter as tk
from tkinter import filedialog

def load_shortcuts_from_file(file_path):
    if not os.path.exists(file_path):
        print(f"Shortcuts file not found at {file_path}.")
        return None

    with open(file_path, 'r', encoding='utf-8') as f:
        shortcuts = [line.strip() for line in f]
    return shortcuts

def get_shortcut_target(shortcut_path):
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(shortcut_path)
        return shortcut.Targetpath
    except Exception as e:
        print(f"Error resolving shortcut target for {shortcut_path}: {e}")
        return ''

def save_executables_to_file(executables, file_path):
    with open(file_path, 'w', encoding='utf-8') as f:
        for exe in executables:
            f.write(f"{exe}\n")

def main():
    # Prompt the user to select a .txt file
    root = tk.Tk()
    root.withdraw()
    shortcuts_file = filedialog.askopenfilename(title="Select Shortcuts File", filetypes=(("Text files", "*.txt"),))
    if not shortcuts_file:
        print("No shortcuts file selected.")
        sys.exit(1)

    print("Resolving executables from shortcut links...")
    shortcuts = load_shortcuts_from_file(shortcuts_file)
    if shortcuts is None:
        sys.exit(1)

    executables = []
    for shortcut in shortcuts:
        if shortcut.startswith("[Folder]"):
            # It's a folder entry; preserve it
            executables.append(shortcut)
        elif shortcut:
            exe_path = get_shortcut_target(shortcut)
            executables.append(exe_path)

    # Save the executables to a .txt file in the same order
    executables_file = os.path.join(os.getcwd(), 'ExecutablePaths.txt')
    save_executables_to_file(executables, executables_file)
    print(f"Executable paths saved to {executables_file}")

if __name__ == '__main__':
    main()
