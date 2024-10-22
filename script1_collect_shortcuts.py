# script1_collect_shortcuts.py

import os

def get_start_menu_shortcuts():
    start_menu_paths = [
        os.path.join(os.environ['PROGRAMDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
        os.path.join(os.environ['APPDATA'], 'Microsoft', 'Windows', 'Start Menu', 'Programs'),
    ]

    shortcut_entries = []

    for path in start_menu_paths:
        if os.path.exists(path):
            traverse_directory(path, shortcut_entries, base_path=path)

    return shortcut_entries

def traverse_directory(current_path, entries_list, base_path):
    # Get a sorted list of entries in the current directory
    entries = os.listdir(current_path)
    entries.sort()  # Sort alphabetically to match Start Menu display

    for entry in entries:
        full_path = os.path.join(current_path, entry)
        relative_path = os.path.relpath(full_path, base_path)

        if os.path.isdir(full_path):
            # Add folder entry (optional)
            entries_list.append(f"[Folder] {relative_path}")
            # Recursively traverse subdirectories
            traverse_directory(full_path, entries_list, base_path)
        elif os.path.isfile(full_path) and entry.lower().endswith('.lnk'):
            entries_list.append(full_path)

def save_shortcuts_to_file(shortcuts, file_path):
    with open(file_path, 'w', encoding='utf-8') as f:
        for shortcut in shortcuts:
            f.write(f"{shortcut}\n")

def main():
    print("Collecting Start Menu shortcut links...")

    output_file = os.path.join(os.getcwd(), 'StartMenuShortcuts.txt')

    shortcuts = get_start_menu_shortcuts()
    save_shortcuts_to_file(shortcuts, output_file)
    print(f"Shortcut links saved to {output_file}")

if __name__ == '__main__':
    main()
