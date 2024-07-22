import json
import sys
import os
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import win32com.client

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

CONFIG_FILE = get_resource_path('config.json')
TF2_URL_SHORTCUT = get_resource_path('Team Fortress 2.url')
ICONS_FOLDER = get_resource_path('Icons')

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as file:
            return json.load(file)
    return {}

def save_config(config):
    with open(CONFIG_FILE, 'w') as file:
        json.dump(config, file)

def show_debug_message(message):
    messagebox.showinfo("Debug", message)

def create_lnk_from_url(url_path, lnk_path):
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(lnk_path)
    with open(url_path, 'r') as url_file:
        for line in url_file:
            if line.startswith("URL="):
                shortcut.TargetPath = line[4:].strip()
    shortcut.save()

def change_icon(tier, level):
    global ico_path, desktop_path

    # Define the path for the URL shortcut in the src folder
    tf2_url_shortcut = os.path.join(os.path.dirname(__file__), 'Team Fortress 2.url')

    tf2_lnk_shortcut = os.path.join(desktop_path, 'Team Fortress 2.lnk')
    
    if not os.path.exists(tf2_url_shortcut):
        show_debug_message("Missing the original TF2 shortcut.\nFIX:\n1. Open steam library\n2. Right click Team Fortress 2\n3. Hover over 'Manage' and select 'Add desktop shortcut'\n4. Place the newly created shortcut inside of Casual Badge Icon's folder\n5. Profit!")
        return
    
    if not os.path.exists(tf2_lnk_shortcut):
        create_lnk_from_url(tf2_url_shortcut, tf2_lnk_shortcut)
    
    ico_file = os.path.join(ico_path, f'{tier} {level}.ico')
    if not os.path.exists(ico_file):
        show_debug_message(f"ICO file {tier} {level}.ico does not exist yet. Check for updates on the Casual Badge Icon's github page!")
        return
    
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(tf2_lnk_shortcut)
    shortcut.IconLocation = ico_file
    shortcut.save()

def get_user_input():
    def on_submit():
        global ico_path, desktop_path
        try:
            tier = int(tier_entry.get())
            level = int(level_entry.get())

            # input validation
            if 1 <= tier <= 8 and 1 <= level <= 150:
                change_icon(tier, level)
                root.destroy()
            else:
                show_debug_message("Tier must be between 1 and 8, and Level must be between 1 and 150.")
        except ValueError:
            show_debug_message("Invalid input! Please enter integers for Tier and Level.")

    def select_icons_folder():
        global ico_path
        ico_path = filedialog.askdirectory(title="Select Icons Folder")
        if ico_path:
            save_config({"ico_path": ico_path, "desktop_path": desktop_path})
            icons_path_label.config(text=ico_path, foreground="green")

    def select_desktop_folder():
        global desktop_path
        desktop_path = filedialog.askdirectory(title="Select Desktop Folder")
        if desktop_path:
            save_config({"ico_path": ico_path, "desktop_path": desktop_path})
            desktop_path_label.config(text=desktop_path, foreground="green")

    # Load configuration
    config = load_config()
    global ico_path, desktop_path
    ico_path = config.get("ico_path", None)
    desktop_path = config.get("desktop_path", None)

    root = tk.Tk()
    root.title("Casual Badge Icon")
    root.configure(bg="#2e2e2e")

    # Create window
    window_width = 350
    window_height = 400
    root.geometry(f"{window_width}x{window_height}")
    root.resizable(False, False)  # Lock window size
    
    # Make frame on canvas
    container = tk.Frame(root, bg="#2e2e2e", padx=10, pady=10)
    container.place(relx=0.5, rely=0.5, anchor="center")

    # Title and subtitle
    title_label = ttk.Label(container, text="Casual Badge Icon", background="#2e2e2e", foreground="#ffffff", font=("Helvetica", 16, "bold"))
    title_label.pack(pady=(15, 0))
    subtitle_label = ttk.Label(container, text="made by greagob on github", background="#2e2e2e", foreground="#ffffff", font=("Helvetica", 10, "italic"))
    subtitle_label.pack(pady=(0, 10))

    desktop_button = ttk.Button(container, text="Select Desktop Folder", command=select_desktop_folder, style="TButton")
    desktop_button.pack(pady=5)
    global desktop_path_label
    if desktop_path:
        desktop_path_label = ttk.Label(container, text=desktop_path, foreground="green", font=("Helvetica", 10, "bold"), background="#2e2e2e")
    else:
        desktop_path_label = ttk.Label(container, text="No path selected", foreground="red", font=("Helvetica", 10, "bold"), background="#2e2e2e")
    desktop_path_label.pack(pady=5)

    icons_button = ttk.Button(container, text="Select Icons Folder", command=select_icons_folder, style="TButton")
    icons_button.pack(pady=5)
    global icons_path_label
    if ico_path:
        icons_path_label = ttk.Label(container, text=ico_path, foreground="green", font=("Helvetica", 10, "bold"), background="#2e2e2e")
    else:
        icons_path_label = ttk.Label(container, text="No path selected", foreground="red", font=("Helvetica", 10, "bold"), background="#2e2e2e")
    icons_path_label.pack(pady=5)

    ttk.Label(container, text="Enter Tier (1-8):", background="#2e2e2e", foreground="#f4a261", font=("Helvetica", 12, "bold")).pack(pady=5)
    tier_entry = ttk.Entry(container, style="TEntry")
    tier_entry.pack(pady=5)

    ttk.Label(container, text="Enter Level (1-150):", background="#2e2e2e", foreground="#f4a261", font=("Helvetica", 12, "bold")).pack(pady=5)
    level_entry = ttk.Entry(container, style="TEntry")
    level_entry.pack(pady=5)

    submit_button = ttk.Button(container, text="Submit", command=on_submit, style="TButton")
    submit_button.pack(pady=10)

    # Apply styles
    style = ttk.Style()
    style.theme_use('clam')
    style.configure("TEntry", background="#2e2e2e", foreground="white", fieldbackground="#2e2e2e", font=("Helvetica", 12))
    style.configure("TButton", background="#f4a261", foreground="black", font=("Helvetica", 12, "bold"), borderwidth=0)
    style.map("TButton", background=[("active", "#e76f51")])

    root.mainloop()

if __name__ == "__main__":
    get_user_input()