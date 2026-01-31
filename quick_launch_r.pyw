"""
SPDX-License-Identifier: MIT
"""

APP_NAME = "Quick Launch R"
APP_REV = "1.0"
APP_AUTHOR = "mifrey00"
APP_LIC = "MIT"
APP_DESCR= "A system tray application for quick access to shortcuts, restoring the functionnality of the Quick Launch toolbar that was removed in Windows 11."
APP_URL = "https://github.com/mifrey00/Quick-Launch-R"

import os
import sys
from pathlib import Path
import pystray
from pystray import MenuItem as Item
from PIL import Image, ImageDraw
import threading
import tkinter as tk
from tkinter import messagebox
from textwrap import dedent

SCRIPT_DIR = Path(__file__).parent

class ShortcutLauncher:
    def __init__(self, shortcuts_path):
        self.shortcuts_path = Path(shortcuts_path)
        self.icon = None
        
    def create_icon_image(self):
        """Load icon image from file"""
        icon_path = SCRIPT_DIR / "icon.png"
        
        if icon_path.exists():
            try:
                return Image.open(icon_path)
            except Exception as e:
                print(f"Error loading icon: {e}, using default icon")
        
        # Fallback to simple generated icon
        width = 64
        height = 64
        image = Image.new('RGB', (width, height), 'blue')
        dc = ImageDraw.Draw(image)
        dc.rectangle([width // 4, height // 4, width * 3 // 4, height * 3 // 4], fill='white')
        return image
    
    def execute_lnk(self, lnk_path):
        """Execute a .lnk file using Windows shell"""
        try:
            os.startfile(lnk_path)
        except Exception as e:
            print(f"Error executing {lnk_path}: {e}")
    
    def make_launch_callback(self, lnk_path):
        """Create a callback function for launching a specific shortcut"""
        def launch(icon, item):
            self.execute_lnk(lnk_path)
        return launch
    
    def is_folder_shortcut(self, lnk_path):
        """Check if a .lnk file points to a folder"""
        try:
            import pythoncom
            import win32com.client
            
            pythoncom.CoInitialize()
            try:
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(str(lnk_path))
                target_path = shortcut.Targetpath
                
                # Store result before cleanup
                is_folder = False
                if target_path and os.path.exists(target_path):
                    is_folder = os.path.isdir(target_path)
                
                # Explicitly release COM objects
                shortcut = None
                shell = None
                
                return is_folder
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Error checking shortcut target: {e}")
        
        return False
    
    def build_menu_from_folder(self, folder_path):
        """Recursively build menu items from folder structure"""
        menu_items = []
        
        if not folder_path.exists():
            return menu_items
        
        # Get all items in the folder, sorted
        items = sorted(folder_path.iterdir(), key=lambda x: (not x.is_dir(), x.name.lower()))
        
        for item in items:
            if item.is_dir():
                # Create submenu for folders
                submenu_items = self.build_menu_from_folder(item)
                if submenu_items:
                    menu_items.append(Item(
                        "📁 " + item.name,
                        pystray.Menu(*submenu_items)
                    ))
            elif item.suffix.lower() == '.lnk':
                # Create menu item for .lnk files
                display_name = item.stem  # Remove .lnk extension
                lnk_path = str(item.resolve())
                
                # Check if shortcut points to a folder
                if self.is_folder_shortcut(item):
                    display_name = "📁 " + display_name
                else:
                    display_name = "↗️ " + display_name
                
                menu_items.append(Item(
                    display_name,
                    self.make_launch_callback(lnk_path)
                ))
        
        return menu_items
    
    def open_shortcuts_folder(self):
        """Open the shortcuts folder in Windows Explorer"""
        try:
            os.startfile(self.shortcuts_path)
        except Exception as e:
            print(f"Error opening folder: {e}")
    
    def show_about(self):
        """Show about dialog"""
        def show_dialog():
            try:                
                root = tk.Tk()
                root.withdraw()

                about_text = dedent(f"""\
                    {APP_NAME}(estored)
                    
                    Author: {APP_AUTHOR}
                    Version: {APP_REV}
                    License: {APP_LIC}
                    
                    {APP_DESCR}
                    
                    {APP_URL}""")
                
                tk.messagebox.showinfo(f"About {APP_NAME}", about_text)
                root.destroy()
            except Exception as e:
                print(f"Error showing about dialog: {e}")
        
        # Run dialog in separate thread
        thread = threading.Thread(target=show_dialog, daemon=True)
        thread.start()
    
    def create_menu(self):
        """Create the main menu"""
        menu_items = self.build_menu_from_folder(self.shortcuts_path)
        
        # Add separator and utility options
        menu_items.append(pystray.Menu.SEPARATOR)
        menu_items.append(Item('Shortcuts Folder', lambda: self.open_shortcuts_folder()))
        menu_items.append(Item('About', lambda: self.show_about()))
        menu_items.append(Item('Exit', self.stop))
        
        return pystray.Menu(*menu_items)
    
    def stop(self):
        """Stop the icon"""
        if self.icon:
            self.icon.stop()
    
    def run(self):
        """Run the system tray icon"""
        image = self.create_icon_image()
        menu = self.create_menu()
        
        self.icon = pystray.Icon(
            "shortcut_launcher",
            image,
            f"{APP_NAME}",
            menu
        )
        
        self.icon.run()

def main():
    # Path to the shortcuts folder
    shortcuts_folder = SCRIPT_DIR / "Shortcuts"
    
    # Create and run the launcher
    launcher = ShortcutLauncher(shortcuts_folder)
    launcher.run()

if __name__ == "__main__":
    main()