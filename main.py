import tkinter as tk
import logging
import sys
import os

# Add the current directory to sys.path to ensure modules can be imported
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from modules.ui.main_window import ExcelProcessorGUI

def main():
    root = tk.Tk()
    # Set icon if available
    try:
        # root.iconbitmap("icon.ico") 
        pass
    except:
        pass
        
    app = ExcelProcessorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()