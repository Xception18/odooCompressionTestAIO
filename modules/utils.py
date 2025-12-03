import os
import sys
import logging
import queue
import tkinter as tk

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class ThreadSafeLogHandler(logging.Handler):
    """Thread-safe logging handler that queues messages for GUI"""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        log_entry = self.format(record)
        # Put message in queue instead of direct GUI access
        self.log_queue.put(log_entry)

class LogHandler(logging.Handler):
    def __init__(self, log_widget=None):
        super().__init__()
        self.log_widget = log_widget

    def emit(self, record):
        log_entry = self.format(record)
        if self.log_widget:
            try:
                # If we have a GUI widget, add to it
                self.log_widget.insert('end', log_entry + '\n')
                self.log_widget.see('end')
                # Update the GUI
                self.log_widget.update()
            except Exception as e:
                # If GUI is not available, print to console
                print(f"GUI Log Error: {e}")
                print(log_entry)
        else:
            # Fallback to console
            print(log_entry)
