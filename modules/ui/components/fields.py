"""Form field components"""
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
from typing import Literal, Callable
from modules.ui.theme import color, font, RADIUS, SIZES, SPACING

class LabeledEntry(ctk.CTkFrame):
    """Entry with label and optional helper text"""
    
    def __init__(self, master, label: str, variable: tk.StringVar = None,
                 helper: str = None, width: int = None, **entry_kwargs):
        """
        Args:
            master: Parent widget
            label: Label text
            variable: StringVar to bind
            helper: Optional helper text below entry
            width: Frame width (None = expand)
            **entry_kwargs: Additional CTkEntry arguments
        """
        super().__init__(master, fg_color="transparent")
        
        if width:
            self.configure(width=width)
        
        # Label
        label_widget = ctk.CTkLabel(
            self,
            text=label,
            font=font("body"),
            anchor="w"
        )
        label_widget.pack(fill="x", pady=(0, SPACING["xs"]))
        
        # Entry
        entry_kwargs.setdefault("height", SIZES["input_height"])
        entry_kwargs.setdefault("corner_radius", RADIUS["input"])
        entry_kwargs.setdefault("font", font("body"))
        
        self.entry = ctk.CTkEntry(self, textvariable=variable, **entry_kwargs)
        self.entry.pack(fill="x")
        
        # Helper text
        if helper:
            helper_widget = ctk.CTkLabel(
                self,
                text=helper,
                font=font("helper"),
                text_color=color("muted"),
                anchor="w"
            )
            helper_widget.pack(fill="x", pady=(SPACING["xs"], 0))


class PasswordEntry(ctk.CTkFrame):
    """Entry with show/hide toggle for passwords"""
    
    def __init__(self, master, label: str, variable: tk.StringVar = None, **kwargs):
        """
        Args:
            master: Parent widget
            label: Label text
            variable: StringVar to bind
        """
        super().__init__(master, fg_color="transparent")
        
        # Label
        label_widget = ctk.CTkLabel(
            self,
            text=label,
            font=font("body"),
            anchor="w"
        )
        label_widget.pack(fill="x", pady=(0, SPACING["xs"]))
        
        # Entry + toggle button container
        entry_frame = ctk.CTkFrame(self, fg_color="transparent")
        entry_frame.pack(fill="x")
        
        # Entry
        self.entry = ctk.CTkEntry(
            entry_frame,
            textvariable=variable,
            show="•",
            height=SIZES["input_height"],
            corner_radius=RADIUS["input"],
            font=font("body")
        )
        self.entry.pack(side="left", fill="x", expand=True)
        
        # Toggle button
        self._visible = False
        self.toggle_btn = ctk.CTkButton(
            entry_frame,
            text="👁",
            width=40,
            height=SIZES["input_height"],
            corner_radius=RADIUS["input"],
            fg_color="transparent",
            hover_color=color("surface_alt"),
            command=self._toggle_visibility
        )
        self.toggle_btn.pack(side="left", padx=(SPACING["xs"], 0))
    
    def _toggle_visibility(self):
        """Toggle password visibility"""
        self._visible = not self._visible
        self.entry.configure(show="" if self._visible else "•")


class FileField(ctk.CTkFrame):
    """
    File/directory picker field with browse button and validation.
    
    Replaces create_file_input and create_directory_input.
    """
    
    def __init__(self, master, label: str, variable: tk.StringVar,
                 mode: Literal["file", "dir"] = "file",
                 filetypes: tuple = (("Excel files", "*.xlsx *.xls"), ("All files", "*.*")),
                 on_change: Callable = None, **kwargs):
        """
        Args:
            master: Parent widget
            label: Field label
            variable: StringVar for path
            mode: "file" or "dir"
            filetypes: File type filter for file mode
            on_change: Callback when path changes
        """
        super().__init__(master, fg_color="transparent")
        
        self.variable = variable
        self.mode = mode
        self.filetypes = filetypes
        self.on_change = on_change
        
        # Label
        label_widget = ctk.CTkLabel(
            self,
            text=label,
            font=font("body"),
            anchor="w"
        )
        label_widget.pack(fill="x", pady=(0, SPACING["xs"]))
        
        # Entry + buttons container
        entry_frame = ctk.CTkFrame(self, fg_color="transparent")
        entry_frame.pack(fill="x")
        
        # Entry
        self.entry = ctk.CTkEntry(
            entry_frame,
            textvariable=variable,
            height=SIZES["input_height"],
            corner_radius=RADIUS["input"],
            font=font("body")
        )
        self.entry.pack(side="left", fill="x", expand=True)
        
        # Validation indicator
        self.status_label = ctk.CTkLabel(
            entry_frame,
            text="",
            width=20,
            font=font("body")
        )
        self.status_label.pack(side="left", padx=(SPACING["xs"], 0))
        
        # Browse button
        self.browse_btn = ctk.CTkButton(
            entry_frame,
            text="📁",
            width=40,
            height=SIZES["input_height"],
            corner_radius=RADIUS["input"],
            fg_color=color("primary"),
            hover_color=color("primary_hover"),
            command=self._browse
        )
        self.browse_btn.pack(side="left", padx=(SPACING["xs"], 0))
        
        # Bind validation
        variable.trace_add("write", lambda *_: self._validate())
        self._validate()
    
    def _browse(self):
        """Open file/directory picker"""
        if self.mode == "file":
            path = filedialog.askopenfilename(filetypes=self.filetypes)
        else:
            path = filedialog.askdirectory()
        
        if path:
            self.variable.set(path)
            if self.on_change:
                self.on_change(path)
    
    def _validate(self):
        """Validate path and update indicator"""
        path_str = self.variable.get()
        
        if not path_str:
            self.status_label.configure(text="")
            self.entry.configure(border_color=color("border"))
            return
        
        path = Path(path_str)
        exists = path.exists()
        
        if exists:
            self.status_label.configure(text="✓", text_color=color("success"))
            self.entry.configure(border_color=color("border"))
        else:
            self.status_label.configure(text="✗", text_color=color("danger"))
            self.entry.configure(border_color=color("danger"))
        
        if self.on_change and exists:
            self.on_change(path_str)
