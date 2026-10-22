"""Modal dialog components - replaces tkinter.messagebox"""
import customtkinter as ctk
from modules.ui.theme import color, font, RADIUS, SPACING
from typing import Literal, Callable

class BaseDialog(ctk.CTkToplevel):
    """Base modal dialog"""
    
    def __init__(self, parent, title: str, width: int = 400, height: int = None):
        super().__init__(parent)
        
        self.title(title)
        self.result = None
        
        # Modal behavior
        self.transient(parent)
        self.grab_set()
        
        # Center on parent
        if height:
            self.geometry(f"{width}x{height}")
        else:
            # Auto-height
            self.geometry(f"{width}x100")  # Temporary, will resize
        
        # Center position
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - width) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")
        
        # Handle close button
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)
    
    def on_cancel(self):
        """Handle cancel/close"""
        self.result = None
        self.destroy()
    
    def wait_result(self):
        """Wait for dialog result (blocking)"""
        self.wait_window()
        return self.result


class InfoDialog(BaseDialog):
    """
    Information dialog - replaces messagebox.showinfo.
    
    Usage:
        InfoDialog(root, "Success", "File uploaded successfully")
    """
    
    def __init__(self, parent, title: str, message: str):
        super().__init__(parent, title, width=450)
        
        # Content frame
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=SPACING["xl"], pady=SPACING["xl"])
        
        # Message
        msg_label = ctk.CTkLabel(
            content,
            text=message,
            font=font("body"),
            wraplength=400,
            justify="left"
        )
        msg_label.pack(pady=(0, SPACING["lg"]))
        
        # OK button
        from modules.ui.components.buttons import PrimaryButton
        ok_btn = PrimaryButton(content, text="OK", command=self.on_ok, width=100)
        ok_btn.pack()
        
        # Bind Enter key
        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_ok())
    
    def on_ok(self):
        """Handle OK"""
        self.result = True
        self.destroy()


class ErrorDialog(BaseDialog):
    """
    Error dialog with copyable details - replaces messagebox.showerror.
    
    Usage:
        ErrorDialog(root, "Error", "Database connection failed", 
                   detail="psycopg2.OperationalError: connection refused")
    """
    
    def __init__(self, parent, title: str, message: str, detail: str = None):
        super().__init__(parent, title, width=500)
        
        # Content frame
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=SPACING["xl"], pady=SPACING["xl"])
        
        # Error icon + message
        msg_frame = ctk.CTkFrame(content, fg_color="transparent")
        msg_frame.pack(fill="x", pady=(0, SPACING["md"]))
        
        icon = ctk.CTkLabel(msg_frame, text="⚠", font=("Arial", 32), 
                           text_color=color("danger"))
        icon.pack(side="left", padx=(0, SPACING["md"]))
        
        msg_label = ctk.CTkLabel(
            msg_frame,
            text=message,
            font=font("body"),
            wraplength=400,
            justify="left"
        )
        msg_label.pack(side="left", fill="x", expand=True)
        
        # Detail (if provided)
        if detail:
            detail_frame = ctk.CTkFrame(content, fg_color=color("surface_alt"),
                                       corner_radius=RADIUS["input"])
            detail_frame.pack(fill="both", expand=True, pady=(0, SPACING["md"]))
            
            detail_text = ctk.CTkTextbox(
                detail_frame,
                height=100,
                font=("Cascadia Mono", 11),
                wrap="word"
            )
            detail_text.pack(fill="both", expand=True, padx=SPACING["sm"], 
                           pady=SPACING["sm"])
            detail_text.insert("1.0", detail)
            detail_text.configure(state="disabled")
        
        # OK button
        from modules.ui.components.buttons import PrimaryButton
        ok_btn = PrimaryButton(content, text="OK", command=self.on_ok, width=100)
        ok_btn.pack()
        
        # Bind keys
        self.bind("<Return>", lambda e: self.on_ok())
        self.bind("<Escape>", lambda e: self.on_ok())
    
    def on_ok(self):
        """Handle OK"""
        self.result = True
        self.destroy()


class ConfirmDialog(BaseDialog):
    """
    Confirmation dialog - replaces messagebox.askyesno / askokcancel.
    
    Usage:
        if ConfirmDialog(root, "Quit", "Stop running process and quit?").wait_result():
            # User clicked Yes
    """
    
    def __init__(self, parent, title: str, message: str, 
                 yes_text: str = "Yes", no_text: str = "No",
                 danger: bool = False):
        """
        Args:
            parent: Parent window
            title: Dialog title
            message: Confirmation message
            yes_text: Positive button text
            no_text: Negative button text
            danger: If True, Yes button is red (destructive action)
        """
        super().__init__(parent, title, width=450)
        
        # Content frame
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=SPACING["xl"], pady=SPACING["xl"])
        
        # Message
        msg_label = ctk.CTkLabel(
            content,
            text=message,
            font=font("body"),
            wraplength=400,
            justify="left"
        )
        msg_label.pack(pady=(0, SPACING["lg"]))
        
        # Buttons
        btn_frame = ctk.CTkFrame(content, fg_color="transparent")
        btn_frame.pack()
        
        from modules.ui.components.buttons import SecondaryButton, PrimaryButton, DangerButton
        
        # No button
        no_btn = SecondaryButton(btn_frame, text=no_text, command=self.on_no, 
                                width=100)
        no_btn.pack(side="left", padx=(0, SPACING["sm"]))
        
        # Yes button
        if danger:
            yes_btn = DangerButton(btn_frame, text=yes_text, command=self.on_yes, 
                                  width=100)
        else:
            yes_btn = PrimaryButton(btn_frame, text=yes_text, command=self.on_yes, 
                                   width=100)
        yes_btn.pack(side="left")
        
        # Bind keys
        self.bind("<Return>", lambda e: self.on_yes())
        self.bind("<Escape>", lambda e: self.on_no())
    
    def on_yes(self):
        """Handle Yes"""
        self.result = True
        self.destroy()
    
    def on_no(self):
        """Handle No"""
        self.result = False
        self.destroy()


# Convenience functions (messagebox API compatibility)

def showinfo(parent, title: str, message: str):
    """Show info dialog - messagebox.showinfo replacement"""
    InfoDialog(parent, title, message)

def showerror(parent, title: str, message: str, detail: str = None):
    """Show error dialog - messagebox.showerror replacement"""
    ErrorDialog(parent, title, message, detail)

def askyesno(parent, title: str, message: str) -> bool:
    """Ask yes/no - messagebox.askyesno replacement"""
    return ConfirmDialog(parent, title, message).wait_result() or False

def askokcancel(parent, title: str, message: str) -> bool:
    """Ask OK/Cancel - messagebox.askokcancel replacement"""
    return ConfirmDialog(parent, title, message, 
                        yes_text="OK", no_text="Cancel").wait_result() or False
