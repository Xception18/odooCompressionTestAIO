"""Toast notification component"""
import customtkinter as ctk
from modules.ui.theme import color, font, RADIUS, SPACING
import threading

class Toast(ctk.CTkFrame):
    """
    Toast notification that appears at bottom-right and auto-dismisses.
    
    Usage:
        Toast.show(root, "File uploaded successfully", level="success")
    """
    
    @staticmethod
    def show(root, message: str, level: str = "info", duration_ms: int = 4000):
        """
        Show a toast notification.
        
        Args:
            root: Root window
            message: Toast message
            level: info/success/warning/error
            duration_ms: Auto-dismiss time (0 = manual dismiss only)
        """
        toast = Toast(root, message, level)
        toast.place(relx=1.0, rely=1.0, anchor="se", x=-20, y=-20)
        
        if duration_ms > 0:
            def dismiss():
                try:
                    toast.dismiss()
                except:
                    pass  # Widget may be destroyed
            
            threading.Timer(duration_ms / 1000, lambda: root.after(0, dismiss)).start()
        
        return toast
    
    def __init__(self, master, message: str, level: str = "info"):
        """
        Args:
            master: Parent widget
            message: Toast text
            level: info/success/warning/error
        """
        # Color based on level
        level_colors = {
            "info": color("primary"),
            "success": color("success"),
            "warning": color("warning"),
            "error": color("danger")
        }
        
        fg = level_colors.get(level, level_colors["info"])
        
        super().__init__(
            master,
            fg_color=fg,
            corner_radius=RADIUS["input"],
            border_width=0
        )
        
        # Content container
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(padx=SPACING["lg"], pady=SPACING["md"])
        
        # Message
        label = ctk.CTkLabel(
            content,
            text=message,
            font=font("body"),
            text_color=("#FFFFFF", "#FFFFFF"),  # Always white on colored bg
            wraplength=300
        )
        label.pack(side="left")
        
        # Close button
        close_btn = ctk.CTkButton(
            content,
            text="×",
            width=20,
            height=20,
            corner_radius=10,
            fg_color="transparent",
            hover_color=("white", "white"),
            text_color=("#FFFFFF", "#FFFFFF"),
            command=self.dismiss
        )
        close_btn.pack(side="left", padx=(SPACING["md"], 0))
    
    def dismiss(self):
        """Dismiss the toast with fade animation"""
        # ponytail: fade animation using alpha/opacity
        # CTk doesn't support widget transparency easily, so just destroy
        self.destroy()


class ToastQueue:
    """
    Queue multiple toasts vertically.
    
    Usage:
        queue = ToastQueue(root)
        queue.info("Processing...")
        queue.success("Done!")
    """
    
    def __init__(self, root, spacing: int = 10):
        """
        Args:
            root: Root window
            spacing: Vertical spacing between toasts
        """
        self.root = root
        self.spacing = spacing
        self.toasts = []
    
    def _reposition(self):
        """Reposition all toasts vertically"""
        y_offset = 20
        for toast in self.toasts:
            if toast.winfo_exists():
                toast.place(relx=1.0, rely=1.0, anchor="se", x=-20, y=-y_offset)
                y_offset += toast.winfo_height() + self.spacing
    
    def show(self, message: str, level: str = "info", duration_ms: int = 4000):
        """Show toast in queue"""
        toast = Toast(self.root, message, level)
        self.toasts.append(toast)
        self._reposition()
        
        if duration_ms > 0:
            def dismiss():
                try:
                    toast.destroy()
                    self.toasts.remove(toast)
                    self._reposition()
                except:
                    pass
            
            threading.Timer(duration_ms / 1000, 
                          lambda: self.root.after(0, dismiss)).start()
    
    def info(self, message: str, duration_ms: int = 4000):
        """Show info toast"""
        self.show(message, "info", duration_ms)
    
    def success(self, message: str, duration_ms: int = 4000):
        """Show success toast"""
        self.show(message, "success", duration_ms)
    
    def warning(self, message: str, duration_ms: int = 4000):
        """Show warning toast"""
        self.show(message, "warning", duration_ms)
    
    def error(self, message: str, duration_ms: int = 4000):
        """Show error toast"""
        self.show(message, "error", duration_ms)
