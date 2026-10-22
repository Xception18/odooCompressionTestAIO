"""GUI editor for beban calculation configuration"""
import customtkinter as ctk
from tkinter import messagebox, filedialog
import json
from pathlib import Path
from typing import Optional
from modules.beban_config import get_beban_config, BebanConfig


class BebanConfigEditor(ctk.CTkToplevel):
    """GUI window for editing beban configuration"""

    # Drag-and-drop tuning
    DRAG_THRESHOLD_PX = 4      # movement needed before a press becomes a drag
    TICK_MS = 16               # drag loop interval (~60 fps)
    SCROLL_MARGIN_PX = 40      # edge zone that triggers auto-scroll
    SCROLL_MAX_SPEED_PX = 18   # max auto-scroll speed per tick

    def __init__(self, parent):
        super().__init__(parent)

        self.title("Beban Configuration Editor")
        self.geometry("900x700")

        self.config_manager = get_beban_config()
        self.current_mutu: Optional[str] = None

        # Drag and drop state
        self.drag_data = self._empty_drag_data()
        self.mutu_items = []    # Store mutu item widgets (in display order)
        self.drag_timer = None  # Drag loop timer id

        self.setup_ui()
        self.load_mutu_list()

    # ------------------------------------------------------------------
    # UI setup
    # ------------------------------------------------------------------
    def setup_ui(self):
        """Setup UI layout"""
        # Main container
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Left panel - Mutu list
        left_panel = ctk.CTkFrame(main_frame, width=200)
        left_panel.pack(side="left", fill="both", padx=(0, 10))
        left_panel.pack_propagate(False)

        title_frame = ctk.CTkFrame(left_panel)
        title_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(
            title_frame,
            text="Mutu List",
            font=("Arial", 14, "bold")
        ).pack(side="left", padx=10)

        # Drag tip label
        ctk.CTkLabel(
            title_frame,
            text="(drag to reorder)",
            font=("Arial", 9),
            text_color="#888888"
        ).pack(side="left")

        # Mutu listbox
        self.mutu_listbox = ctk.CTkScrollableFrame(left_panel)
        self.mutu_listbox.pack(fill="both", expand=True, pady=(0, 10))

        # Mutu buttons
        btn_frame = ctk.CTkFrame(left_panel)
        btn_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkButton(btn_frame, text="Add Mutu", command=self.add_mutu).pack(fill="x", pady=2)
        ctk.CTkButton(btn_frame, text="Delete Mutu", command=self.delete_mutu).pack(fill="x", pady=2)

        # Right panel - Config editor
        right_panel = ctk.CTkFrame(main_frame)
        right_panel.pack(side="right", fill="both", expand=True)

        # Title
        self.title_label = ctk.CTkLabel(right_panel, text="Select a Mutu",
                                        font=("Arial", 16, "bold"))
        self.title_label.pack(pady=10)

        # Scrollable config area
        self.config_frame = ctk.CTkScrollableFrame(right_panel)
        self.config_frame.pack(fill="both", expand=True, pady=(0, 10))

        # Bottom buttons
        bottom_frame = ctk.CTkFrame(main_frame)
        bottom_frame.pack(side="bottom", fill="x", pady=(10, 0))

        ctk.CTkButton(bottom_frame, text="Save", command=self.save_config,
                     fg_color="green", hover_color="darkgreen").pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="Apply", command=self.apply_config,
                     fg_color="blue", hover_color="darkblue").pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="Import", command=self.import_config).pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="Export", command=self.export_config).pack(side="left", padx=5)
        ctk.CTkButton(bottom_frame, text="Close", command=self.destroy).pack(side="right", padx=5)

    def load_mutu_list(self):
        """Load mutu names into listbox with drag-and-drop support"""
        self._cancel_drag_timer()
        self.drag_data = self._empty_drag_data()

        # Clear existing
        for widget in self.mutu_listbox.winfo_children():
            widget.destroy()

        self.mutu_items = []

        # Add mutu items with drag handles
        for mutu_name in self.config_manager.get_all_mutu_names():
            item = self.create_mutu_item(mutu_name)
            self.mutu_items.append(item)

    def create_mutu_item(self, mutu_name: str) -> ctk.CTkFrame:
        """Create a draggable mutu list item"""
        # Container frame
        item_frame = ctk.CTkFrame(self.mutu_listbox, height=40)
        item_frame.pack(fill="x", pady=2)
        item_frame.pack_propagate(False)

        # Drag handle (visual indicator + main grip)
        drag_handle = ctk.CTkLabel(
            item_frame,
            text="⋮⋮",
            width=28,
            font=ctk.CTkFont(size=14),
            cursor="fleur"
        )
        drag_handle.pack(side="left", padx=(5, 0), fill="y")

        # Mutu name button
        btn = ctk.CTkButton(
            item_frame,
            text=mutu_name,
            command=lambda m=mutu_name: self.select_mutu(m),
            anchor="w",
            height=30
        )
        btn.pack(side="left", fill="x", expand=True, padx=(5, 5))

        # Store mutu name in frame
        item_frame.mutu_name = mutu_name
        item_frame.original_color = item_frame.cget("fg_color")

        # Bind drag events. The frame itself is passed (not an index) so the
        # binding stays correct after items have been reordered.
        for widget in (drag_handle, item_frame):
            widget.bind("<ButtonPress-1>", lambda e, f=item_frame: self.on_drag_start(e, f))
            widget.bind("<B1-Motion>", self.on_drag_motion)
            widget.bind("<ButtonRelease-1>", self.on_drag_release)

        return item_frame

    # ------------------------------------------------------------------
    # Drag and drop
    # ------------------------------------------------------------------
    @staticmethod
    def _empty_drag_data() -> dict:
        return {
            "item": None,         # frame being dragged
            "active": False,      # True once the threshold has been passed
            "press_y": 0,         # screen Y at button press
            "grab_offset": 0,     # pointer offset inside the item at press
            "item_h": 0,          # real item height (px)
            "slot": 0,            # real distance between two rows (px)
            "top_offset": 0,      # first row top relative to list content top
            "index": -1,          # current index of the dragged item
            "start_index": -1,    # index at drag start (to detect changes)
        }

    def _get_canvas(self):
        """Return the canvas that scrolls the mutu list (or None)"""
        return getattr(self.mutu_listbox, "_parent_canvas", None)

    def _cancel_drag_timer(self):
        if self.drag_timer is not None:
            try:
                self.after_cancel(self.drag_timer)
            except Exception:
                pass
            self.drag_timer = None

    def on_drag_start(self, event, item):
        """Handle button press: remember geometry, drag starts after threshold"""
        if item not in self.mutu_items:
            return

        self._cancel_drag_timer()
        self.update_idletasks()  # make sure geometry is current before measuring

        index = self.mutu_items.index(item)
        first = self.mutu_items[0]

        # Measure real geometry (works with any DPI scaling)
        if len(self.mutu_items) > 1:
            slot = self.mutu_items[1].winfo_rooty() - first.winfo_rooty()
        else:
            slot = item.winfo_height()
        slot = max(slot, 1)

        self.drag_data = {
            "item": item,
            "active": False,
            "press_y": event.y_root,
            "grab_offset": event.y_root - item.winfo_rooty(),
            "item_h": item.winfo_height(),
            "slot": slot,
            "top_offset": first.winfo_rooty() - self.mutu_listbox.winfo_rooty(),
            "index": index,
            "start_index": index,
        }

    def on_drag_motion(self, event):
        """Activate the drag once the pointer moved past the threshold"""
        d = self.drag_data
        if d["item"] is None or d["active"]:
            return

        if abs(event.y_root - d["press_y"]) < self.DRAG_THRESHOLD_PX:
            return

        d["active"] = True
        d["item"].configure(fg_color="#4A90E2")  # highlight dragged item
        self._drag_tick()

    def _drag_tick(self):
        """Drag loop: auto-scroll + reposition, independent of mouse events"""
        d = self.drag_data
        if d["item"] is None or not d["active"]:
            return

        pointer_y = self.winfo_pointery()
        self._auto_scroll(pointer_y)
        self._update_drop_position(pointer_y)

        self.drag_timer = self.after(self.TICK_MS, self._drag_tick)

    def _auto_scroll(self, pointer_y: int):
        """Scroll the list when the pointer is near the top/bottom edge"""
        canvas = self._get_canvas()
        if canvas is None:
            return

        view_top = canvas.winfo_rooty()
        view_h = canvas.winfo_height()
        margin = self.SCROLL_MARGIN_PX

        if pointer_y < view_top + margin:
            ratio = min(1.0, (view_top + margin - pointer_y) / margin)
            step = -ratio * self.SCROLL_MAX_SPEED_PX
        elif pointer_y > view_top + view_h - margin:
            ratio = min(1.0, (pointer_y - (view_top + view_h - margin)) / margin)
            step = ratio * self.SCROLL_MAX_SPEED_PX
        else:
            return

        bbox = canvas.bbox("all")
        if not bbox:
            return
        content_h = bbox[3] - bbox[1]
        if content_h <= view_h:
            return  # nothing to scroll

        first = canvas.yview()[0]
        max_first = 1.0 - (view_h / content_h)
        new_first = min(max(first + step / content_h, 0.0), max_first)
        canvas.yview_moveto(new_first)

    def _target_index(self, pointer_y: int) -> int:
        """Slot index under the dragged item's centre (no hysteresis)"""
        d = self.drag_data
        list_top = self.mutu_listbox.winfo_rooty() + d["top_offset"]
        dragged_center = pointer_y - d["grab_offset"] + d["item_h"] / 2
        idx = int((dragged_center - list_top) // d["slot"])
        return max(0, min(idx, len(self.mutu_items) - 1))

    def _update_drop_position(self, pointer_y: int):
        d = self.drag_data
        new_index = self._target_index(pointer_y)
        if new_index != d["index"]:
            self.reorder_items(d["index"], new_index)
            d["index"] = new_index

    def on_drag_release(self, event):
        """Handle drag release"""
        d = self.drag_data
        if d["item"] is None:
            return

        self._cancel_drag_timer()

        item = d["item"]
        was_active = d["active"]
        changed = d["index"] != d["start_index"]

        if was_active:
            item.configure(fg_color=item.original_color)
            if changed:
                self.save_mutu_order()

        self.drag_data = self._empty_drag_data()

    def reorder_items(self, from_index: int, to_index: int):
        """Move one item to a new slot (only the dragged widget is re-packed)"""
        if from_index == to_index or from_index < 0 or to_index < 0:
            return
        if from_index >= len(self.mutu_items) or to_index >= len(self.mutu_items):
            return

        item = self.mutu_items.pop(from_index)
        self.mutu_items.insert(to_index, item)

        # Re-position just this widget instead of forgetting/re-packing all rows
        if to_index == len(self.mutu_items) - 1:
            item.pack_configure(fill="x", pady=2, after=self.mutu_items[to_index - 1])
        else:
            item.pack_configure(fill="x", pady=2, before=self.mutu_items[to_index + 1])

    def save_mutu_order(self):
        """Save the current mutu order to config"""
        # Get current order from UI
        new_order = [item.mutu_name for item in self.mutu_items]

        # Create new ordered config dict
        old_config = self.config_manager.config.copy()
        new_config = {}

        for mutu_name in new_order:
            if mutu_name in old_config:
                new_config[mutu_name] = old_config[mutu_name]

        # Replace config (in memory; user must click Save to persist)
        self.config_manager.config = new_config

    def destroy(self):
        """Make sure the drag loop stops before the window goes away"""
        self._cancel_drag_timer()
        super().destroy()

    # ------------------------------------------------------------------
    # Editor panel
    # ------------------------------------------------------------------
    def select_mutu(self, mutu_name: str):
        """Load mutu config into editor"""
        self.current_mutu = mutu_name
        self.title_label.configure(text=f"Edit: {mutu_name}")

        config = self.config_manager.get_mutu_config(mutu_name)
        if not config:
            return

        # Clear existing widgets
        for widget in self.config_frame.winfo_children():
            widget.destroy()

        self.entry_widgets = {}

        # Umur 7 section
        ctk.CTkLabel(self.config_frame, text="Umur 7 Hari",
                    font=("Arial", 14, "bold")).pack(anchor="w", pady=(10, 5))

        if "umur_7" in config:
            frame7 = ctk.CTkFrame(self.config_frame)
            frame7.pack(fill="x", pady=5)

            self.entry_widgets["umur_7_min"] = self.create_entry_row(
                frame7, "Min:", config["umur_7"].get("min", 0))
            self.entry_widgets["umur_7_max"] = self.create_entry_row(
                frame7, "Max:", config["umur_7"].get("max", 0))

        # Umur 28 section
        ctk.CTkLabel(self.config_frame, text="Umur 28 Hari",
                    font=("Arial", 14, "bold")).pack(anchor="w", pady=(20, 5))

        if "umur_28" in config:
            u28 = config["umur_28"]

            # Initial range
            ctk.CTkLabel(self.config_frame, text="Initial Range:",
                        font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 2))
            frame_init = ctk.CTkFrame(self.config_frame)
            frame_init.pack(fill="x", pady=5)
            self.entry_widgets["umur_28_initial_min"] = self.create_entry_row(
                frame_init, "Min:", u28.get("initial", {}).get("min", 0))
            self.entry_widgets["umur_28_initial_max"] = self.create_entry_row(
                frame_init, "Max:", u28.get("initial", {}).get("max", 0))

            # Mean
            frame_mean = ctk.CTkFrame(self.config_frame)
            frame_mean.pack(fill="x", pady=5)
            self.entry_widgets["umur_28_mean"] = self.create_entry_row(
                frame_mean, "Mean:", u28.get("mean", 0))

            # Same docket adjustments
            ctk.CTkLabel(self.config_frame, text="Same Docket Adjustment:",
                        font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 2))
            frame_same = ctk.CTkFrame(self.config_frame)
            frame_same.pack(fill="x", pady=5)
            self.entry_widgets["same_docket_up_min"] = self.create_entry_row(
                frame_same, "Up Min:", u28.get("same_docket_adjust_up", {}).get("min", 0))
            self.entry_widgets["same_docket_up_max"] = self.create_entry_row(
                frame_same, "Up Max:", u28.get("same_docket_adjust_up", {}).get("max", 0))
            self.entry_widgets["same_docket_down_min"] = self.create_entry_row(
                frame_same, "Down Min:", u28.get("same_docket_adjust_down", {}).get("min", 0))
            self.entry_widgets["same_docket_down_max"] = self.create_entry_row(
                frame_same, "Down Max:", u28.get("same_docket_adjust_down", {}).get("max", 0))

            # Different docket adjustments
            ctk.CTkLabel(self.config_frame, text="Different Docket Adjustment:",
                        font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 2))
            frame_diff = ctk.CTkFrame(self.config_frame)
            frame_diff.pack(fill="x", pady=5)
            self.entry_widgets["diff_docket_up_min"] = self.create_entry_row(
                frame_diff, "Up Min:", u28.get("diff_docket_adjust_up", {}).get("min", 0))
            self.entry_widgets["diff_docket_up_max"] = self.create_entry_row(
                frame_diff, "Up Max:", u28.get("diff_docket_adjust_up", {}).get("max", 0))
            self.entry_widgets["diff_docket_down_min"] = self.create_entry_row(
                frame_diff, "Down Min:", u28.get("diff_docket_adjust_down", {}).get("min", 0))
            self.entry_widgets["diff_docket_down_max"] = self.create_entry_row(
                frame_diff, "Down Max:", u28.get("diff_docket_adjust_down", {}).get("max", 0))

            # Range limits
            ctk.CTkLabel(self.config_frame, text="Final Range Limits:",
                        font=("Arial", 12, "bold")).pack(anchor="w", pady=(10, 2))
            frame_range = ctk.CTkFrame(self.config_frame)
            frame_range.pack(fill="x", pady=5)
            self.entry_widgets["range_min"] = self.create_entry_row(
                frame_range, "Min:", u28.get("range_min", 0))
            self.entry_widgets["range_max"] = self.create_entry_row(
                frame_range, "Max:", u28.get("range_max", 0))

    def create_entry_row(self, parent, label: str, value) -> ctk.CTkEntry:
        """Create label + entry row"""
        row = ctk.CTkFrame(parent)
        row.pack(fill="x", pady=2)

        ctk.CTkLabel(row, text=label, width=120).pack(side="left", padx=5)
        entry = ctk.CTkEntry(row)
        entry.insert(0, str(value))
        entry.pack(side="left", fill="x", expand=True, padx=5)

        return entry

    def add_mutu(self):
        """Add new mutu configuration"""
        dialog = ctk.CTkInputDialog(text="Enter Mutu Name:", title="Add Mutu")
        mutu_name = dialog.get_input()

        if not mutu_name:
            return

        if mutu_name in self.config_manager.get_all_mutu_names():
            messagebox.showerror("Error", "Mutu already exists!")
            return

        # Template config
        template = {
            "umur_7": {"min": 0.0, "max": 0.0},
            "umur_28": {
                "initial": {"min": 0.0, "max": 0.0},
                "mean": 0.0,
                "same_docket_adjust_up": {"min": 0.0, "max": 0.0},
                "same_docket_adjust_down": {"min": 0.0, "max": 0.0},
                "diff_docket_adjust_up": {"min": 0.0, "max": 0.0},
                "diff_docket_adjust_down": {"min": 0.0, "max": 0.0},
                "range_min": 0.0,
                "range_max": 0.0
            }
        }

        self.config_manager.add_mutu(mutu_name, template)
        self.load_mutu_list()
        self.select_mutu(mutu_name)

        # Show tip about dragging
        messagebox.showinfo("Tip",
            "You can drag the mutu items up or down to reorder them.\n\n"
            "Click on the drag handle (⋮⋮) or the item itself and drag.")

    def delete_mutu(self):
        """Delete current mutu"""
        if not self.current_mutu:
            messagebox.showwarning("Warning", "No mutu selected!")
            return

        if messagebox.askyesno("Confirm", f"Delete mutu '{self.current_mutu}'?"):
            self.config_manager.delete_mutu(self.current_mutu)
            self.current_mutu = None
            self.load_mutu_list()

            # Clear editor
            for widget in self.config_frame.winfo_children():
                widget.destroy()
            self.title_label.configure(text="Select a Mutu")

    def apply_config(self):
        """Apply current editor values to config (in memory only)"""
        if not self.current_mutu:
            messagebox.showwarning("Warning", "No mutu selected!")
            return

        try:
            config = {
                "umur_7": {
                    "min": float(self.entry_widgets["umur_7_min"].get()),
                    "max": float(self.entry_widgets["umur_7_max"].get())
                },
                "umur_28": {
                    "initial": {
                        "min": float(self.entry_widgets["umur_28_initial_min"].get()),
                        "max": float(self.entry_widgets["umur_28_initial_max"].get())
                    },
                    "mean": float(self.entry_widgets["umur_28_mean"].get()),
                    "same_docket_adjust_up": {
                        "min": float(self.entry_widgets["same_docket_up_min"].get()),
                        "max": float(self.entry_widgets["same_docket_up_max"].get())
                    },
                    "same_docket_adjust_down": {
                        "min": float(self.entry_widgets["same_docket_down_min"].get()),
                        "max": float(self.entry_widgets["same_docket_down_max"].get())
                    },
                    "diff_docket_adjust_up": {
                        "min": float(self.entry_widgets["diff_docket_up_min"].get()),
                        "max": float(self.entry_widgets["diff_docket_up_max"].get())
                    },
                    "diff_docket_adjust_down": {
                        "min": float(self.entry_widgets["diff_docket_down_min"].get()),
                        "max": float(self.entry_widgets["diff_docket_down_max"].get())
                    },
                    "range_min": float(self.entry_widgets["range_min"].get()),
                    "range_max": float(self.entry_widgets["range_max"].get())
                }
            }

            self.config_manager.add_mutu(self.current_mutu, config)
            messagebox.showinfo("Success", "Configuration applied in memory.\nClick 'Save' to persist to file.")

        except ValueError as e:
            messagebox.showerror("Error", f"Invalid number format: {e}")

    def save_config(self):
        """Save config to file"""
        if self.config_manager.save():
            messagebox.showinfo("Success",
                "Configuration saved to file!\n\n"
                "Note: The mutu order has been preserved.")
        else:
            messagebox.showerror("Error", "Failed to save configuration!")

    def import_config(self):
        """Import config from JSON file"""
        filepath = filedialog.askopenfilename(
            title="Import Config",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not filepath:
            return

        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                imported = json.load(f)

            # Validate structure
            if not isinstance(imported, dict):
                raise ValueError("Invalid config format")

            # Merge or replace?
            if messagebox.askyesno("Import Mode",
                                  "Replace existing config?\n\n" +
                                  "Yes = Replace all\n" +
                                  "No = Merge (keep existing)"):
                self.config_manager.config = imported
            else:
                self.config_manager.config.update(imported)

            self.load_mutu_list()
            messagebox.showinfo("Success", f"Imported {len(imported)} mutu configurations!")

        except Exception as e:
            messagebox.showerror("Error", f"Import failed: {e}")

    def export_config(self):
        """Export config to JSON file"""
        filepath = filedialog.asksaveasfilename(
            title="Export Config",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )

        if not filepath:
            return

        if self.config_manager.save(filepath):
            messagebox.showinfo("Success", f"Exported to:\n{filepath}")
        else:
            messagebox.showerror("Error", "Export failed!")


def open_config_editor(parent=None):
    """Open config editor window"""
    editor = BebanConfigEditor(parent)
    editor.grab_set()  # Modal
    editor.focus()