# Migration Plan: Tkinter → CustomTkinter (Modern UI Redesign)

**Project:** Excel Data Processor – ODOO Test Integration
**Source file analysed:** `modules/ui/main_window.py` (`ExcelProcessorGUI`)
**Goal:** Finish the move to CustomTkinter, remove the remaining legacy Tk/ttk widgets, and give the app a modern, consistent look without breaking the existing workflow (Excel → BJDT ID → CSV → DB → ODOO sync, plus Selenium "Rencana Benda Uji").

---

## 1. Current State Audit

The code is **already partly on CustomTkinter**. Most widgets are `ctk.*`. The migration is therefore a *completion + redesign*, not a rewrite from scratch.

### 1.1 What is already CTk
`CTkLabel`, `CTkButton`, `CTkEntry`, `CTkFrame`, `CTkScrollableFrame`, `CTkTabview`, `CTkFont`.

### 1.2 What is still legacy Tk / ttk

| Legacy usage | Where | Problem | Target |
|---|---|---|---|
| `ttk.Progressbar` (`style="Custom.Horizontal.TProgressbar"`) | Process tab | Does not follow theme; the style name is never defined, so it falls back to the OS look. The inline comment claims CTk has no indeterminate mode, but `CTkProgressBar(mode="indeterminate")` supports it | `ctk.CTkProgressBar` |
| `ScrolledText` | Log tab | Hard-coded colours, native scrollbar, no log-level colouring | `ctk.CTkTextbox` + tag colours |
| `tkinter.messagebox` (≈ 15 call sites) | Everywhere, **including worker threads** | Native OS dialogs break the look and are not thread-safe | In-app modal (`CTkToplevel`) + toast notifications |
| `tk.StringVar` | All form fields | Fine – CTk uses the same variables | **Keep** |
| `filedialog` | Browse buttons | Native dialog is acceptable | **Keep** |
| `root.configure(bg='#212121')` and hex colours scattered in code | Many places | No central palette, no light mode | Theme JSON + `theme.py` tokens |
| Emoji as icons in button/tab text | Everywhere | Renders differently per OS, inconsistent baseline | `CTkImage` icons (Lucide/Phosphor PNG) |
| `CTkTabview` with 5 emoji tabs | Main navigation | Cramped for 5+ sections, dated look | Sidebar navigation |

### 1.3 Structural issues that affect the redesign
- One 1,000-line class mixes **UI, Excel logic, HTTP calls, DB upload and thread control**. UI work is hard to isolate and test.
- `create_file_input(..., row)` and `create_directory_input(..., row)` take a `row` argument they never use, and are near-duplicates.
- Button enable/disable is hand-written per button (`disable_buttons` / `enable_buttons`), and the Grid and Beban buttons are missing from it.
- Colours for the "toggle" buttons (Start/Stop) are duplicated as literals in 6+ places.

---

## 2. Bugs to Fix During the Migration

These were found while reading the code. They matter for the UI work because the new notification/dialog layer must be called from the main thread only.

1. **Race condition in `full_process`.** Step 3 calls `self.generate_csv_files()`, which only *starts a new thread* and returns immediately. Step 4 (`upload_to_database`) then runs before the CSV exists. Call `_generate_csv_files_logic()` directly inside the worker.
2. **Tk calls from worker threads.** `messagebox.*`, `progress.start()/stop()` and `sync_btn.configure()` are called from `threading.Thread` targets (`full_process`, `copy_data_process`, `bjdt_id_process`, `upload_db_process`, `disable_buttons`, `enable_buttons`). Route everything through `root.after(0, ...)` via a single dispatcher (see §6.3).
3. **False success message.** `upload_db_process` shows "uploaded successfully" in `finally`, even when the upload raised or failed with a duplicate-key error. Return a result object and report real status.
4. **UI-blocking stop.** `stop_rencana_process` calls `thread.join(timeout=1.0)` on the main thread. Move the wait to a worker and update the UI via callback.
5. **Hard-coded DB password** (`'adhimix'`) and credentials built into a connection URL string. Load defaults from `config.cnf` or environment variables, and build the URL with `sqlalchemy.engine.URL.create(...)` so special characters in passwords do not break it.
6. **`process_log_queue()` is started before `setup_ui()`**, so early logs are silently dropped when `log_text` does not exist yet. Buffer messages until the log widget exists.
7. **Unbounded log widget.** The log text grows forever. Cap it (e.g. 5,000 lines) when using `CTkTextbox`.
8. **`on_closing`** calls `sys.exit(0)` after `destroy()`, and does not stop the daemon or Selenium thread first.
9. **`requests.get` without a timeout** in `get_bjdt_ids`; a stalled server freezes the whole BJDT step.
10. **Label typo / inconsistency:** button text "Continue to ODOO Sync" later becomes "Start ODOO Sync". Pick one wording.

---

## 3. Goals and Non-Goals

**Goals**
- 100 % of visible widgets are CustomTkinter (except native file dialogs).
- Single source of truth for colours, fonts, spacing and radius.
- Dark and light mode with a runtime switch (default: dark, matching today).
- Sidebar navigation, card-based layout, consistent buttons, toasts instead of native message boxes.
- UI stays responsive; all UI mutations happen on the main thread.
- Business logic extracted from the window class so UI can change freely.

**Non-goals**
- Changing Excel/BEBAN, BJDT, CSV or DB behaviour (beyond the bug fixes above).
- Rewriting `grid_benda_uji.py` (wx). It stays a separate process, as it is today.
- Rewriting `daemon_sync.py` / `input_rencana_benda_uji.py`.

---

## 4. Target Architecture

```
modules/
├── ui/
│   ├── app.py                  # ctk.CTk root, sidebar + page container, status bar
│   ├── theme.py                # design tokens, fonts, apply_theme()
│   ├── theme.json              # CustomTkinter colour theme
│   ├── dispatcher.py           # UiDispatcher (thread → main thread)
│   ├── icons.py                # load_icon("name") -> CTkImage (cached)
│   ├── components/
│   │   ├── card.py             # SectionCard
│   │   ├── fields.py           # FileField, DirField, LabeledEntry, PasswordEntry
│   │   ├── buttons.py          # PrimaryButton, SecondaryButton, DangerButton, ToggleButton
│   │   ├── status.py           # StatusBadge, StepIndicator
│   │   ├── log_console.py      # LogConsole (CTkTextbox + level tags)
│   │   ├── toast.py            # Toast notifications
│   │   └── dialogs.py          # ConfirmDialog, InfoDialog, ErrorDialog
│   ├── pages/
│   │   ├── files_page.py       # was: File Configuration tab
│   │   ├── database_page.py    # was: Database configure tab
│   │   ├── process_page.py     # was: Process Control tab
│   │   ├── rencana_page.py     # was: Input Rencana Benda Uji tab
│   │   └── logs_page.py        # was: Process Logs tab
│   └── beban_config_editor.py  # existing – restyle with the same components
├── services/                   # extracted from ExcelProcessorGUI (no tkinter imports)
│   ├── excel_service.py        # copy_excel_data, update_idpengujian
│   ├── bjdt_service.py         # get_bjdt_ids
│   ├── csv_service.py          # generate_csv_files
│   └── db_service.py           # test_connection, upload_csv
└── controllers/
    └── process_controller.py   # orchestrates services, runs in threads, emits events
```

**Rule:** `services/` never imports `tkinter`/`customtkinter`. They take plain arguments (a `Settings` dataclass) and a `progress/log` callback, and return a result object (`ok`, `message`, `counts`).

```python
@dataclass
class Settings:
    file_rencana: Path
    file_pengujian: Path
    output_dir: Path
    sheet_source: str = "Report Excel"
    sheet_odoo: str = "ODOO"
    sheet_pengujian: str = "pengujian"
    sheet_bjdt: str = "bjdt_id"
    db: DbSettings = field(default_factory=DbSettings)
```

`Settings` is built from the existing `StringVar`s, which also gives a natural place to **persist user settings** between runs (see §9).

---

## 5. Design System

### 5.1 Palette (dark first, light mapped)

Keeps the existing brand blue (`#4A90E2`), green (`#28A745`) and red (`#DC3545`), but defines proper surface layers instead of one flat `#212121`.

| Token | Dark | Light | Use |
|---|---|---|---|
| `bg` | `#121417` | `#F4F6F9` | Window background |
| `surface` | `#1B1F24` | `#FFFFFF` | Sidebar, cards |
| `surface_alt` | `#232830` | `#EEF1F5` | Inputs, nested panels |
| `border` | `#2E343D` | `#D9DEE5` | Card/input outlines |
| `text` | `#E6E9EE` | `#1B1F24` | Primary text |
| `text_muted` | `#8B94A1` | `#667085` | Secondary text |
| `primary` | `#4A90E2` | `#2F7BD6` | Main actions, active nav |
| `primary_hover` | `#3A7BC8` | `#2467B8` | |
| `success` | `#2EB872` | `#1E9E5C` | Sync start, success badge |
| `danger` | `#E5484D` | `#D13438` | Stop / destructive |
| `warning` | `#F5A524` | `#D98A00` | Warnings in logs |

### 5.2 Typography
One family (e.g. **Inter**, falling back to `Segoe UI`) via `CTkFont`.

| Role | Size | Weight |
|---|---|---|
| Page title | 24 | bold |
| Section title | 16 | bold |
| Body / labels | 13 | normal |
| Helper text | 11 | normal, `text_muted` |
| Log (monospace) | 12 | `Cascadia Mono` / `Consolas` |

### 5.3 Spacing, radius, sizing
- 4-pt grid: spacing tokens `4, 8, 12, 16, 24, 32`.
- Corner radius: cards `12`, inputs/buttons `8`, badges `999`.
- Control height: inputs `36`, buttons `38` (primary `44`).
- Sidebar width `220` (collapses to `64` icon-only below 1000 px window width).

### 5.4 Theme implementation

```python
# theme.py
import customtkinter as ctk
from modules.utils import resource_path

TOKENS = {
    "dark":  {"bg": "#121417", "surface": "#1B1F24", "surface_alt": "#232830",
              "border": "#2E343D", "text": "#E6E9EE", "muted": "#8B94A1",
              "primary": "#4A90E2", "primary_hover": "#3A7BC8",
              "success": "#2EB872", "danger": "#E5484D", "warning": "#F5A524"},
    "light": {"bg": "#F4F6F9", "surface": "#FFFFFF", "surface_alt": "#EEF1F5",
              "border": "#D9DEE5", "text": "#1B1F24", "muted": "#667085",
              "primary": "#2F7BD6", "primary_hover": "#2467B8",
              "success": "#1E9E5C", "danger": "#D13438", "warning": "#D98A00"},
}

def color(name: str) -> tuple[str, str]:
    """(light, dark) tuple — the format CTk uses to switch automatically."""
    return (TOKENS["light"][name], TOKENS["dark"][name])

def apply_theme(mode: str = "dark") -> None:
    ctk.set_appearance_mode(mode)                       # "dark" | "light" | "system"
    ctk.set_default_color_theme(resource_path("modules/ui/theme.json"))
```

Using `(light, dark)` tuples means widgets re-colour themselves when `set_appearance_mode` is called, so **no per-widget colour code** is needed for the dark/light switch.

---

## 6. UI Redesign

### 6.1 Shell layout

Replace `CTkTabview` with a **sidebar + page container + status bar**.

```
┌──────────────┬──────────────────────────────────────────────┐
│  ▣ Excel     │  Page title                       [🌙 Theme] │
│    Processor │  ─────────────────────────────────────────── │
│              │                                              │
│  ▸ Dashboard │   ┌──────────── Card ────────────────────┐   │
│  ▸ Files     │   │                                      │   │
│  ▸ Database  │   └──────────────────────────────────────┘   │
│  ▸ Rencana   │   ┌──────────── Card ────────────────────┐   │
│  ▸ Logs      │   │                                      │   │
│              │   └──────────────────────────────────────┘   │
│  ─────────   ├──────────────────────────────────────────────┤
│  ● Sync: Off │  Ready · last run 22:10 · DB: connected      │
└──────────────┴──────────────────────────────────────────────┘
```

- **Sidebar:** logo/title, nav buttons (icon + label, active state highlighted with `primary`), ODOO sync status pill at the bottom, theme toggle.
- **Page container:** one `CTkFrame` holding all pages; switching = `grid_remove()` the old page and `grid()` the new one (pages are created once, so state is kept).
- **Status bar:** current step text, thin indeterminate `CTkProgressBar` (visible only while busy), DB connection indicator.

### 6.2 Page designs

**Dashboard (replaces "Process Control")** – the page people use daily.
- *Pipeline card:* a horizontal 4-step stepper `Copy Data → Get BJDT ID → Generate CSV → Upload DB`. Each step shows state: pending / running (spinner colour) / done ✓ / failed ✗. "Start Full Process" primary button runs all four.
- *Individual actions card:* the same actions as today (Copy Data, Get BJDT ID, Generate CSV, Upload CSV file, Upload to DB) as a 3-column grid of secondary buttons with icons. Selected CSV name shown beside the upload button (today the chosen file is invisible).
- *Tools card:* Open Grid Benda Uji, Beban Config Editor.
- *ODOO sync card:* large toggle button (green "Start sync" ↔ red "Stop sync") with a status pill and the thread state (running/stopped).

**Files (replaces "File Configuration")**
- *Files card:* three `FileField` rows (Rencana, Pengujian, Output dir) with browse icon button, validation tick/cross showing whether the path exists.
- *Sheets card:* 2×2 grid of `LabeledEntry` (Source, ODOO, Pengujian, BJDT ID). Optionally turn these into `CTkComboBox` populated from the workbook's real sheet names after a file is chosen.

**Database (replaces "Database configure")**
- Single card, two-column form (Host/Port, User/Password, Database/Table).
- Password field with a show/hide eye toggle.
- "Test connection" button with an inline result badge (✓ Connected · 42 ms / ✗ error text) instead of a popup.

**Rencana Benda Uji**
- Hero card with description, Start/Stop toggle button (same `ToggleButton` component), live status badge, and a compact "last run" summary (success / failed / skipped counts as three stat tiles) instead of only a final messagebox.

**Logs**
- Toolbar: level filter (`All / Info / Warning / Error` via `CTkSegmentedButton`), search box, Copy, Save to file, Clear, "Auto-scroll" switch.
- `LogConsole` with coloured levels (INFO muted-white, WARNING amber, ERROR red), monospaced font, line cap.

### 6.3 Interaction patterns

| Old | New |
|---|---|
| `messagebox.showinfo("Success", ...)` | **Toast** (bottom-right, auto-dismiss 4 s), e.g. "CSV files generated" |
| `messagebox.showerror(...)` | **Error dialog** (`CTkToplevel`, modal, copyable detail) + log entry |
| `messagebox.askokcancel("Quit")` | **ConfirmDialog**, only if a process/daemon is running |
| `progress.start()` in a tab | Global status-bar progress + per-step stepper state |
| Disabled buttons while busy | `set_busy(True)` disables a registered group of buttons and shows the spinner |
| Hand-rolled start/stop colour swaps | `ToggleButton.set_state("running" | "idle")` |

**Thread-safe dispatcher** (fixes bug #2 once, everywhere):

```python
# dispatcher.py
import queue

class UiDispatcher:
    """Call .post(fn, *args) from any thread; fn runs on the Tk main loop."""
    def __init__(self, root, interval_ms: int = 50):
        self._root, self._q, self._interval = root, queue.Queue(), interval_ms
        self._pump()

    def post(self, fn, *args, **kwargs):
        self._q.put((fn, args, kwargs))

    def _pump(self):
        try:
            while True:
                fn, a, kw = self._q.get_nowait()
                fn(*a, **kw)
        except queue.Empty:
            pass
        self._root.after(self._interval, self._pump)
```

Services emit events (`step_started`, `step_finished`, `log`, `failed`) through the controller, which forwards them to the dispatcher. Pages never call `messagebox` or touch widgets from threads.

### 6.4 Component specs (key reusable pieces)

```python
# components/card.py
class SectionCard(ctk.CTkFrame):
    def __init__(self, master, title: str, subtitle: str | None = None, icon=None):
        super().__init__(master, fg_color=color("surface"), corner_radius=12,
                         border_width=1, border_color=color("border"))
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(header, text=title, image=icon, compound="left",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(side="left")
        if subtitle:
            ctk.CTkLabel(self, text=subtitle, text_color=color("muted"),
                         font=ctk.CTkFont(size=11)).pack(anchor="w", padx=16)
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self.body.pack(fill="both", expand=True, padx=16, pady=(8, 16))
```

```python
# components/fields.py  – replaces create_file_input / create_directory_input
class FileField(ctk.CTkFrame):
    def __init__(self, master, label, variable, mode="file",
                 filetypes=(("Excel", "*.xlsx *.xls"),), on_change=None):
        ...
        # label (muted) / entry (read+edit) / browse CTkButton (icon) / validity icon
        # mode "dir" uses filedialog.askdirectory; validity = Path(variable.get()).exists()
```

```python
# components/log_console.py
class LogConsole(ctk.CTkTextbox):
    MAX_LINES = 5000
    def __init__(self, master, **kw):
        super().__init__(master, font=("Cascadia Mono", 12), wrap="none", **kw)
        self.tag_config("INFO",    foreground="#C9D1D9")
        self.tag_config("WARNING", foreground="#F5A524")
        self.tag_config("ERROR",   foreground="#E5484D")
        self.configure(state="disabled")

    def append(self, line: str, level: str = "INFO"):
        self.configure(state="normal")
        self.insert("end", line + "\n", level)
        if int(self.index("end-1c").split(".")[0]) > self.MAX_LINES:
            self.delete("1.0", "500.0")
        self.configure(state="disabled")
        if self._autoscroll: self.see("end")
```

> `CTkTextbox` wraps a real `tk.Text`, so `tag_config` colours work; per-tag `font` is **not** supported by CTkTextbox, so level styling is colour-only.

```python
# components/buttons.py
class ToggleButton(ctk.CTkButton):
    """Start (primary/success) ↔ Stop (danger) in one place."""
    def set_state(self, running: bool):
        if running:
            self.configure(text=self._stop_text, fg_color=color("danger"),
                           hover_color=("#B02A2E", "#C23A3F"))
        else:
            self.configure(text=self._start_text, fg_color=self._idle_color,
                           hover_color=self._idle_hover)
```

### 6.5 Icons
- Add a small `assets/icons/` set (24 px PNG, light + dark variants, e.g. from Lucide) and a cached `load_icon(name)` returning `ctk.CTkImage(light_image=..., dark_image=..., size=(20, 20))`.
- Replace all emoji prefixes in labels and tab names.
- Bundle `assets/` with PyInstaller (`--add-data`) and always resolve through `resource_path()`.

### 6.6 Widget mapping cheat-sheet

| Old | New |
|---|---|
| `ctk.CTkTabview` | Sidebar + page frames |
| `ttk.Progressbar` | `ctk.CTkProgressBar(mode="indeterminate")` |
| `ScrolledText` | `LogConsole(ctk.CTkTextbox)` |
| `messagebox.*` | `Toast` / `dialogs.*` |
| `create_file_input` / `create_directory_input` | `FileField(mode="file"/"dir")` |
| `create_sheet_input` | `LabeledEntry` (optionally `CTkComboBox`) |
| `create_db_input` | `LabeledEntry` / `PasswordEntry` |
| Literal hex colours | `color("token")` |
| `root.configure(bg=...)` | `fg_color=color("bg")` on the root frame |
| `tk.StringVar` | unchanged |

---

## 7. Implementation Phases

Each phase leaves the app runnable, so it can be merged and tested independently.

### Phase 0 – Preparation (0.5 day)
- [ ] Create a `ui-modernization` branch; tag the current version for rollback.
- [ ] Pin `customtkinter>=5.2` (and `Pillow`) in `requirements.txt`.
- [ ] Take screenshots of every current tab (baseline for visual comparison).
- [ ] Write a manual smoke-test checklist (see §8).

### Phase 1 – Foundations (1 day)
- [ ] Add `theme.py`, `theme.json`, `dispatcher.py`, `icons.py`.
- [ ] Call `apply_theme()` before the root window is created in `main.py` (root becomes `ctk.CTk`, if it is not already).
- [ ] Set `ctk.deactivate_automatic_dpi_awareness(False)` only if scaling problems appear; otherwise keep CTk defaults.
- **Done when:** the existing UI still works and the root background comes from the theme.

### Phase 2 – Component library (1.5 days)
- [ ] `SectionCard`, `FileField`, `LabeledEntry`, `PasswordEntry`, button variants, `StatusBadge`.
- [ ] `LogConsole`, `Toast`, `dialogs`.
- [ ] A throw-away `demo_components.py` that renders every component in both themes.
- **Done when:** every component looks right in dark and light mode.

### Phase 3 – Shell + navigation (1 day)
- [ ] `app.py`: sidebar, page container, status bar, theme toggle.
- [ ] Register the five pages as empty placeholders; verify switching and window resize.
- [ ] Keep the old `ExcelProcessorGUI` reachable behind a `--legacy-ui` flag until Phase 7.

### Phase 4 – Page migration (2–3 days)
Migrate in order of risk, lowest first:
1. [ ] **Database page** (+ inline test-connection result)
2. [ ] **Files page**
3. [ ] **Logs page** (needs `LogConsole`, ties into the existing `ThreadSafeLogHandler`/`log_queue`)
4. [ ] **Rencana page** (Start/Stop `ToggleButton`, stat tiles)
5. [ ] **Dashboard** (stepper, actions grid, sync card)
6. [ ] Restyle `beban_config_editor.py` with the same components

### Phase 5 – Extract services + controller (2 days)
- [ ] Move logic out of the window class into `services/*` with the `Settings` dataclass.
- [ ] Implement `process_controller.py` (runs steps in worker threads, emits events).
- [ ] Apply the bug fixes from §2 (#1, #3, #4, #5, #6, #9).
- **Done when:** `full_process` runs steps strictly in order, and success is only reported on real success.

### Phase 6 – Polish (1 day)
- [ ] Keyboard: `Ctrl+1…5` page switching, `Esc` closes dialogs, `Enter` triggers the primary action.
- [ ] Tooltips on icon-only buttons.
- [ ] Empty states and validation hints (red border + helper text on invalid path / port).
- [ ] Persist settings and theme choice (§9).
- [ ] Window minimum size (≈ 960×640) and sidebar collapse behaviour.

### Phase 7 – Packaging, QA, cutover (1 day)
- [ ] PyInstaller spec: include `customtkinter` data (themes/assets) and `assets/icons`, `theme.json`.
- [ ] Test on Windows 10/11 at 100 %, 125 %, 150 % display scaling (and a second monitor with different scaling).
- [ ] Remove `--legacy-ui` and the old `main_window.py` after one release cycle.

**Estimated total:** ~9–11 working days for one developer (≈ 5 days if only Phases 1–4 are done and logic extraction is deferred).

---

## 8. Test Plan

**Manual smoke checklist (run in dark and light)**
- [ ] App starts; all five pages open; resizing does not clip or overlap.
- [ ] Browse buttons set paths; invalid paths show the red state.
- [ ] Test DB connection: success and failure (wrong password) both display inline.
- [ ] Copy Data → Get BJDT ID → Generate CSV → Upload to DB individually and via Full Process, in order.
- [ ] Upload a duplicate CSV → error shown, **no** success toast.
- [ ] Rencana: Start → Stop mid-run; UI stays responsive; button returns to idle.
- [ ] ODOO sync: Start → Stop; status pill and sidebar indicator agree with thread state.
- [ ] Close window while a process runs → confirm dialog, clean shutdown (no orphan Selenium/Chrome).
- [ ] Log page: colours, filter, search, clear, save, line cap (generate > 5,000 lines).
- [ ] Beban Config Editor opens in matching style.

**Automated (where practical)**
- Unit tests for `services/*` using small fixture workbooks (no GUI needed – a benefit of the extraction).
- A headless-friendly import test that constructs each page against a `ctk.CTk()` root and destroys it.

---

## 9. Optional Enhancements (post-migration)

- Persist settings (paths, sheet names, DB host/port/user, theme) to `settings.json`; keep the **password out of the file** (OS keyring via `keyring`, or ask each session).
- Per-step timing and a "last run" summary on the Dashboard.
- Recent-files dropdown in `FileField`.
- Export logs for a run with one click.
- Sheet-name `CTkComboBox` populated from the selected workbook.

---

## 10. Risks and Mitigations

| Risk | Impact | Mitigation |
|---|---|---|
| CTk scaling/blurry text on mixed-DPI displays | Medium | Test at 100/125/150 %; `ctk.set_widget_scaling()` if needed |
| `CTkScrollableFrame` is slow with many children | Low–Med | Pages are small; avoid nesting scrollable frames; use plain frames for fixed content |
| `CTkTextbox` slowdown with huge logs | Medium | Line cap, batch inserts from the queue every 100 ms |
| Native `filedialog` doesn't follow the theme | Low | Accepted; it is an OS dialog |
| Worker thread touching widgets → random `TclError`/crashes | High | Single `UiDispatcher`; code-review rule "no widget calls outside the main thread" |
| PyInstaller build missing CTk assets/icons | High | Add `--add-data` entries in Phase 7 and test the **built** exe, not just source |
| wx grid process and Tk in the same app | Medium | Keep the existing subprocess approach; never import `wx` in the UI process |
| Behaviour regressions while extracting services | Medium | Move code verbatim first, then refactor; compare outputs of old vs. new on the same sample files |

---

## 11. Definition of Done

- No `ttk`, `ScrolledText` or `tkinter.messagebox` imports remain in `modules/ui/`.
- All colours/fonts come from `theme.py`; dark and light mode both verified.
- Every UI update originates from the main thread.
- Full Process runs steps in order and reports accurate success/failure.
- No hard-coded credentials in source.
- The built executable runs on a clean Windows machine with the new look, and the legacy UI flag has been removed.

---

## 12. Dependencies

```
customtkinter>=5.2
Pillow>=10          # CTkImage icons
# existing: pandas, openpyxl, xlwings, requests, sqlalchemy, psycopg2
# optional: keyring (credential storage)
```
