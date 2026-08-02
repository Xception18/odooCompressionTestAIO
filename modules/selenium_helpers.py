# selenium_helpers.py — re-export shim
# All functionality has been migrated to playwright_helpers.py.
# This file exists for backward compatibility so other modules
# (e.g. shutdown_manager.py) that import from here continue to work.

from modules.playwright_helpers import (  # noqa: F401
    BrowserWrapper,
    setup_driver,
    wait_for_loading_overlay_to_disappear,
    is_click_intercepted_error,
    handle_duplicate_docket_error,
    fill_field,
    select_first_row_in_modal_and_confirm,
    quick_delete_all,
    quick_delete_excess_rows,
)
