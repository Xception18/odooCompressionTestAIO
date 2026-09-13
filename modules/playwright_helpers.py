import time
import logging
try:
    # pyrefly: ignore [missing-import]
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
except ImportError:  # pragma: no cover
    raise RuntimeError("playwright is required: pip install playwright && playwright install chromium")
from modules.exceptions import DuplicateDocketException, NoRecordFoundException

logger = logging.getLogger(__name__)


class BrowserWrapper:
    """
    Wraps playwright + browser + page into a single object that exposes
    a .quit() method so existing callsites (shutdown_manager, etc.) work
    without modification.
    """

    def __init__(self, playwright, browser, page):
        self._playwright = playwright
        self._browser = browser
        self.page = page

    # --- Selenium-compatible surface ---

    def get(self, url: str):
        self.page.goto(url)

    def maximize_window(self):
        try:
            size = self.page.evaluate("({w: window.screen.width, h: window.screen.height})")
            self.page.set_viewport_size({"width": size["w"], "height": size["h"]})
        except Exception:
            self.page.set_viewport_size({"width": 1920, "height": 1080})

    @property
    def current_url(self) -> str:
        return self.page.url

    def refresh(self):
        self.page.reload()

    def find_elements(self, by, value):
        """Compatibility shim used by a few callsites that still call driver.find_elements."""
        locator = self._build_locator(value, by)
        try:
            return locator.all()
        except Exception:
            return []

    def find_element(self, by, value):
        """Compatibility shim."""
        locator = self._build_locator(value, by)
        return locator

    def execute_script(self, script: str, *args):
        """
        Playwright equivalent for driver.execute_script().
        Only the subset actually used in the codebase is supported.
        """
        # The only patterns used:
        #   "arguments[0].scrollIntoView(...);"  -> scroll_into_view_if_needed
        #   "arguments[0].click();"              -> force click
        if args and "scrollIntoView" in script:
            el = args[0]
            if hasattr(el, "scroll_into_view_if_needed"):
                el.scroll_into_view_if_needed()
            return
        if args and script.strip() in (
            "arguments[0].click();",
            "arguments[0].click()",
        ):
            el = args[0]
            if hasattr(el, "click"):
                el.click(force=True)
            return
        # Generic fallback
        js_args = list(args)
        return self.page.evaluate(script, js_args if js_args else None)

    def quit(self):
        try:
            self._browser.close()
        except Exception as e:
            logger.warning(f"Error closing browser: {e}")
        try:
            self._playwright.stop()
        except Exception as e:
            logger.warning(f"Error stopping playwright: {e}")

    # --- Internal ---

    def _build_locator(self, value: str, by: str):
        by = by.lower()
        if by in ("id", "by.id"):
            return self.page.locator(f"#{value}")
        if by in ("css selector", "by.css_selector"):
            return self.page.locator(value)
        if by in ("xpath", "by.xpath"):
            return self.page.locator(f"xpath={value}")
        if by in ("link text", "by.link_text"):
            return self.page.get_by_text(value, exact=True)
        if by in ("tag name", "by.tag_name"):
            return self.page.locator(value)
        return self.page.locator(value)


def setup_driver() -> BrowserWrapper:
    """Setup Playwright browser (Chrome/Chromium). Drop-in for Selenium setup_driver()."""
    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        channel="chrome",  # uses installed Chrome; falls back to bundled Chromium if absent
        headless=False,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--start-maximized"],
    )
    page = browser.new_page(no_viewport=True)
    wrapper = BrowserWrapper(pw, browser, page)
    wrapper.maximize_window()
    return wrapper


# ---------------------------------------------------------------------------
# Timeout helpers
# ---------------------------------------------------------------------------

def _ms(seconds: float) -> int:
    """Convert seconds to milliseconds for Playwright timeouts."""
    return int(seconds * 1000)


def wait_for_loading_overlay_to_disappear(driver: BrowserWrapper):
    """Wait for blockUI loading overlay to disappear."""
    page = driver.page
    loading_selectors = [
        "div.blockUI.blockMsg.blockPage",
        "div.blockUI.blockOverlay",
        "div.oe_blockui_spin_container",
    ]
    try:
        for selector in loading_selectors:
            try:
                page.wait_for_selector(selector, state="hidden", timeout=500)
            except PlaywrightTimeoutError:
                pass
        time.sleep(0.5)
        return True
    except Exception as e:
        logger.warning(f"Error waiting for loading overlay: {e}")
        return False


def is_click_intercepted_error(error_message: str) -> bool:
    """Check if error is due to element click intercepted."""
    error_str = str(error_message).lower()
    return "element click intercepted" in error_str and (
        "blockui" in error_str or "blockoverlay" in error_str
    )


def handle_duplicate_docket_error(driver: BrowserWrapper):
    page = driver.page
    try:
        error_modal = page.locator("div.modal-content.o_error_dialog")
        error_modal.wait_for(state="visible", timeout=1000)
        modal_text = error_modal.inner_text()
        if "Docket ini sudah pernah dibuat benda uji !" in modal_text:
            logger.warning("Detected duplicate docket error!")
            close_btn = error_modal.locator(
                ".modal-footer button.btn-primary, .modal-footer button.btn-secondary"
            ).first
            close_btn.click()
            error_modal.wait_for(state="hidden")
            raise DuplicateDocketException("Docket ini sudah pernah dibuat benda uji !")
    except PlaywrightTimeoutError:
        pass
    except DuplicateDocketException:
        raise
    except Exception as e:
        logger.error(f"An unexpected error occurred while checking for modals: {e}")

def no_record_found(driver: BrowserWrapper, no_docket):
    page = driver.page
    try:
        no_record_selector = "li.o_m2o_no_result a#rel_docket_0_0_0"
        no_record_found = False
        try:
            # Short timeout because we are checking if it appears after typing
            page.wait_for_selector(no_record_selector, state="visible", timeout=1000)
            no_record_found = True
        except:
            no_record_found = False

        if no_record_found:
            logger.warning(f"No record found for '{no_docket}'. Skipping to next.")
            raise NoRecordFoundException(f"No record found for '{no_docket}'. Skipping to next.")
    except PlaywrightTimeoutError:
        pass
    except NoRecordFoundException:
        raise
    except Exception as e:
        logger.error(f"An unexpected error occurred while checking for modals: {e}")
    

def fill_field(driver: BrowserWrapper, locator: str = "", value="", field_name: str = ""):
    """Generic function to fill form fields (ID-based)."""
    page = driver.page
    logger.info(f"Filling {field_name} field with value: {value}")
    try:
        field = page.locator(f"#{locator}")
        field.wait_for(state="visible")
        field.scroll_into_view_if_needed()
        field.click()
        field.fill("")
        time.sleep(0.5)
        field.fill(str(value))

        # Verify value
        try:
            page.wait_for_function(
                f"document.getElementById('{locator}')?.value?.includes('{str(value)}')")
            logger.info(f"{field_name} value entered successfully")
        except PlaywrightTimeoutError:
            logger.warning(f"{field_name} field did not reflect the typed value immediately")
    except PlaywrightTimeoutError:
        logger.error(f"Could not find {field_name} field")
        raise


def select_first_row_in_modal_and_confirm(
    driver: BrowserWrapper,
    row_text: str | None = None,
    absolute_xpath: str | None = None,
):
    """Select first row in modal and confirm selection."""
    page = driver.page
    max_attempts = 1
    last_error = None

    for attempt in range(1, max_attempts + 1):
        try:
            # Find visible modal
            modal_sel = "div.modal.show, div.modal.in, div.modal-content, div.modal[style*='display: block']"
            modal = page.locator(modal_sel).last
            modal.wait_for(state="visible")

            # Find table in modal
            table = None
            for selector in [
                ".modal-content table.o_list_view",
                ".modal-content table.table.o_list_view",
                ".modal-content table.table-condensed.table-striped.o_list_view",
            ]:
                t = modal.locator(selector)
                if t.count() > 0:
                    table = t.first
                    break

            if table is None:
                t = modal.locator("xpath=.//table[contains(@class,'o_list_view')]")
                if t.count() == 0:
                    raise Exception("No list table found inside modal")
                table = t.first

            target_row = None

            # Search by row_text
            if row_text:
                row_xpath = f"xpath=.//tbody/tr[.//td[contains(normalize-space(), \"{row_text}\")] or contains(normalize-space(), \"{row_text}\") or contains(., \"{row_text}\")]"
                matching = table.locator(row_xpath)
                if matching.count() > 0:
                    target_row = matching.first

            # Fallback to absolute xpath
            if target_row is None and absolute_xpath:
                t = page.locator(f"xpath={absolute_xpath}")
                if t.count() > 0:
                    target_row = t.first

            # Fallback to first row
            if target_row is None:
                rows = table.locator("tbody tr")
                if rows.count() == 0:
                    raise Exception("No rows found inside the modal list view")
                target_row = rows.first

            # Scroll and click
            target_row.scroll_into_view_if_needed()
            time.sleep(0.05)

            # Find clickable element
            clickable = None
            for sel in [
                "td.o_list_record_selector input",
                "td.o_list_record_selector",
                "td a",
                "td",
            ]:
                c = target_row.locator(sel)
                if c.count() > 0:
                    clickable = c.first
                    break

            if clickable is None:
                clickable = target_row

            try:
                clickable.click()
            except PlaywrightError:
                clickable.click(force=True)

            # Click Select button
            try:
                select_btn = modal.locator(".modal-footer .btn.btn-primary, .modal-footer .o_select_button").first
                if select_btn.count() > 0:
                    try:
                        select_btn.click()
                    except PlaywrightError:
                        select_btn.click(force=True)
            except Exception:
                pass

            # Wait for modal to disappear
            gone_sel = "div.modal.show, div.modal.in, div.modal[style*='display: block']"
            try:
                page.wait_for_selector(gone_sel, state="hidden", timeout=5000)
            except PlaywrightTimeoutError:
                # Check if modal is already gone
                if page.locator(gone_sel).count() == 0:
                    return
            return

        except Exception as e:
            last_error = e
            logger.info(f"Retry selecting row in modal due to transient error: {e} (attempt {attempt}/{max_attempts})")
            time.sleep(0.2)
            continue

    # Check if modal is gone
    gone_sel = "div.modal.show, div.modal.in, div.modal[style*='display: block']"
    if page.locator(gone_sel).count() == 0:
        logger.info("Modal no longer visible; treating selection as successful.")
        return

    logger.error(f"Failed selecting row in modal after {max_attempts} attempts: {last_error}")
    if last_error is not None:
        raise last_error
    raise Exception("Failed selecting row in modal")


def quick_delete_all(driver: BrowserWrapper):
    """Delete all rows by clicking delete buttons."""
    page = driver.page
    deleted_count = 0
    try:
        while True:
            delete_buttons = page.locator("td.o_list_record_remove.text-center > button")
            if delete_buttons.count() == 0:
                break
            first_button = delete_buttons.first
            try:
                first_button.scroll_into_view_if_needed()
                time.sleep(0.1)
                first_button.click()
                deleted_count += 1
                time.sleep(0.2)
            except Exception as e:
                logger.error(f"Failed to click delete button: {e}")
                break

        logger.info(f"Quick delete: {deleted_count} rows deleted!")
        return deleted_count
    except Exception as e:
        logger.error(f"Error in quick_delete_all: {e}")
        return deleted_count


def quick_delete_excess_rows(driver: BrowserWrapper, rows_to_delete: int):
    """Delete specific number of excess rows starting from the last row."""
    page = driver.page
    deleted_count = 0
    try:
        for i in range(rows_to_delete):
            delete_buttons = page.locator("td.o_list_record_remove.text-center > button")
            if delete_buttons.count() == 0:
                break
            last_button = delete_buttons.last
            try:
                last_button.scroll_into_view_if_needed()
                time.sleep(0.1)
                last_button.click()
                deleted_count += 1
                time.sleep(0.2)
            except Exception as e:
                logger.error(f"Failed to click delete button: {e}")
                break

        logger.info(f"Quick delete excess: {deleted_count} rows deleted from bottom!")
        return deleted_count
    except Exception as e:
        logger.error(f"Error in quick_delete_excess_rows: {e}")
        return deleted_count
