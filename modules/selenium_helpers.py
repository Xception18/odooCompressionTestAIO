import time
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException, StaleElementReferenceException

from modules.utils import resource_path

logger = logging.getLogger(__name__)

def setup_driver():
    """Setup Chrome driver"""
    chromedriver_path = resource_path("chromedriver.exe")
    service = Service(executable_path=chromedriver_path)
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(service=service, options=options)
    driver.maximize_window()
    return driver

def wait_for_loading_overlay_to_disappear(driver, wait, max_wait=20):
    """Wait for blockUI loading overlay to disappear"""
    try:
        loading_selectors = [
            "div.blockUI.blockMsg.blockPage",
            "div.blockUI.blockOverlay",
            "div.oe_blockui_spin_container"
        ]
        for selector in loading_selectors:
            try:
                wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, selector)))
            except:
                pass
        time.sleep(0.5)
        return True
    except Exception as e:
        logger.warning(f"Error waiting for loading overlay: {e}")
        return False

def is_click_intercepted_error(error_message):
    """Check if error is due to element click intercepted"""
    error_str = str(error_message).lower()
    return "element click intercepted" in error_str and ("blockui" in error_str or "blockoverlay" in error_str)

def fill_field(driver, wait, xpath, value, field_name):
    """Generic function to fill form fields"""
    logger.info(f"Filling {field_name} field with value: {value}")
    try:
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
        time.sleep(0.2)
        field.click()
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys(Keys.DELETE)
        field.send_keys(str(value))
        field.send_keys(Keys.TAB)
        
        # Verify value entered
        try:
            wait.until(lambda d: str(value) in (field.get_attribute('value') or ''))
            logger.info(f"{field_name} value entered successfully")
        except TimeoutException:
            logger.warning(f"{field_name} field did not reflect the typed value immediately")
    except TimeoutException:
        logger.error(f"Could not find {field_name} field")
        raise

def select_first_row_in_modal_and_confirm(driver, wait, row_text: str | None = None, absolute_xpath: str | None = None):
    """Select first row in modal and confirm selection"""
    max_attempts = 1
    last_error = None
    
    for attempt in range(1, max_attempts + 1):
        try:
            # Find visible modal
            visible_modals = driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']")
            modal = visible_modals[-1] if visible_modals else wait.until(EC.visibility_of_element_located((
                By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"
            )))

            # Find table in modal
            table = None
            for selector in [
                ".modal-content table.o_list_view",
                ".modal-content table.table.o_list_view",
                ".modal-content table.table-condensed.table-striped.o_list_view"
            ]:
                try:
                    table = modal.find_element(By.CSS_SELECTOR, selector)
                    break
                except NoSuchElementException:
                    continue
            
            if table is None:
                try:
                    table = modal.find_element(By.XPATH, ".//table[contains(@class,'o_list_view')]")
                except NoSuchElementException:
                    raise NoSuchElementException("No list table found inside modal")

            target_row = None

            # Search by row_text first
            if row_text:
                row_xpath = f".//tbody/tr[.//td[contains(normalize-space(), \"{row_text}\")] or contains(normalize-space(), \"{row_text}\") or contains(., \"{row_text}\")]"
                matching_rows = table.find_elements(By.XPATH, row_xpath)
                if matching_rows:
                    target_row = matching_rows[0]

            # Fallback to absolute xpath
            if target_row is None and absolute_xpath:
                try:
                    target_row = driver.find_element(By.XPATH, absolute_xpath)
                except NoSuchElementException:
                    target_row = None

            # Fallback to first row
            if target_row is None:
                rows = table.find_elements(By.CSS_SELECTOR, "tbody tr")
                if not rows:
                    rows = table.find_elements(By.XPATH, ".//tbody/tr")
                if not rows:
                    raise NoSuchElementException("No rows found inside the modal list view")
                target_row = rows[0]

            # Scroll and click
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", target_row)
            time.sleep(0.05)

            # Find clickable element
            clickable_selectors = [
                "td.o_list_record_selector input",
                "td.o_list_record_selector",
                "td a",
                "td"
            ]
            clickable = None
            for sel in clickable_selectors:
                try:
                    clickable = target_row.find_element(By.CSS_SELECTOR, sel)
                    break
                except NoSuchElementException:
                    continue
            
            if clickable is None:
                clickable = target_row

            try:
                driver.execute_script("arguments[0].click();", clickable)
            except WebDriverException:
                ActionChains(driver).move_to_element(clickable).pause(0.05).click().perform()

            # Click Select button in modal footer
            try:
                select_btn = modal.find_element(By.CSS_SELECTOR, ".modal-footer .btn.btn-primary, .modal-footer .o_select_button")
                try:
                    driver.execute_script("arguments[0].click();", select_btn)
                except WebDriverException:
                    select_btn.click()
            except NoSuchElementException:
                pass

            # Wait for modal to disappear
            try:
                WebDriverWait(driver, 5).until(EC.invisibility_of_element_located((
                    By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"
                )))
            except TimeoutException:
                try:
                    WebDriverWait(driver, 2).until(EC.staleness_of(modal))
                except TimeoutException:
                    if not driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"):
                        return
            return

        except StaleElementReferenceException as e:
            last_error = e
            logger.info(f"Retry selecting row in modal due to stale element (attempt {attempt}/{max_attempts})")
            time.sleep(0.2)
            continue
        except (TimeoutException, NoSuchElementException, WebDriverException) as e:
            last_error = e
            logger.info(f"Retry selecting row in modal due to transient error: {e} (attempt {attempt}/{max_attempts})")
            time.sleep(0.2)
            continue

    # Check if modal is gone (success)
    if not driver.find_elements(By.CSS_SELECTOR, "div.modal.show, div.modal.in, div.modal[style*='display: block']"):
        logger.info("Modal no longer visible; treating selection as successful.")
        return

    logger.error(f"Failed selecting row in modal after {max_attempts} attempts: {last_error}")
    if last_error is not None:
        raise last_error
    else:
        raise Exception("Failed selecting row in modal")

def quick_delete_all(driver):
    """Delete all rows by clicking delete buttons"""
    deleted_count = 0
    try:
        while True:
            delete_buttons = driver.find_elements(By.CSS_SELECTOR, 'tr[data-id^="one2many_v_id_"] td.o_list_record_delete .fa-trash-o')
            if not delete_buttons:
                break
            
            first_button = delete_buttons[0]
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", first_button)
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

def quick_delete_excess_rows(driver, rows_to_delete):
    """Delete specific number of excess rows starting from the last row"""
    deleted_count = 0
    try:
        for i in range(rows_to_delete):
            delete_buttons = driver.find_elements(By.CSS_SELECTOR, 'tr[data-id^="one2many_v_id_"] td.o_list_record_delete .fa-trash-o')
            if not delete_buttons:
                break
            
            # Get the last button instead of first
            last_button = delete_buttons[-1]
            try:
                driver.execute_script("arguments[0].scrollIntoView(true);", last_button)
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
