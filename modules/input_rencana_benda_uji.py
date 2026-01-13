# ODOO Automation Script - Excel Integration
import os, sys, time, logging
from pathlib import Path
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException

# Import from modules
from modules.utils import resource_path
from modules.selenium_helpers import (
    setup_driver, 
    wait_for_loading_overlay_to_disappear, 
    is_click_intercepted_error, 
    fill_field, 
    select_first_row_in_modal_and_confirm, 
    quick_delete_excess_rows
)
from modules.data_generators import (
    generate_random_slump_test, 
    generate_random_yield, 
    calculate_jam_sample
)
from modules.excel_handler import ExcelDataProcessor

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(funcName)s : %(lineno)d - %(message)s')
logger = logging.getLogger('ExcelProcessor')
log_file = "automation_log.txt"

def logger_debug(pesan):
    # Keep this for backward compatibility if used elsewhere, or move to utils
    from datetime import datetime
    waktu = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    log_message = f"{waktu} {pesan}"
    
    try:
        with open(log_file, "a", encoding="utf-8") as f:
            f.write(f"{log_message}\n")
    except Exception as e:
        safe_message = log_message.encode('ascii', 'ignore').decode('ascii')
        with open(log_file, "a") as f:
            f.write(f"{safe_message}\n")
    
    print(log_message, flush=True)

def login(driver, wait):
    """Handle login process"""
    username = "oenoseven@gmail.com"
    password = "Odoo2026"
    login_url = "https://rmc.adhimix.web.id/web/login"
    logger.info("Navigating to login page...")
    driver.get(login_url)
    logger.info("Entering credentials...")
    username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
    password_field = driver.find_element(By.ID, "password")
    username_field.send_keys(username)
    password_field.send_keys(password)
    logger.info("Clicking login button...")
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='submit' and contains(@class, 'btn-primary') and text()='Log in']")))
    login_button.click()
    time.sleep(1)
    logger.info(f"Current URL after login: {driver.current_url}")

def navigate_and_create(driver, wait):
    wait_for_loading_overlay_to_disappear(driver, wait)
    """Navigate to create page and click create button"""
    create_rbu_url = "https://rmc.adhimix.web.id/web#model=schedule.truck.mixer.benda.uji&view_type=list&cids=124&menu_id=669"
    logger.info("Navigating to CREATE RENCANA BENDA UJI page...")
    driver.get(create_rbu_url)
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("CREATE RENCANA BENDA UJI...")
    create_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "body > div.o_action_manager > div > div.o_control_panel.d-flex.flex-column.gap-3.gap-lg-1.px-3.pt-2.pb-3 > div > div.o_control_panel_breadcrumbs.d-flex.align-items-center.gap-1.order-0.h-lg-100 > div.o_control_panel_main_buttons.d-flex.gap-1.d-empty-none.d-print-none > div.d-none.d-xl-inline-flex.gap-1 > button")))
    create_button.click()

def fill_proyek_form(driver, wait, row_data):
    """Fill main form fields using Excel data"""
    # Date field - from Excel column 1 (index 0)
    wait_for_loading_overlay_to_disappear(driver, wait)
    tgl_mulai_prod = str(row_data.iloc[0]) if len(row_data) > 0 else "None"
    logger.info(f"Filling Date form with: {tgl_mulai_prod}")
    tgl_field = wait.until(EC.element_to_be_clickable((By.ID, "mulai_produksi_docket_0")))
    tgl_field.clear()
    tgl_field.send_keys(tgl_mulai_prod)

def fill_docket_form(driver, wait, row_data):
    """Fill docket form using Excel data"""
    # No. Docket field - from Excel column 2 (index 1)
    wait_for_loading_overlay_to_disappear(driver, wait)
    no_docket = str(row_data.iloc[1]) if len(row_data) > 1 else "None"
    logger.info(f"Filling No. Docket field with: {no_docket}")
    # Click and fill the No. Docket field
    no_docket_field = wait.until(EC.element_to_be_clickable((By.ID, "rel_docket_0")))
    driver.execute_script("arguments[0].scrollIntoView(true);", no_docket_field)
    time.sleep(1)
    no_docket_field.click()
    no_docket_field.clear()
    no_docket_field.send_keys(no_docket)
    
    try:
        time.sleep(1)
        # Check if the element exists without waiting
        xpath = f"//a[@role='option' and contains(@class, 'dropdown-item') and contains(@class, 'ui-menu-item-wrapper') and normalize-space()='{no_docket}']"
        elements = driver.find_elements(By.XPATH, xpath)
        
        if elements and elements[0].is_displayed():
            # Element found, click it
            logger.info(f"Found {no_docket} in autocomplete dropdown")
            driver.execute_script("arguments[0].click();", elements[0])
            time.sleep(1)
            logger.info(f"Successfully clicked on {no_docket} from autocomplete dropdown")

        else:
            logger.info(f"Element with text '{no_docket}' not found in autocomplete dropdown")
            logger.info("Using 'Search more...' option as fallback")
            # Search more option
            search_more = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[2]/li[8]/a")))
            search_more.click()
            # Search in modal
            modal = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, ".modal-content")))
            try:
                search_input = modal.find_element(By.CSS_SELECTOR, ".o_searchview input.o_searchview_input")
            except NoSuchElementException:
                search_input = modal.find_element(By.CSS_SELECTOR, ".o_searchview input.o_input")
            search_input.click()
            search_input.clear()
            search_input.send_keys(no_docket)
            search_input.send_keys(Keys.ENTER)
            # Wait for search results and select
            wait.until(EC.visibility_of_element_located((By.XPATH, f"//div[contains(@class,'modal-content')]//div[contains(@class,'o_searchview_facet')][.//span[contains(@class,'o_searchview_facet_label')][normalize-space()='No. Docket']]//div[contains(@class,'o_facet_values')]//span[contains(normalize-space(), \"{no_docket}\")]")))
            time.sleep(3)
            select_first_row_in_modal_and_confirm(driver, wait, row_text=no_docket)   
    except Exception as e:
        logger.error(f"Error in docket selection: {str(e)}")

    # Fill remaining fields using Excel data
    slump_value = str(row_data.iloc[6]) if len(row_data) > 6 else "12" # Column 7 (index 6)
    slump_mapping = {
        "Slump 12.0 +1.5/-1.5": "12",
        "Slump 12.0 +2.0/-2.0": "12",
        "Slump 10.0 +2.0/-2.0": "10",
        "Slump 55.0 +10.0/-0.0": "55"
    }
    slump_rencana = next(
    (value for key, value in slump_mapping.items() if key in slump_value),
    "default_value"  # fallback if no match
    )
    slump_test = generate_random_slump_test(slump_rencana)
    yield_value = generate_random_yield()
    nama_teknisi = str(row_data.iloc[4]) if len(row_data) > 4 else "TEKNISI"  # Column 5 (index 4)
    # Try to extract base_jam from web
    try:
        # Selector for: <div name="jam_berangkat" ...><span class="text-truncate">...</span></div>
        jam_berangkat_elem = driver.find_element(By.CSS_SELECTOR, "div[name='jam_berangkat'] span.text-truncate")
        base_jam = jam_berangkat_elem.text.strip()
        logger.info(f"Extracted base_jam from web: {base_jam}")
    except Exception as e:
        logger.warning(f"Could not extract base_jam from web, falling back to Excel. Error: {e}")
        base_jam = str(row_data.iloc[8]) if len(row_data) > 8 else "10:30"  # Column 9 (index 8)
    jam_sample = calculate_jam_sample(base_jam)
    logger.info(f"Calculated jam_sample: {jam_sample}")
    base_xpath = "/html/body/div[2]/div/div/div[2]/div/div/div[2]/div[1]/div[2]/"
    form_fields = [
        (f"{base_xpath}div[3]/div[2]/div/input", slump_rencana, "Slump Rencana"),
        (f"{base_xpath}div[4]/div[2]/div/input", slump_test, "Slump Test"),
        (f"{base_xpath}div[5]/div[2]/div/input", yield_value, "Yield"),
    ]
    for xpath, value, field_name in form_fields:
        time.sleep(1)
        fill_field(driver, wait, xpath, value, field_name)
    
    # Fill remaining fields using Excel data
    time.sleep(1)
    jam_sample_field = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[id='date_sample_0'][data-field='date_sample']")))
    jam_sample_field.click()
    jam_sample_field.send_keys(jam_sample)
    jam_sample_field.send_keys(Keys.TAB)
    time.sleep(1)
    nama_teknisi_field = wait.until(EC.element_to_be_clickable((By.ID, "nama_teknisi_0")))
    nama_teknisi_field.click()
    nama_teknisi_field.clear()
    nama_teknisi_field.send_keys(nama_teknisi)
    time.sleep(1)
    jam_sample_0 = wait.until(EC.element_to_be_clickable((By.ID, "jam_sample_0")))
    jam_sample_0.click()
    jam_sample_0.clear()

    # Format HH:MM specifically for jam_sample_0 while keeping jam_sample intact
    jam_to_input = jam_sample
    if ' ' in jam_sample:
        try:
            # Extract HH:MM from "DD/MM/YYYY HH:MM:SS"
            jam_to_input = jam_sample.split(' ')[1][:5]
        except:
            pass

    jam_sample_0.send_keys(jam_to_input)

def data_to_input(driver, no_urut, row_data, is_first_row=False, stop_event=None):
    """Input data to the table row using Excel data with robust selectors"""
    wait = WebDriverWait(driver, 5)
    # Determine test age based on sequence number
    rencana_umur_test = "7" if no_urut in [1, 2] else "28"
    # Get kode benda uji from Excel (column 3, index 2)
    kode_benda_uji = str(row_data.iloc[2]) if len(row_data) > 2 else f" Isi Kode Benda Uji - {no_urut}"
    bentuk_benda_uji = "Silinder 15 x 30 cm"
    logger.info(f"Input Row {no_urut} on data table...")
    time.sleep(1)

    # 1. FIND THE TARGET ROW
    # Use tr.o_data_row to find all data rows.
    # Note: Odoo 16/17+ often uses virtual IDs or just indices.
    rows = driver.find_elements(By.CSS_SELECTOR, "div[name='benda_uji_ids'] table tbody tr.o_data_row")
    
    if is_first_row:
        # Assuming the first row is the first one in the table
        target_row = rows[0]
        # Click it to activate edit mode if needed (though usually "Add a line" focuses implicitly)
        try:
             target_row.click()
        except:
             pass 
    else:
        # Find the row corresponding to no_urut (1-based index)
        if len(rows) >= no_urut:
            target_row = rows[no_urut - 1]
            try:
                target_row.click() # Ensure row is active
            except:
                pass
        else:
            logger.warning(f"Row {no_urut} not found, using last available row")
            target_row = rows[-1]
            try:
                target_row.click()
            except:
                pass

    # Helper function to fill fields within the specific row
    def fill_row_field(row_element, field_name, value, is_select=False, is_autocomplete=False):
        try:
            # Find the cell by name attribute on the td
            cell = row_element.find_element(By.CSS_SELECTOR, f"td[name='{field_name}']")
            
            if is_select:
                input_el = cell.find_element(By.TAG_NAME, "select")
                input_el.click()
                time.sleep(0.5)
                # Select option by text content if possible, or value
                # Using xpath to find option with text containing value
                # Since input_el is a select, we can use Select class or clicking options
                # Simplified: assume value is what we want to select
                try:
                    option = input_el.find_element(By.XPATH, f".//option[contains(text(), '{value}')]")
                    option.click()
                except:
                   # If exact text fails, try standard keys or value
                   # Try selecting by index if specific values known
                   if value == "Internal":
                       # Assuming 2nd option as per original code logic (index 1)
                       input_el.find_elements(By.TAG_NAME, "option")[1].click()
                   
            elif is_autocomplete:
                input_el = cell.find_element(By.TAG_NAME, "input")
                input_el.click()
                input_el.clear()
                input_el.send_keys(str(value))
                time.sleep(1)
                # Wait for autocomplete dropdown
                try:
                     # Odoo autocomplete usually appears in a dropdown
                     # We can press Enter to select first option if it highlights, 
                     # but finding the dropdown result is safer. 
                     # Pressing arrow down + enter is a common fallback
                     input_el.send_keys(Keys.ENTER)
                except Exception as e:
                     logger.warning(f"Autocomplete selection failed for {field_name}: {e}")

            else: # Standard input
                input_el = cell.find_element(By.TAG_NAME, "input")
                input_el.click()
                input_el.clear()
                input_el.send_keys(str(value))
                
        except Exception as e:
            logger.error(f"Error filling {field_name}: {e}")

    # 2. FILL DATA
    if stop_event and stop_event.is_set():
        return

    # No Urut
    fill_row_field(target_row, "nomor_urut", str(no_urut))
    time.sleep(0.5)

    # Kode Benda Uji
    fill_row_field(target_row, "kode_benda_uji", kode_benda_uji)
    time.sleep(0.5)

    # Rencana Umur Test (benda_uji_id)
    # Value is "7" or "28" (days) - assuming text match works for autocomplete
    fill_row_field(target_row, "benda_uji_id", rencana_umur_test, is_autocomplete=True)
    time.sleep(0.5)

    # Bentuk Benda Uji
    fill_row_field(target_row, "bentuk_benda_uji", bentuk_benda_uji, is_autocomplete=True)
    time.sleep(0.5)

    # Tempat Pengetesan
    fill_row_field(target_row, "tempat_pengetesan", "Internal", is_select=True)
    time.sleep(0.5)
    
    wait_for_loading_overlay_to_disappear(driver, wait)


def add_table_rows(driver, wait, row_data, stop_event=None):
    """Add and fill table rows using Excel data"""
    logger.info("Processing table rows...")
    
    # Check how many existing rows
    # Check how many existing rows
    # Use robust selector for rows found via live web analysis
    # Scoped to benda_uji_ids to ensure we target the correct table
    existing_rows = driver.find_elements(By.CSS_SELECTOR, "div[name='benda_uji_ids'] .o_data_row")
    existing_count = len(existing_rows)
    logger.info(f"Found {existing_count} existing rows")
    
    # Delete excess rows if more than 4 exist
    if existing_count > 4:
        logger.info(f"More than 4 rows exist ({existing_count}), deleting excess rows...")
        rows_to_delete = existing_count - 4
        # Note: quick_delete_excess_rows might need check if it uses old selectors
        # For now, assuming it works or we should refactor it too if finding row fails
        # But let's proceed with filling first. 
        quick_delete_excess_rows(driver, rows_to_delete) 
        time.sleep(1)
        existing_rows = driver.find_elements(By.CSS_SELECTOR, "div[name='benda_uji_ids'] .o_data_row")
        existing_count = len(existing_rows)
    
    # Get add item link
    # New selector: scoped to benda_uji_ids
    try:
        add_item_link = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[name='benda_uji_ids'] .o_field_x2many_list_row_add a")))
    except:
        # Fallback to text
        add_item_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Add a line")))
    
    if existing_count < 4:
        # Fill existing rows first
        logger.info("Less than 4 rows exist, filling existing rows first...")
        for no_urut in range(1, existing_count + 1):
            if stop_event and stop_event.is_set(): return
            logger.info(f"Filling existing row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)
        
        # Add and fill remaining rows
        for no_urut in range(existing_count + 1, 5):
            if stop_event and stop_event.is_set(): return
            logger.info(f"Adding new row {no_urut}...")
            # Click add line
            add_item_link.click()
            time.sleep(1) 
            data_to_input(driver, no_urut, row_data, is_first_row=False, stop_event=stop_event)
    else:
        # Exactly 4 or more, fill first 4
        logger.info("Sufficient rows exist, filling first 4...")
        for no_urut in range(1, 5):
            if stop_event and stop_event.is_set(): return
            logger.info(f"Filling row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)

def save_form(driver, wait):
    """Save the form"""
    time.sleep(2)
    logger.info("Saving Rencana Benda Uji...")
    save_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".o_form_button_save")))
    logger.info("Save button found, clicking...")
    save_button.click()
    time.sleep(3)

def create_form(wait):
    """Create new form"""
    logger.info("Creating new form...")
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[1]/div/div[1]/div[1]/div[2]/button")))
    logger.info("Create button found, clicking...")
    create_button.click()

def duplicate_form(driver, wait, next_row_data, stop_event=None):
    """Duplicate form for next entry with same kode_benda_uji and proyek"""
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Duplicating Rencana Benda Uji...")
    action_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div[2]/div/div/button/i")))
    action_button.click()
    time.sleep(1)
    duplicate_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/div/div[1]/div/div[1]/div[2]/div/div[2]/div/div/div/span[1]")))
    logger.info("Duplicate button found, clicking...")
    duplicate_button.click()
    time.sleep(2)
    # Update the duplicated form with next row data
    logger.info("Updating duplicated form with next row data...")
    fill_docket_form(driver, wait, next_row_data)
    logger.info("Input data to the table row using Excel data")
    if stop_event and stop_event.is_set():
        driver.quit()
        return
    add_table_rows(driver, wait, next_row_data, stop_event=stop_event)

def alternative_form(driver, wait, next_row_data, stop_event=None):
    # Update the duplicated form with next row data
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Updating duplicated form with next row data...")
    fill_docket_form(driver, wait, next_row_data)
    logger.info("Input data to the table row using Excel data")
    if stop_event and stop_event.is_set():
        driver.quit()
        return
    add_table_rows(driver, wait, next_row_data, stop_event=stop_event)

def refresh_and_wait(driver, wait):
    """Refresh page and wait for it to load"""
    logger.info("Refreshing page...")
    driver.refresh()
    time.sleep(3)
    wait_for_loading_overlay_to_disappear(driver, wait)
    time.sleep(2)

def process_excel_row_with_retry(driver, wait, excel_processor, row_data, row_index, max_retries=3, stop_event=None):
    """Process single Excel row with retry logic for click intercepted errors"""
    no_docket = row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing Excel row {row_index + 1} - No. Docket: {no_docket}")

    for attempt in range(max_retries):
        try:
            if stop_event and stop_event.is_set():
                driver.quit()
                return False, no_docket, "Stopped by user"
            navigate_and_create(driver, wait)

            wait_for_loading_overlay_to_disappear(driver, wait)
            fill_proyek_form(driver, wait, row_data)

            wait_for_loading_overlay_to_disappear(driver, wait)
            fill_docket_form(driver, wait, row_data)

            wait_for_loading_overlay_to_disappear(driver, wait)
            if stop_event and stop_event.is_set():
                driver.quit()
                return False, no_docket, "Stopped by user"
            add_table_rows(driver, wait, row_data, stop_event=stop_event)

            save_form(driver, wait)
            logger.info(f"Success processing row {row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""
            
        except ElementClickInterceptedException as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {row_index + 1} (No. Docket: {no_docket})")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    continue
                else:
                    logger.error(f"Max retries reached for row {row_index + 1} (No. Docket: {no_docket}). Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                # Not a blockUI error, don't retry
                return False, no_docket, error_message
                
        except Exception as e:
            error_message = str(e)
            logger.warning(f"Click intercepted detected on attempt {attempt + 1}/{max_retries} for row {row_index + 1}")
                
            if attempt < max_retries - 1:
                logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                refresh_and_wait(driver, wait)
                continue
            else:
                logger.error(f"Max retries reached for row {row_index + 1}. Skipping...")
                return False, no_docket, f"Click intercepted after {max_retries} attempts"
            
    return False, no_docket, "Unknown error after all retries"

def process_duplicate_row_with_retry(driver, wait, next_row_data, next_row_index, max_retries=3, stop_event=None):
    """Process next row using duplicate form with retry logic"""
    no_docket = next_row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing row {next_row_index + 1} using duplicate form - No. Docket: {no_docket}")
    
    # Flag untuk mengontrol dua opsi fungsi
    use_duplicate = True
    
    for attempt in range(max_retries):
        try:
            # DUA OPSI: duplicate_form atau alternative_form
            if use_duplicate:
                logger.info(f"Using duplicate_form on attempt {attempt + 1}")
                duplicate_form(driver, wait, next_row_data, stop_event=stop_event)
            else:
                logger.info(f"Using alternative_form on attempt {attempt + 1}")
                alternative_form(driver, wait, next_row_data, stop_event=stop_event)
            
            # Selalu panggil save_form setelah form processing
            save_form(driver, wait)
            logger.info(f"Success processing row {next_row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""
            
        except ElementClickInterceptedException as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {next_row_index + 1}")
                
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver, wait)
                    navigate_and_create(driver, wait)
                    fill_proyek_form(driver, wait, next_row_data)
                    
                    # SWITCH STRATEGI: Jika duplicate gagal, gunakan alternative
                    if use_duplicate:
                        use_duplicate = False
                        logger.info("Switching from duplicate_form to alternative_form strategy")
                    
                    continue  # Kembali ke awal loop dengan strategi baru
                else:
                    logger.error(f"Max retries reached for row {next_row_index + 1}. Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                return False, no_docket, error_message
                
        except Exception as e:
            error_message = str(e)
            logger.warning(f"Click intercepted detected on attempt {attempt + 1}/{max_retries}")
                
            if attempt < max_retries - 1:
                logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                refresh_and_wait(driver, wait)
                navigate_and_create(driver, wait)
                fill_proyek_form(driver, wait, next_row_data)
                    
                    # SWITCH STRATEGI: Dari duplicate ke alternative
                if use_duplicate:
                    use_duplicate = False
                    logger.info("Exception occurred, switching from duplicate_form to alternative_form strategy")
                    
                continue  # Kembali ke loop dengan fungsi berbeda
            else:
                logger.error(f"Max retries reached for row {next_row_index + 1}. Skipping...")
                return False, no_docket, f"Click intercepted after {max_retries} attempts"
    
    return False, no_docket, "Unknown error after all retries"

def prepare_for_next_row(driver, wait, excel_processor, row_index):
    """Prepare for next row - either duplicate or create new form"""
    try:
        wait_for_loading_overlay_to_disappear(driver, wait)
        if excel_processor.should_duplicate(row_index):
            logger.info("Next row has same kode_benda_uji and proyek - will use duplicate")
            return "duplicate"
        else:
            logger.info("Next row has different kode_benda_uji or proyek - creating new form")
            create_form(wait)
            return "create"
        
    except Exception as e:
        logger.error(f"Error preparing for next row after {row_index + 1}: {e}")
        return "error"

def log_processing_summary(successful_rows, failed_rows, skipped_rows, last_success_info, last_failure_info):
    """Log processing summary with last success/failure details"""
    logger.info(f"{'='*60}")
    logger_debug("="*60)
    logger.info("PROCESSING SUMMARY")
    logger_debug("PROCESSING SUMMARY")
    logger.info(f"{'='*60}")
    logger_debug("="*60)
    logger.info(f"Total successful rows: {len(successful_rows)}")
    logger_debug(f"Total successful rows: {len(successful_rows)}")
    logger.info(f"Total failed rows: {len(failed_rows)}")
    logger_debug(f"Total failed rows: {len(failed_rows)}")
    logger.info(f"Total skipped rows (after retries): {len(skipped_rows)}")
    logger_debug(f"Total skipped rows (after retries): {len(skipped_rows)}")
    logger.info(f"\n{'='*120}")
    logger_debug("="*120)
    
    if successful_rows:
        logger.info(f"\nSuccessful rows:")
        for row_info in successful_rows[-5:]:  # Show last 5 successful
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']}")
        if len(successful_rows) > 5:
            logger.info(f"  ... and {len(successful_rows) - 5} more")
    
    if failed_rows:
        logger.info(f"\nFailed rows:")
        for row_info in failed_rows:
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']} - {row_info['error']}")
    
    if skipped_rows:
        logger.info(f"\nSkipped rows (after max retries):")
        for row_info in skipped_rows:
            logger.info(f"  - Row {row_info['index']}: No. Docket {row_info['no_docket']} - {row_info['error']}")
    
    if last_success_info:
        logger.info(f"\nLast successful row: {last_success_info['index']} - No. Docket {last_success_info['no_docket']}")
    
    if last_failure_info:
        logger.info(f"Last failed row: {last_failure_info['index']} - No. Docket {last_failure_info['no_docket']}")
    
    logger.info(f"{'='*60}")

def initialize_components(excel_file_path):
    """Initialize Excel processor and web driver"""
    try:
        if not os.path.exists(excel_file_path):
            logger.error(f"Excel file not found: {excel_file_path}")
            return None, None
        
        excel_processor = ExcelDataProcessor(excel_file_path)
        driver = setup_driver()
        
        return driver, excel_processor
    except Exception as e:
        logger.error(f"Failed to initialize components: {e}")
        return None, None


def process_all_rows(driver, wait, excel_processor, stop_event=None):
    """
    Process all rows from Excel with proper tracking and stop event support
    
    Args:
        driver: Selenium WebDriver instance
        wait: WebDriverWait instance
        excel_processor: ExcelDataProcessor instance
        stop_event: threading.Event object for stopping the process (optional)
    
    Returns:
        dict: Results containing successful_rows, failed_rows, skipped_rows, and stopped status
    """
    results = {
        'successful_rows': [],
        'failed_rows': [],
        'skipped_rows': [],
        'last_success_info': None,
        'last_failure_info': None,
        'stopped': False
    }
    
    total_rows = len(excel_processor.data)
    logger.info(f"Starting to process {total_rows} rows from Excel")
    
    row_index = 0
    while row_index < total_rows:
        # CHECK STOP EVENT at start of each row
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received at row {row_index + 1}/{total_rows}")
            results['stopped'] = True
            driver.quit()
            break
        
        row_data = excel_processor.get_row_data(row_index)
        
        if row_data is None:
            logger.warning(f"Skipping empty row {row_index + 1}")
            row_index += 1
            continue
        
        # Process current row
        no_docket = row_data.get('No. Docket', 'Unknown')
        log_row_header(row_index + 1, total_rows, no_docket)
        
        # CHECK STOP EVENT before processing
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received before processing row {row_index + 1}")
            results['stopped'] = True
            driver.quit()
            break
        
        success, processed_no_docket, error_message = process_excel_row_with_retry(
            driver, wait, excel_processor, row_data, row_index, stop_event=stop_event
        )
        
        # CHECK STOP EVENT after processing
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after processing row {row_index + 1}")
            results['stopped'] = True
            driver.quit()
            break
        
        if success:
            # Handle successful row processing
            row_index = handle_successful_row(
                driver, wait, excel_processor, results, 
                row_index, total_rows, processed_no_docket, stop_event
            )
        else:
            # Handle failed row processing
            handle_failed_row(
                driver, wait, results, row_index, 
                processed_no_docket, error_message
            )
        
        row_index += 1
        
        # CHECK STOP EVENT before sleep
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after row {row_index}")
            results['stopped'] = True
            driver.quit()
            break
            
        time.sleep(2)
    
    # Final log if stopped
    if results['stopped']:
        logger.warning(f"Processing stopped at row {row_index}/{total_rows}")
        logger.warning(f"Completed: {len(results['successful_rows'])} rows")
        logger.warning(f"Failed: {len(results['failed_rows'])} rows")
    
    return results


def handle_successful_row(driver, wait, excel_processor, results, 
                         row_index, total_rows, processed_no_docket, stop_event=None):
    """Handle successful row processing and potential duplicates with stop event support"""
    success_info = create_row_info(row_index + 1, processed_no_docket)
    results['successful_rows'].append(success_info)
    results['last_success_info'] = success_info
    
    logger.info(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")
    logger_debug(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")
    
    # CHECK STOP EVENT before handling next row
    if stop_event and stop_event.is_set():
        logger.warning(f"STOP signal received after successful row {row_index + 1}")
        results['stopped'] = True
        driver.quit()
        return row_index
    
    # Handle next row preparation and potential duplicates
    if row_index + 1 < total_rows:
        return handle_next_row_preparation(
            driver, wait, excel_processor, results, 
            row_index, total_rows, stop_event
        )
    else:
        logger.info("This is the last row - no next action needed")
    
    return row_index


def handle_next_row_preparation(driver, wait, excel_processor, results, 
                               row_index, total_rows, stop_event=None):
    """Handle preparation for next row and potential duplicate processing with stop event support"""
    
    # CHECK STOP EVENT before preparing next row
    if stop_event and stop_event.is_set():
        logger.warning(f"STOP signal received before preparing row {row_index + 2}")
        results['stopped'] = True
        driver.quit()
        return row_index
    
    next_action = prepare_for_next_row(driver, wait, excel_processor, row_index)
    
    if next_action == "duplicate":
        return process_duplicate_sequence(
            driver, wait, excel_processor, results, 
            row_index, total_rows, stop_event
        )
    elif next_action == "error":
        logger.error(f"Error preparing for next row after {row_index + 1}")
    
    return row_index


def process_duplicate_sequence(driver, wait, excel_processor, results, 
                              row_index, total_rows, stop_event=None):
    """Process sequence of duplicate rows with stop event support"""
    current_row = row_index
    
    while current_row + 1 < total_rows:
        # CHECK STOP EVENT at start of duplicate loop
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received during duplicate sequence at row {current_row + 1}")
            results['stopped'] = True
            driver.quit()
            break
        
        current_row += 1
        next_row_data = excel_processor.get_row_data(current_row)
        
        if next_row_data is None:
            logger.warning(f"Next row {current_row + 1} is empty - skipping duplicate")
            break
        
        next_no_docket = next_row_data.get('No. Docket', 'Unknown')
        log_duplicate_header(current_row + 1, total_rows, next_no_docket)
        
        # CHECK STOP EVENT before processing duplicate
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received before duplicate row {current_row + 1}")
            results['stopped'] = True
            driver.quit()
            break
        
        duplicate_success, duplicate_no_docket, duplicate_error = process_duplicate_row_with_retry(
            driver, wait, next_row_data, current_row, stop_event=stop_event
        )
        
        # CHECK STOP EVENT after processing duplicate
        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after duplicate row {current_row + 1}")
            results['stopped'] = True
            driver.quit()
            break
        
        if duplicate_success:
            handle_successful_duplicate(results, current_row, duplicate_no_docket)
            
            # Check if there's another duplicate
            if current_row + 1 < total_rows:
                # CHECK STOP EVENT before checking next duplicate
                if stop_event and stop_event.is_set():
                    logger.warning(f"STOP signal received before checking next duplicate")
                    results['stopped'] = True
                    driver.quit()
                    break
                    
                next_action = prepare_for_next_row(driver, wait, excel_processor, current_row)
                if next_action != "duplicate":
                    break
            else:
                break
        else:
            handle_failed_duplicate(driver, wait, results, current_row, 
                                  duplicate_no_docket, duplicate_error)
            break
    
    return current_row


def handle_successful_duplicate(results, row_index, no_docket):
    """Handle successful duplicate processing"""
    success_info = create_row_info(row_index + 1, no_docket)
    results['successful_rows'].append(success_info)
    results['last_success_info'] = success_info
    
    logger.info(f"Row {row_index + 1} successfully processed via duplicate - No. Docket: {no_docket}")
    logger_debug(f"Row {row_index + 1} successfully processed via duplicate - No. Docket: {no_docket}")


def handle_failed_duplicate(driver, wait, results, row_index, no_docket, error_message):
    """Handle failed duplicate processing"""
    if is_max_retry_error(error_message):
        skipped_info = create_error_info(row_index + 1, no_docket, error_message)
        results['skipped_rows'].append(skipped_info)
        logger.warning(f"Row {row_index + 1} skipped after max retries - No. Docket: {no_docket}")
    else:
        failure_info = create_error_info(row_index + 1, no_docket, error_message)
        results['failed_rows'].append(failure_info)
        results['last_failure_info'] = failure_info
    
    log_failed_duplicate(row_index + 1, no_docket, error_message)
    refresh_and_wait(driver, wait)


def handle_failed_row(driver, wait, results, row_index, no_docket, error_message):
    """Handle failed row processing"""
    if is_max_retry_error(error_message):
        skipped_info = create_error_info(row_index + 1, no_docket, error_message)
        results['skipped_rows'].append(skipped_info)
        logger.warning(f"Row {row_index + 1} skipped after max retries - No. Docket: {no_docket}")
    else:
        failure_info = create_error_info(row_index + 1, no_docket, error_message)
        results['failed_rows'].append(failure_info)
        results['last_failure_info'] = failure_info
        input("Press Enter to close the browser...")
        driver.quit()

def cleanup_resources_no_input(driver):
    """Clean up resources without user input (for automated runs)"""
    if driver:
        logger.info("Closing browser...")
        driver.quit()

# Helper functions for better code organization
def create_row_info(index, no_docket):
    """Create row info dictionary"""
    return {'index': index, 'no_docket': no_docket}


def create_error_info(index, no_docket, error):
    """Create error info dictionary"""
    return {'index': index, 'no_docket': no_docket, 'error': error}


def is_max_retry_error(error_message):
    """Check if error is due to max retry attempts"""
    return "after" in error_message and "attempts" in error_message


def log_row_header(row_num, total_rows, no_docket):
    """Log row processing header"""
    logger.info(f"{'='*100}")
    logger.info(f"Processing Row {row_num}/{total_rows} - No. Docket: {no_docket}")
    logger.info(f"{'='*100}")
    logger_debug(f"{'='*100}")


def log_duplicate_header(row_num, total_rows, no_docket):
    """Log duplicate row processing header"""
    logger.info(f"\n{'='*50}")
    logger.info(f"Processing Row {row_num}/{total_rows} (via duplicate) - No. Docket: {no_docket}")
    logger.info(f"{'='*50}")


def log_failed_row(row_num, no_docket, error_message):
    """Log failed row processing"""
    logger.error(f"Row {row_num} processing failed - No. Docket: {no_docket} - Error: {error_message}")
    logger.error(f"Skipped processing row {row_num} - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")
    logger_debug(f"Row {row_num} Skipped - Processing failed - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")


def log_failed_duplicate(row_num, no_docket, error_message):
    """Log failed duplicate processing"""
    logger.error(f"Row {row_num} duplicate processing failed - No. Docket: {no_docket} - Error: {error_message}")
    logger.error(f"Skipped processing row {row_num} - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")
    logger_debug(f"Row {row_num} Skipped - Duplicate Processing failed - No. Docket: {no_docket} - Error: {'Tidak ada No. Docket / Sudah Pernah diinput'}")


# Configuration class for better maintainability
class ProcessingConfig:
    """Configuration class for processing parameters"""
    EXCEL_FILE_PATH = "data.xlsx"
    WAIT_TIMEOUT = 10
    PROCESSING_DELAY = 2
    ERROR_MESSAGE_DUPLICATE = "Tidak ada No. Docket / Sudah Pernah diinput"

def main():
    """Main execution function - original entry point"""
    driver = None
    excel_processor = None
    
    try:
        script_dir = Path(__file__).parent.parent
        excel_file_path = script_dir / 'Data Input Pengujian' / 'data.xlsx'
        
        # Call dengan stop_event=None untuk backward compatibility
        results = run_with_custom_path_and_stop(excel_file_path, stop_event=None)
        
        if results:
            logger.info("="*60)
            logger.info("📊 FINAL RESULTS:")
            logger.info(f"✅ Successful: {results['successful_rows']}")
            logger.info(f"Failed: {results['failed_rows']}")
            logger.info(f"⏭️ Skipped: {results.get('skipped_rows', 0)}")
            logger.info(f"Stopped: {results.get('stopped', False)}")
            logger.info("="*60)

    except Exception as e:
        logger.error(f"Unexpected error in main: {e}")

def run_with_custom_path_and_stop(excel_path, stop_event=None):
    """
    Run the process with stop event support - COMPLETE IMPLEMENTATION
    
    Args:
        excel_path: Path to Excel file
        stop_event: threading.Event object for stopping the process
    
    Returns:
        dict: Results with successful_rows, failed_rows, and stopped status
    """
    driver = None
    excel_processor = None
    
    try:
        logger.info("="*60)
        logger.info("Starting Rencana Benda Uji process with stop event support")
        logger.info(f"Excel file: {excel_path}")
        logger.info("="*60)
        
        # Initialize components
        driver, excel_processor = initialize_components(excel_path)
        if not driver or not excel_processor:
            logger.error("Failed to initialize components")
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': False,
                'error': 'Initialization failed'
            }
        
        # CHECK STOP EVENT sebelum login
        if stop_event and stop_event.is_set():
            logger.warning("STOP signal received before login")
            driver.quit()
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': True
            }
        
        wait = WebDriverWait(driver, 10)
        login(driver, wait)
        
        # CHECK STOP EVENT setelah login
        if stop_event and stop_event.is_set():
            logger.warning("STOP signal received after login")
            driver.quit()
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': True
            }
        
        # Process all rows dengan stop event
        results = process_all_rows(driver, wait, excel_processor, stop_event)
        
        # Format hasil untuk return
        return_results = {
            'successful_rows': len(results['successful_rows']),
            'failed_rows': len(results['failed_rows']),
            'skipped_rows': len(results['skipped_rows']),
            'stopped': results.get('stopped', False),
            'details': results
        }
        
        # Log summary
        log_processing_summary(
            results['successful_rows'], 
            results['failed_rows'], 
            results['skipped_rows'],
            results['last_success_info'], 
            results['last_failure_info']
        )
        
        if results.get('stopped'):
            logger.warning("="*60)
            logger.warning("PROCESS STOPPED BY USER")
            logger.warning("="*60)
        
        return return_results
        
    except Exception as e:
        logger.error(f"Fatal error in run_with_custom_path_and_stop: {e}")
        import traceback
        logger.error(traceback.format_exc())
        return {
            'successful_rows': 0,
            'failed_rows': 0,
            'stopped': False,
            'error': str(e)
        }
        
    finally:
        # Clean up Selenium driver
        if driver:
            try:
                logger.info("Closing Selenium driver...")
                driver.quit()
                logger.info("Selenium driver closed successfully")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")


if __name__ == "__main__":
    main()