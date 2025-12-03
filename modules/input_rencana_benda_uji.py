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
    password = "rmc"
    login_url = "https://rmc.adhimix.web.id/web/login"
    logger.info("Navigating to login page...")
    driver.get(login_url)
    logger.info("Entering credentials...")
    username_field = wait.until(EC.presence_of_element_located((By.ID, "login")))
    password_field = driver.find_element(By.ID, "password")
    username_field.send_keys(username)
    password_field.send_keys(password)
    logger.info("Clicking login button...")
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".btn.btn-primary")))
    login_button.click()
    time.sleep(3)
    logger.info(f"Current URL after login: {driver.current_url}")

def navigate_and_create(driver, wait):
    wait_for_loading_overlay_to_disappear(driver, wait)
    """Navigate to create page and click create button"""
    create_rbu_url = "https://rmc.adhimix.web.id/web?#min=1&limit=80&view_type=list&model=schedule.truck.mixer.benda.uji&menu_id=535"
    logger.info("Navigating to CREATE RENCANA BENDA UJI page...")
    driver.get(create_rbu_url)
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("CREATE RENCANA BENDA UJI...")
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/button[1]")))
    create_button.click()

def fill_proyek_form(driver, wait, row_data):
    """Fill main form fields using Excel data"""
    # Date field - from Excel column 1 (index 0)
    wait_for_loading_overlay_to_disappear(driver, wait)
    tgl_mulai_prod = str(row_data.iloc[0]) if len(row_data) > 0 else "None"
    logger.info(f"Filling Date form with: {tgl_mulai_prod}")
    tgl_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[1]/td[2]/div/input")))
    tgl_field.clear()
    tgl_field.send_keys(tgl_mulai_prod)
    # Proyek field - from Excel column 4 (index 3)
    proyek = str(row_data.iloc[3]) if len(row_data) > 3 else "-"
    logger.info(f"Filling Proyek field with: {proyek}")
    proyek_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[2]/td[2]/div/div/input")))
    proyek_field.clear()
    proyek_field.send_keys(proyek)
    time.sleep(3)
    
    if proyek == "JALAN TOL AKSES PATIMBAN":
        proyek_option = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[1]/li[2]")))
        proyek_option.click()
        logger.info("Selected 'JALAN TOL AKSES PATIMBAN' from dropdown")
    else:
        try:
            dropdown_option = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[1]/li[1]")))
            dropdown_option.click()
            logger.info("Proyek selected")
        except TimeoutException:
            logger.error("Proyek dropdown not found")

def fill_docket_form(driver, wait, row_data):
    """Fill docket form using Excel data"""
    # No. Docket field - from Excel column 2 (index 1)
    wait_for_loading_overlay_to_disappear(driver, wait)
    no_docket = str(row_data.iloc[1]) if len(row_data) > 1 else "None"
    logger.info(f"Filling No. Docket field with: {no_docket}")
    # Click and fill the No. Docket field
    no_docket_field = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[1]/tbody/tr[3]/td[2]/div/div/input")))
    driver.execute_script("arguments[0].scrollIntoView(true);", no_docket_field)
    time.sleep(1)
    no_docket_field.click()
    no_docket_field.clear()
    no_docket_field.send_keys(no_docket)
    time.sleep(3)
    
    try:
        time.sleep(1)
        # Check if the element exists without waiting
        xpath = f"//ul[@class='ui-autocomplete ui-front ui-menu ui-widget ui-widget-content']//a[text()='{no_docket}']"
        elements = driver.find_elements(By.XPATH, xpath)
        
        if elements and elements[0].is_displayed():
            # Element found, click it
            logger.info(f"Found {no_docket} in autocomplete dropdown")
            elements[0].click()
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
    base_jam = str(row_data.iloc[8]) if len(row_data) > 8 else "10:30"  # Column 9 (index 8)
    jam_sample = calculate_jam_sample(base_jam)
    base_xpath = "/html/body/div[1]/div/div[2]/div/div/div/div/div[1]/table[2]/tbody/tr"
    form_fields = [
        (f"{base_xpath}[2]/td[2]/input", slump_rencana, "Slump Rencana"),
        (f"{base_xpath}[3]/td[2]/input", slump_test, "Slump Test"),
        (f"{base_xpath}[4]/td[2]/input", yield_value, "Yield"),
        (f"{base_xpath}[5]/td[2]/input", nama_teknisi, "Nama Teknisi"),
        (f"{base_xpath}[6]/td[2]/input", jam_sample, "Jam Sample")
    ]
    for xpath, value, field_name in form_fields:
        time.sleep(1)
        fill_field(driver, wait, xpath, value, field_name)

def data_to_input(driver, no_urut, row_data, is_first_row=False, stop_event=None):
    """Input data to the table row using Excel data"""
    wait = WebDriverWait(driver, 5)
    # Determine test age based on sequence number
    rencana_umur_test = "7" if no_urut in [1, 2] else "28"
    # Get kode benda uji from Excel (column 3, index 2)
    kode_benda_uji = str(row_data.iloc[2]) if len(row_data) > 2 else f" Isi Kode Benda Uji - {no_urut}"
    bentuk_benda_uji = "Silinder 15 x 30 cm"
    logger.info(f"Input Row {no_urut} on data table...")
    time.sleep(1)
    # Only click on the first row if it's the first iteration
    if is_first_row:
        first_row = wait.until(EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_'] td.o_list_number[data-field='nomor_urut']")
                ))
        first_row.click()

    else:
        # For rows 2-4, try to find and click the specific row
        try:
            # Try to find the row with the specific number
            rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_'] td.o_list_number[data-field='nomor_urut']")
            if len(rows) >= no_urut:
                target_row = rows[no_urut - 1]  # Index is 0-based, no_urut is 1-based
                target_row.click()
            else:
                logger.warning(f"Row {no_urut} not found, using first available row")
                target_row = rows[0] if rows else None
                if target_row:
                    target_row.click()
        except Exception as e:
            logger.warning(f"Error finding row {no_urut}: {e}")
    # Fill all fields for the row
    base_xpath = "/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/div/div/table/tbody/tr/td/div/div[2]/div[1]"
    fields_data = [
        (f"{base_xpath}/input[1]", str(no_urut), "No Urut"),
        (f"{base_xpath}/input[2]", kode_benda_uji, "Kode Benda Uji"),
        (f"{base_xpath}/div[1]/div/input", rencana_umur_test, "Rencana Umur Test"),
    ]
    for xpath, value, field_name in fields_data:
        if stop_event and stop_event.is_set():
            logger.warning("Stop signal received in data_to_input loop")
            driver.quit()
            return
        logger.info(f"Filling {field_name} with value: {value}")
        field = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", field)
        field.send_keys(Keys.CONTROL, "a")
        field.send_keys(Keys.DELETE)
        field.send_keys(str(value))
        time.sleep(1)

    wait_for_loading_overlay_to_disappear(driver, wait)

    if rencana_umur_test == "7":
        umur7 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[3]/li[1]/a")))
        umur7.click()
    else:
        umur28 = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[3]/li/a")))
        umur28.click()

    # Fill Bentuk Benda Uji
    logger.info(f"Filling Bentuk Benda Uji field for row {no_urut}...")
    bentuk_benda_uji_field = driver.find_element(By.CSS_SELECTOR, '[data-fieldname="bentuk_benda_uji"] .o_form_input')
    bentuk_benda_uji_field.click()
    bentuk_benda_uji_field.send_keys(Keys.CONTROL, "a")
    bentuk_benda_uji_field.send_keys(Keys.DELETE) 
    bentuk_benda_uji_field.send_keys(bentuk_benda_uji)
    time.sleep(1)
    silinder_select = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/ul[4]/li[1]/a")))
    silinder_select.click()
    time.sleep(1)

    # Fill Tempat Pengetesan
    logger.info(f"Filling Tempat Pengetesan field for row {no_urut}...")
    tempat_field = wait.until(EC.element_to_be_clickable((By.XPATH, f"{base_xpath}/select")))
    tempat_field.click()
    time.sleep(1)
    tempat_internal_select = wait.until(EC.element_to_be_clickable((By.XPATH, f"{base_xpath}/select/option[2]")))
    tempat_internal_select.click()
    time.sleep(1)

def add_table_rows(driver, wait, row_data, stop_event=None):
    """Add and fill table rows using Excel data"""
    logger.info("Processing table rows...")
    
    # Check how many existing rows with data-id^='one2many_v_id' exist
    existing_rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_']")
    existing_count = len(existing_rows)
    logger.info(f"Found {existing_count} existing rows")
    
    # Delete excess rows if more than 4 exist
    if existing_count > 4:
        logger.info(f"More than 4 rows exist ({existing_count}), deleting excess rows...")
        rows_to_delete = existing_count - 4
        deleted_count = quick_delete_excess_rows(driver, rows_to_delete)
        logger.info(f"Deleted {deleted_count} excess rows")
        time.sleep(1)
        # Recheck existing count after deletion
        existing_rows = driver.find_elements(By.CSS_SELECTOR, "tr[data-id^='one2many_v_id_']")
        existing_count = len(existing_rows)
        logger.info(f"After deletion: {existing_count} rows remain")
    
    # Get add item link
    add_item_link = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Add an item")))
    
    if existing_count < 4:
        # Fill existing rows first, then add new ones
        logger.info("Less than 4 rows exist, filling existing rows first...")
        
        # Fill existing rows in order
        for no_urut in range(1, existing_count + 1):
            if stop_event and stop_event.is_set():
                logger.warning("Stop signal received in add_table_rows loop 1")
                driver.quit()
                return
            logger.info(f"Filling existing row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)
        
        # Add and fill remaining rows using add item link
        for no_urut in range(existing_count + 1, 5):
            if stop_event and stop_event.is_set():
                logger.warning("Stop signal received in add_table_rows loop 2")
                driver.quit()
                return
            logger.info(f"Adding new row {no_urut}...")
            add_item_link.click()
            time.sleep(1)  # Wait for row to be added
            data_to_input(driver, no_urut, row_data, is_first_row=False, stop_event=stop_event)
    else:
        # If exactly 4 rows exist, just fill them
        logger.info("Exactly 4 rows exist, filling existing rows...")
        for no_urut in range(1, 5):
            if stop_event and stop_event.is_set():
                logger.warning("Stop signal received in add_table_rows loop 3")
                driver.quit()
                return
            logger.info(f"Filling row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)

def save_form(driver, wait):
    """Save the form"""
    time.sleep(2)
    tablist = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[2]/div/div/div/div/div[2]/ul")))
    tablist.click()
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Saving Rencana Benda Uji...")
    save_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/div[2]/button[1]")))
    logger.info("Save button found, clicking...")
    save_button.click()
    time.sleep(3)

def create_form(wait):
    """Create new form"""
    logger.info("Creating new form...")
    create_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[1]/div/div[1]/button[2]")))
    logger.info("Create button found, clicking...")
    create_button.click()

def duplicate_form(driver, wait, next_row_data, stop_event=None):
    """Duplicate form for next entry with same kode_benda_uji and proyek"""
    wait_for_loading_overlay_to_disappear(driver, wait)
    logger.info("Duplicating Rencana Benda Uji...")
    action_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/button")))
    action_button.click()
    time.sleep(1)
    duplicate_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div[2]/div[2]/div/div[2]/ul/li/a")))
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