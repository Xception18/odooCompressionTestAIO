# ODOO Automation Script - Excel Integration (Playwright)
import os, sys, time, logging, re
from pathlib import Path
try:
    # pyrefly: ignore [missing-import]
    from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
except ImportError:  # pragma: no cover
    PlaywrightTimeoutError = TimeoutError  # type: ignore[assignment,misc]
    PlaywrightError = Exception  # type: ignore[assignment,misc]

# Import from modules
from modules.playwright_helpers import (
    BrowserWrapper,
    setup_driver,
    wait_for_loading_overlay_to_disappear,
    is_click_intercepted_error,
    fill_field,
    handle_duplicate_docket_error,
    quick_delete_excess_rows,
    no_record_found,
)
from modules.data_generators import (
    generate_random_slump_test,
    generate_random_yield,
    calculate_jam_sample,
)
from modules.excel_handler import ExcelDataProcessor
from modules.exceptions import DuplicateDocketException, NoRecordFoundException

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(funcName)s : %(lineno)d - %(message)s')
logger = logging.getLogger('Input_Rencana_Benda_Uji')
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


def login(driver: BrowserWrapper):
    """Handle login process"""
    from modules.config import get_config
    config = get_config()

    username = config.get_selenium_user()
    password = config.get_selenium_password()
    login_url = config.get_selenium_url()

    if not username or not password:
        raise ValueError("SELENIUM_USER and SELENIUM_PASSWORD environment variables are required")

    page = driver.page
    logger.info("Navigating to login page...")
    page.goto(login_url)
    logger.info("Entering credentials...")
    page.locator("#login").wait_for(state="visible")
    page.locator("#login").fill(username)
    page.locator("#password").fill(password)
    logger.info("Clicking login button...")
    page.locator("button[type='submit'].btn-primary:has-text('Log in')").click()
    time.sleep(1)
    logger.info(f"Current URL after login: {page.url}")


def navigate_and_create(driver: BrowserWrapper):
    """Navigate to create page and click create button"""
    wait_for_loading_overlay_to_disappear(driver)
    create_rbu_url = "https://rmc.adhimix.web.id/web#model=schedule.truck.mixer.benda.uji&view_type=list&cids=124&menu_id=669"
    logger.info("Navigating to CREATE RENCANA BENDA UJI page...")
    page = driver.page
    page.goto(create_rbu_url)
    wait_for_loading_overlay_to_disappear(driver)
    logger.info("CREATE RENCANA BENDA UJI...")
    page.get_by_role("button", name="New").click()


def fill_proyek_form(driver: BrowserWrapper, row_data=None):
    """Fill main form fields using Excel data"""
    # Date field - from Excel column 1 (index 0)
    wait_for_loading_overlay_to_disappear(driver)
    page = driver.page
    tgl_mulai_prod = str(row_data.iloc[0]) if len(row_data) > 0 else "None"
    logger.info(f"Filling Date form with: {tgl_mulai_prod}")
    tgl_field = page.locator("#mulai_produksi_docket_0")
    tgl_field.wait_for(state="visible")
    tgl_field.fill("")
    tgl_field.fill(tgl_mulai_prod)


def fill_docket_form(driver: BrowserWrapper, row_data=None):
    """Fill docket form using Excel data"""
    # No. Docket field - from Excel column 2 (index 1)
    wait_for_loading_overlay_to_disappear(driver)
    page = driver.page
    no_docket = str(row_data.iloc[1]) if len(row_data) > 1 else "None"
    logger.info(f"Filling No. Docket field with: {no_docket}")

    no_docket_field = page.locator("#rel_docket_0")
    no_docket_field.wait_for(state="visible")
    no_docket_field.scroll_into_view_if_needed()
    no_docket_field.click()
    no_docket_field.fill("")
    no_docket_field.fill(no_docket)

    try:
        no_record_found(driver, no_docket)
        # Check if autocomplete item with exact text exists
        xpath = f"//li[contains(@class, 'o-autocomplete--dropdown-item')]//a[contains(@class, 'ui-menu-item-wrapper') and normalize-space()='{no_docket}']"
        try:
            element = page.locator(f"xpath={xpath}")
            element.wait_for(state="visible")
            found = element.count() > 0 and element.is_visible()
        except PlaywrightTimeoutError:
            found = False

        if found:
            logger.info(f"Found {no_docket} in autocomplete dropdown")
            docket_link = page.locator("#rel_docket_0_0_0")
            docket_link.wait_for(state="visible", timeout=500)
            docket_link.click()
            handle_duplicate_docket_error(driver)
            logger.info(f"Successfully clicked on {no_docket} from autocomplete dropdown")
        else:
            logger.info(f"Element with text '{no_docket}' not found in autocomplete dropdown")
            logger.info("Using 'Search more...' option as fallback")
            search_more = page.locator("#rel_docket_0_0_1")
            search_more.wait_for(state="visible", timeout=500)
            search_more.click()
            selector = "tr.o_data_row td.o_data_cell"
            row = page.locator(selector).first
            row.wait_for(state="visible")
            row.click()
            handle_duplicate_docket_error(driver)

    except (DuplicateDocketException, NoRecordFoundException):
        raise
    except Exception as e:
        logger.error(f"Error in docket selection: {str(e)}")

    # Fill remaining fields using Excel data
    time.sleep(0.5)
    slump_value = str(row_data.iloc[6]) if len(row_data) > 6 else "12"  # Column 7 (index 6)

    # Extract number after "Slump" dynamically
    match = re.search(r'[Ss]lump\s+(\d+(?:\.\d+)?)', slump_value)
    slump_rencana = match.group(1) if match else "12"
    slump_test = generate_random_slump_test(slump_rencana)
    yield_value = generate_random_yield()
    nama_teknisi = str(row_data.iloc[4]) if len(row_data) > 4 else "TEKNISI"  # Column 5 (index 4)

    # Try to extract base_jam from web
    try:
        jam_berangkat_elem = page.locator("div[name='jam_berangkat'] span.text-truncate")
        base_jam = jam_berangkat_elem.inner_text().strip()
        logger.info(f"Extracted base_jam from web: {base_jam}")
    except Exception as e:
        logger.warning(f"Could not extract base_jam from web, falling back to Excel. Error: {e}")
        base_jam = str(row_data.iloc[8]) if len(row_data) > 8 else "10:30"  # Column 9 (index 8)
    jam_sample = calculate_jam_sample(base_jam)
    logger.info(f"Calculated jam_sample: {jam_sample}")

    # Fill remaining fields
    date_sample_field = page.locator("#date_sample_0")
    date_sample_field.wait_for(state="visible", timeout=5000)
    date_sample_field.click()
    date_sample_field.fill(jam_sample)
    page.keyboard.press("Tab")

    # Format HH:MM for jam_sample_0
    jam_to_input = jam_sample
    if ' ' in jam_sample:
        try:
            jam_to_input = jam_sample.split(' ')[1][:5]
        except:
            pass
    time.sleep(0.5)
    form_fields = [
        ("slump_rencana_0", slump_rencana, "Slump Rencana"),
        ("slump_test_0", slump_test, "Slump Test"),
        ("yield_test_0", yield_value, "Yield"),
        ("nama_teknisi_0", nama_teknisi, "Nama Teknisi"),
        ("jam_sample_0", jam_to_input, "Jam Sampel"),
    ]
    for field_id, value, field_name in form_fields:
        fill_field(driver, field_id, value, field_name)


def data_to_input(driver: BrowserWrapper, no_urut, row_data, is_first_row=False, stop_event=None):
    """Input data to the table row using Excel data with robust selectors"""
    page = driver.page
    rencana_umur_test = "7" if no_urut in [1, 2] else "28"
    kode_benda_uji = str(row_data.iloc[2]) if len(row_data) > 2 else f" Isi Kode Benda Uji - {no_urut}"
    bentuk_benda_uji = "Silinder 15 x 30 cm"
    logger.info(f"Input Row {no_urut} on data table...")
    time.sleep(1)

    # 1. FIND THE TARGET ROW
    row_selector = "div[name='benda_uji_ids'] table.o_list_table tbody tr.o_data_row"
    rows = page.locator(row_selector)

    if rows.count() == 0:
        logger.error("No rows found in benda_uji_ids table.")
        return

    # Select target row based on index
    idx = 0 if is_first_row else (no_urut - 1)
    count = rows.count()
    target_row = rows.nth(idx) if count > idx else rows.last

    target_row.scroll_into_view_if_needed()
    target_row.click()

    def fill_row_field(field_name, value, field_type='input'):
        try:
            cell_xpath = f"xpath=.//td[@name='{field_name}' or @data-name='{field_name}']"
            cell = target_row.locator(cell_xpath)
            cell.click()

            if field_type == 'select':
                # Odoo Selection field
                select_el = page.locator("select").first
                select_el.wait_for(state="visible")
                try:
                    select_el.select_option(label=value)
                except Exception:
                    # Fallback: try select by value substring
                    options = select_el.evaluate("el => Array.from(el.options).map(o => ({value: o.value, text: o.text}))")
                    for opt in options:
                        if value.lower() in opt['value'].lower():
                            select_el.select_option(value=opt['value'])
                            break

            elif field_type == 'autocomplete':
                # Odoo Many2one field
                input_el = cell.locator("input").first
                input_el.fill("")
                input_el.fill(str(value))
                time.sleep(0.8)
                page.keyboard.press("Enter")

            else:  # Standard Input
                input_el = cell.locator("input").first
                page.keyboard.press("Control+a")
                page.keyboard.press("Backspace")
                input_el.fill(str(value))
                page.keyboard.press("Tab")

        except Exception as e:
            logger.error(f"Error filling {field_name}: {e}")

    # 2. FILL DATA
    fill_row_field("nomor_urut", no_urut)
    fill_row_field("kode_benda_uji", kode_benda_uji)
    fill_row_field("benda_uji_id", rencana_umur_test, field_type='autocomplete')
    fill_row_field("bentuk_benda_uji", bentuk_benda_uji, field_type='autocomplete')
    fill_row_field("tempat_pengetesan", "Internal", field_type='select')
    #wait_for_loading_overlay_to_disappear(driver, wait=None)


def add_table_rows(driver: BrowserWrapper, row_data=None, stop_event=None):
    """Add and fill table rows using Excel data"""
    page = driver.page
    logger.info("Processing table rows...")

    existing_rows = page.locator("div[name='benda_uji_ids'] .o_data_row")
    existing_count = existing_rows.count()
    logger.info(f"Found {existing_count} existing rows")

    # Delete excess rows if more than 4 exist
    if existing_count > 4:
        logger.info(f"More than 4 rows exist ({existing_count}), deleting excess rows...")
        rows_to_delete = existing_count - 4
        quick_delete_excess_rows(driver, rows_to_delete)
        time.sleep(1)
        existing_count = page.locator("div[name='benda_uji_ids'] .o_data_row").count()

    # Get add item link
    try:
        add_item_link = page.locator("div[name='benda_uji_ids'] .o_field_x2many_list_row_add a")
        add_item_link.wait_for(state="visible")
    except PlaywrightTimeoutError:
        add_item_link = page.get_by_text("Add a line", exact=True)

    if existing_count < 4:
        logger.info("Less than 4 rows exist, filling existing rows first...")
        for no_urut in range(1, existing_count + 1):
            if stop_event and stop_event.is_set():
                return
            logger.info(f"Filling existing row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)

        for no_urut in range(existing_count + 1, 5):
            if stop_event and stop_event.is_set():
                return
            logger.info(f"Adding new row {no_urut}...")
            add_item_link.click()
            time.sleep(1)
            data_to_input(driver, no_urut, row_data, is_first_row=False, stop_event=stop_event)
    else:
        logger.info("Sufficient rows exist, filling first 4...")
        for no_urut in range(1, 5):
            if stop_event and stop_event.is_set():
                return
            logger.info(f"Filling row {no_urut}...")
            data_to_input(driver, no_urut, row_data, is_first_row=(no_urut == 1), stop_event=stop_event)


def save_form(driver: BrowserWrapper):
    """Save the form"""
    page = driver.page
    logger.info("Saving Rencana Benda Uji...")
    save_button = page.locator(".o_form_button_save")
    save_button.wait_for(state="visible", timeout=10000)
    logger.info("Save button found, clicking...")
    save_button.click()

def create_form(driver: BrowserWrapper):
    """Create new form"""
    page = driver.page
    logger.info("Creating new form...")
    page.get_by_role("button", name="New").click()

def duplicate_form(driver: BrowserWrapper, next_row_data=None, stop_event=None):
    """Duplicate form for next entry with same kode_benda_uji and proyek"""
    page = driver.page
    wait_for_loading_overlay_to_disappear(driver)
    logger.info("Click Duplicate Form...")
    page.locator("button:has(i[data-hotkey='u'])").click()
    page.get_by_role("menuitem", name="Duplicate").click()
    fill_docket_form(driver, next_row_data)
    logger.info("Input data to the table row using Excel data")
    if stop_event and stop_event.is_set():
        driver.quit()
        return
    add_table_rows(driver, next_row_data, stop_event=stop_event)


def alternative_form(driver: BrowserWrapper, next_row_data=None, stop_event=None):
    wait_for_loading_overlay_to_disappear(driver)
    logger.info("Updating duplicated form with next row data...")
    fill_docket_form(driver, next_row_data)
    logger.info("Input data to the table row using Excel data")
    if stop_event and stop_event.is_set():
        driver.quit()
        return
    add_table_rows(driver, next_row_data, stop_event=stop_event)


def refresh_and_wait(driver: BrowserWrapper):
    """Refresh page and wait for it to load"""
    logger.info("Refreshing page...")
    driver.page.reload()
    navigate_and_create(driver)
    wait_for_loading_overlay_to_disappear(driver)


def process_excel_row_with_retry(driver: BrowserWrapper, excel_processor=None, row_data=None, row_index=0, max_retries=3, stop_event=None):
    """Process single Excel row with retry logic for click intercepted errors"""
    no_docket = row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing Excel row {row_index + 1} - No. Docket: {no_docket}")

    for attempt in range(max_retries):
        try:
            if stop_event and stop_event.is_set():
                driver.quit()
                return False, no_docket, "Stopped by user"
            navigate_and_create(driver)

            wait_for_loading_overlay_to_disappear(driver)
            fill_proyek_form(driver, row_data)

            wait_for_loading_overlay_to_disappear(driver)
            fill_docket_form(driver, row_data)

            wait_for_loading_overlay_to_disappear(driver)
            if stop_event and stop_event.is_set():
                driver.quit()
                return False, no_docket, "Stopped by user"
            add_table_rows(driver, row_data, stop_event=stop_event)

            save_form(driver)
            logger.info(f"Success processing row {row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""

        except (DuplicateDocketException, NoRecordFoundException) as e:
            logger.warning(f"Duplicate or No Record Found detected for row {row_index + 1} (No. Docket: {no_docket}). Skipping...")
            refresh_and_wait(driver)
            return False, no_docket, str(e)

        except PlaywrightError as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {row_index + 1} (No. Docket: {no_docket})")
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver)
                    continue
                else:
                    logger.error(f"Max retries reached for row {row_index + 1} (No. Docket: {no_docket}). Skipping...")
                    return False, no_docket, f"Click intercepted after {max_retries} attempts"
            else:
                return False, no_docket, error_message

        except Exception as e:
            error_message = str(e)
            logger.warning(f"Click intercepted detected on attempt {attempt + 1}/{max_retries} for row {row_index + 1}")
            if attempt < max_retries - 1:
                logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                refresh_and_wait(driver)
                continue
            else:
                logger.error(f"Max retries reached for row {row_index + 1}. Skipping...")
                return False, no_docket, f"Click intercepted after {max_retries} attempts"

    return False, no_docket, "Unknown error after all retries"


def process_duplicate_row_with_retry(driver: BrowserWrapper, next_row_data=None, next_row_index=0, max_retries=3, stop_event=None):
    """Process next row using duplicate form with retry logic"""
    no_docket = next_row_data.get('No. Docket', 'Unknown')
    logger.info(f"Processing row {next_row_index + 1} using duplicate form - No. Docket: {no_docket}")

    use_duplicate = True

    for attempt in range(max_retries):
        try:
            if use_duplicate:
                logger.info(f"Using duplicate_form on attempt {attempt + 1}")
                duplicate_form(driver, next_row_data, stop_event=stop_event)
            else:
                logger.info(f"Using alternative_form on attempt {attempt + 1}")
                alternative_form(driver, next_row_data, stop_event=stop_event)

            save_form(driver)
            logger.info(f"Success processing row {next_row_index + 1}: No. Docket {no_docket}")
            return True, no_docket, ""

        except (DuplicateDocketException, NoRecordFoundException) as e:
            logger.warning(f"Duplicate or No Record Found detected for row {next_row_index + 1} (No. Docket: {no_docket}). Skipping...")
            refresh_and_wait(driver)
            return False, no_docket, str(e)

        except PlaywrightError as e:
            error_message = str(e)
            if is_click_intercepted_error(error_message):
                logger.warning(f"Click intercepted on attempt {attempt + 1}/{max_retries} for row {next_row_index + 1}")
                if attempt < max_retries - 1:
                    logger.info(f"Retrying after refresh for No. Docket: {no_docket}")
                    refresh_and_wait(driver)
                    navigate_and_create(driver)
                    fill_proyek_form(driver, next_row_data)
                    if use_duplicate:
                        use_duplicate = False
                        logger.info("Switching from duplicate_form to alternative_form strategy")
                    continue
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
                refresh_and_wait(driver)
                navigate_and_create(driver)
                fill_proyek_form(driver, next_row_data)
                if use_duplicate:
                    use_duplicate = False
                    logger.info("Exception occurred, switching from duplicate_form to alternative_form strategy")
                continue
            else:
                logger.error(f"Max retries reached for row {next_row_index + 1}. Skipping...")
                return False, no_docket, f"Click intercepted after {max_retries} attempts"

    return False, no_docket, "Unknown error after all retries"


def prepare_for_next_row(driver: BrowserWrapper, excel_processor=None, row_index=0):
    """Prepare for next row - either duplicate or create new form"""
    try:
        wait_for_loading_overlay_to_disappear(driver)
        if excel_processor.should_duplicate(row_index):
            logger.info("Next row has same kode_benda_uji and proyek - will use duplicate")
            return "duplicate"
        else:
            logger.info("Next row has different kode_benda_uji or proyek - creating new form")
            create_form(driver)
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
        for row_info in successful_rows[-5:]:
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


def process_all_rows(driver: BrowserWrapper, excel_processor=None, stop_event=None):
    """
    Process all rows from Excel with proper tracking and stop event support

    Args:
        driver: BrowserWrapper instance
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
        'stopped': False,
        'driver_closed_unexpectedly': False,
    }

    total_rows = len(excel_processor.data)
    logger.info(f"Starting to process {total_rows} rows from Excel")

    row_index = 0
    while row_index < total_rows:
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

        no_docket = row_data.get('No. Docket', 'Unknown')
        log_row_header(row_index + 1, total_rows, no_docket)

        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received before processing row {row_index + 1}")
            results['stopped'] = True
            driver.quit()
            break

        try:
            success, processed_no_docket, error_message = process_excel_row_with_retry(
                driver, excel_processor, row_data, row_index, stop_event=stop_event
            )
        except Exception as e:
            logger.error(f"Driver error while processing row {row_index + 1}: {e}")
            success = False
            processed_no_docket = no_docket
            error_message = f"Driver error: {str(e)}"
            results['driver_closed_unexpectedly'] = True

        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after processing row {row_index + 1}")
            results['stopped'] = True
            driver.quit()
            break

        if success:
            row_index = handle_successful_row(
                driver, excel_processor, results,
                row_index, total_rows, processed_no_docket, stop_event
            )
        else:
            handle_failed_row(
                driver, results, row_index,
                processed_no_docket, error_message
            )

        row_index += 1

        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after row {row_index}")
            results['stopped'] = True
            driver.quit()
            break

        time.sleep(2)

    if row_index < total_rows and (not results['stopped'] or results['driver_closed_unexpectedly']):
        logger.warning(f"Driver closed or error occurred - marking {total_rows - row_index} remaining rows as failed")
        for remaining_row_idx in range(row_index, total_rows):
            remaining_row_data = excel_processor.get_row_data(remaining_row_idx)
            if remaining_row_data is not None:
                remaining_no_docket = remaining_row_data.get('No. Docket', 'Unknown')
                failure_info = create_error_info(
                    remaining_row_idx + 1,
                    remaining_no_docket,
                    "Driver closed before row could be processed",
                )
                results['failed_rows'].append(failure_info)
                logger.error(f"Row {remaining_row_idx + 1} marked as failed - No. Docket: {remaining_no_docket} - Driver closed unexpectedly")

    if results['stopped']:
        logger.warning(f"Processing stopped at row {row_index}/{total_rows}")
        logger.warning(f"Completed: {len(results['successful_rows'])} rows")
        logger.warning(f"Failed: {len(results['failed_rows'])} rows")

    return results


def handle_successful_row(driver: BrowserWrapper, excel_processor, results,
                          row_index, total_rows, processed_no_docket, stop_event=None):
    """Handle successful row processing and potential duplicates with stop event support"""
    success_info = create_row_info(row_index + 1, processed_no_docket)
    results['successful_rows'].append(success_info)
    results['last_success_info'] = success_info

    logger.info(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")
    logger_debug(f"Row {row_index + 1} successfully saved - No. Docket: {processed_no_docket}")

    if stop_event and stop_event.is_set():
        logger.warning(f"STOP signal received after successful row {row_index + 1}")
        results['stopped'] = True
        driver.quit()
        return row_index

    if row_index + 1 < total_rows:
        return handle_next_row_preparation(
            driver, excel_processor, results,
            row_index, total_rows, stop_event
        )
    else:
        logger.info("This is the last row - no next action needed")

    return row_index


def handle_next_row_preparation(driver: BrowserWrapper, excel_processor, results,
                                row_index, total_rows, stop_event=None):
    """Handle preparation for next row and potential duplicate processing with stop event support"""
    if stop_event and stop_event.is_set():
        logger.warning(f"STOP signal received before preparing row {row_index + 2}")
        results['stopped'] = True
        driver.quit()
        return row_index

    next_action = prepare_for_next_row(driver, excel_processor, row_index)

    if next_action == "duplicate":
        return process_duplicate_sequence(
            driver, excel_processor, results,
            row_index, total_rows, stop_event
        )
    elif next_action == "error":
        logger.error(f"Error preparing for next row after {row_index + 1}")

    return row_index


def process_duplicate_sequence(driver: BrowserWrapper, excel_processor, results,
                               row_index, total_rows, stop_event=None):
    """Process sequence of duplicate rows with stop event support"""
    current_row = row_index

    while current_row + 1 < total_rows:
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

        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received before duplicate row {current_row + 1}")
            results['stopped'] = True
            driver.quit()
            break

        duplicate_success, duplicate_no_docket, duplicate_error = process_duplicate_row_with_retry(
            driver, next_row_data, current_row, stop_event=stop_event
        )

        if stop_event and stop_event.is_set():
            logger.warning(f"STOP signal received after duplicate row {current_row + 1}")
            results['stopped'] = True
            driver.quit()
            break

        if duplicate_success:
            handle_successful_duplicate(results, current_row, duplicate_no_docket)

            if current_row + 1 < total_rows:
                if stop_event and stop_event.is_set():
                    logger.warning(f"STOP signal received before checking next duplicate")
                    results['stopped'] = True
                    driver.quit()
                    break
                next_action = prepare_for_next_row(driver, excel_processor, current_row)
                if next_action != "duplicate":
                    break
            else:
                break
        else:
            handle_failed_duplicate(driver, results, current_row,
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


def handle_failed_duplicate(driver: BrowserWrapper, results, row_index, no_docket, error_message):
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
    refresh_and_wait(driver)


def handle_failed_row(driver: BrowserWrapper, results, row_index, no_docket, error_message):
    """Handle failed row processing"""
    if is_max_retry_error(error_message) or "Duplicate Docket" in error_message or "Docket ini sudah pernah dibuat benda uji" in error_message or "No record found for" in error_message:
        skipped_info = create_error_info(row_index + 1, no_docket, error_message)
        results['skipped_rows'].append(skipped_info)
        logger.warning(f"Row {row_index + 1} skipped - No. Docket: {no_docket} - Reason: {error_message}")
    else:
        failure_info = create_error_info(row_index + 1, no_docket, error_message)
        results['failed_rows'].append(failure_info)
        results['last_failure_info'] = failure_info
        input("Press Enter to close the browser...")
        if driver:
            try:
                driver.quit()
            except Exception as e:
                logger.error(f"Error closing driver: {e}")


def cleanup_resources_no_input(driver: BrowserWrapper):
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
    try:
        script_dir = Path(__file__).parent.parent
        excel_file_path = script_dir / 'Data Input Pengujian' / 'data.xlsx'

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

        driver, excel_processor = initialize_components(excel_path)
        if not driver or not excel_processor:
            logger.error("Failed to initialize components")
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': False,
                'error': 'Initialization failed',
            }

        if stop_event and stop_event.is_set():
            logger.warning("STOP signal received before login")
            driver.quit()
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': True,
            }

        login(driver)

        if stop_event and stop_event.is_set():
            logger.warning("STOP signal received after login")
            driver.quit()
            return {
                'successful_rows': 0,
                'failed_rows': 0,
                'stopped': True,
            }

        results = process_all_rows(driver, excel_processor=excel_processor, stop_event=stop_event)

        return_results = {
            'successful_rows': len(results['successful_rows']),
            'failed_rows': len(results['failed_rows']),
            'skipped_rows': len(results['skipped_rows']),
            'stopped': results.get('stopped', False),
            'details': results,
        }

        log_processing_summary(
            results['successful_rows'],
            results['failed_rows'],
            results['skipped_rows'],
            results['last_success_info'],
            results['last_failure_info'],
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
            'error': str(e),
        }

    finally:
        if driver:
            try:
                logger.info("Closing Playwright browser...")
                driver.quit()
                logger.info("Playwright browser closed successfully")
            except Exception as e:
                logger.error(f"Error closing driver: {e}")


if __name__ == "__main__":
    main()