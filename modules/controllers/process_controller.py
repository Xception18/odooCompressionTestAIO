"""Process controller - orchestrates services for the full pipeline"""
import logging
from dataclasses import dataclass
from typing import Callable, Optional
from modules.services import (
    copy_excel_data,
    get_bjdt_ids,
    generate_csv_files,
    upload_today_csv,
    ExcelResult,
    BjdtResult,
    CsvResult,
    DbResult
)

logger = logging.getLogger(__name__)

@dataclass
class ProcessSettings:
    """Settings for process pipeline"""
    # File paths
    rencana_path: str
    pengujian_path: str
    output_dir: str
    
    # Sheet names
    sheet_source: str = "Report Excel"
    sheet_odoo: str = "ODOO"
    sheet_pengujian: str = "pengujian"
    sheet_bjdt: str = "bjdt_id"
    
    # Database
    db_host: str = "localhost"
    db_port: str = "5432"
    db_user: str = "postgres"
    db_password: str = ""
    db_name: str = "alatujidb"
    db_table: str = "pengujian"
    
    # API
    bjdt_timeout: int = 10


@dataclass
class ProcessResult:
    """Result of full process pipeline"""
    success: bool
    message: str
    steps_completed: int = 0
    excel_result: Optional[ExcelResult] = None
    bjdt_result: Optional[BjdtResult] = None
    csv_result: Optional[CsvResult] = None
    db_result: Optional[DbResult] = None
    error: Optional[str] = None


class ProcessController:
    """
    Controller for orchestrating the full data processing pipeline.
    
    Emits events for UI updates via callbacks.
    """
    
    def __init__(
        self,
        settings: ProcessSettings,
        on_step_started: Optional[Callable[[str], None]] = None,
        on_step_completed: Optional[Callable[[str, bool], None]] = None,
        on_log: Optional[Callable[[str, str], None]] = None
    ):
        """
        Args:
            settings: Process settings
            on_step_started: Callback(step_name)
            on_step_completed: Callback(step_name, success)
            on_log: Callback(message, level)
        """
        self.settings = settings
        self.on_step_started = on_step_started or (lambda x: None)
        self.on_step_completed = on_step_completed or (lambda x, y: None)
        self.on_log = on_log or (lambda x, y: None)
    
    def run_full_process(self) -> ProcessResult:
        """
        Run complete pipeline: Copy → BJDT → CSV → Upload.
        
        Returns:
            ProcessResult with overall status
        """
        steps_completed = 0
        
        try:
            # Step 1: Copy Data
            self.on_step_started("Copy Data")
            self.on_log("Starting data copy process...", "INFO")
            
            excel_result = copy_excel_data(
                self.settings.rencana_path,
                self.settings.pengujian_path,
                self.settings.sheet_source,
                self.settings.sheet_odoo,
                self.settings.sheet_pengujian
            )
            
            if not excel_result.success:
                self.on_step_completed("Copy Data", False)
                self.on_log(f"Copy failed: {excel_result.message}", "ERROR")
                return ProcessResult(
                    success=False,
                    message=f"Step 1 failed: {excel_result.message}",
                    steps_completed=0,
                    excel_result=excel_result,
                    error=excel_result.error
                )
            
            self.on_step_completed("Copy Data", True)
            self.on_log(f"Copy completed: {excel_result.message}", "SUCCESS")
            steps_completed += 1
            
            # Step 2: Get BJDT IDs
            self.on_step_started("Get BJDT ID")
            self.on_log("Fetching BJDT IDs from API...", "INFO")
            
            bjdt_result = get_bjdt_ids(
                self.settings.pengujian_path,
                self.settings.sheet_bjdt,
                self.settings.bjdt_timeout
            )
            
            if not bjdt_result.success:
                self.on_step_completed("Get BJDT ID", False)
                self.on_log(f"BJDT fetch failed: {bjdt_result.message}", "ERROR")
                return ProcessResult(
                    success=False,
                    message=f"Step 2 failed: {bjdt_result.message}",
                    steps_completed=1,
                    excel_result=excel_result,
                    bjdt_result=bjdt_result,
                    error=bjdt_result.errors[0] if bjdt_result.errors else None
                )
            
            self.on_step_completed("Get BJDT ID", True)
            self.on_log(f"BJDT completed: {bjdt_result.message}", "SUCCESS")
            steps_completed += 1
            
            # Step 3: Generate CSV
            self.on_step_started("Generate CSV")
            self.on_log("Generating CSV files...", "INFO")
            
            csv_result = generate_csv_files(
                self.settings.pengujian_path,
                self.settings.sheet_pengujian,
                self.settings.sheet_odoo,
                self.settings.sheet_bjdt,
                self.settings.output_dir
            )
            
            if not csv_result.success:
                self.on_step_completed("Generate CSV", False)
                self.on_log(f"CSV generation failed: {csv_result.message}", "ERROR")
                return ProcessResult(
                    success=False,
                    message=f"Step 3 failed: {csv_result.message}",
                    steps_completed=2,
                    excel_result=excel_result,
                    bjdt_result=bjdt_result,
                    csv_result=csv_result,
                    error=csv_result.error
                )
            
            self.on_step_completed("Generate CSV", True)
            self.on_log(f"CSV completed: {csv_result.message}", "SUCCESS")
            steps_completed += 1
            
            # Step 4: Upload to Database
            self.on_step_started("Upload DB")
            self.on_log("Uploading to database...", "INFO")
            
            db_result = upload_today_csv(
                self.settings.output_dir,
                self.settings.db_host,
                self.settings.db_port,
                self.settings.db_user,
                self.settings.db_password,
                self.settings.db_name,
                self.settings.db_table
            )
            
            if not db_result.success:
                self.on_step_completed("Upload DB", False)
                if db_result.duplicate:
                    self.on_log("Upload failed: Duplicate data detected", "WARNING")
                else:
                    self.on_log(f"Upload failed: {db_result.message}", "ERROR")
                
                return ProcessResult(
                    success=False,
                    message=f"Step 4 failed: {db_result.message}",
                    steps_completed=3,
                    excel_result=excel_result,
                    bjdt_result=bjdt_result,
                    csv_result=csv_result,
                    db_result=db_result,
                    error=db_result.error
                )
            
            self.on_step_completed("Upload DB", True)
            self.on_log(f"Upload completed: {db_result.message}", "SUCCESS")
            steps_completed += 1
            
            # Success!
            return ProcessResult(
                success=True,
                message="Full process completed successfully",
                steps_completed=4,
                excel_result=excel_result,
                bjdt_result=bjdt_result,
                csv_result=csv_result,
                db_result=db_result
            )
            
        except Exception as e:
            logger.error(f"Unexpected error in full process: {str(e)}")
            self.on_log(f"Unexpected error: {str(e)}", "ERROR")
            
            return ProcessResult(
                success=False,
                message=f"Unexpected error: {str(e)}",
                steps_completed=steps_completed,
                error=str(e)
            )
    
    def run_copy_only(self) -> ExcelResult:
        """Run only Copy Data step"""
        self.on_step_started("Copy Data")
        self.on_log("Starting data copy...", "INFO")
        
        result = copy_excel_data(
            self.settings.rencana_path,
            self.settings.pengujian_path,
            self.settings.sheet_source,
            self.settings.sheet_odoo,
            self.settings.sheet_pengujian
        )
        
        self.on_step_completed("Copy Data", result.success)
        if result.success:
            self.on_log(f"Copy completed: {result.message}", "SUCCESS")
        else:
            self.on_log(f"Copy failed: {result.message}", "ERROR")
        
        return result
    
    def run_bjdt_only(self) -> BjdtResult:
        """Run only Get BJDT ID step"""
        self.on_step_started("Get BJDT ID")
        self.on_log("Fetching BJDT IDs...", "INFO")
        
        result = get_bjdt_ids(
            self.settings.pengujian_path,
            self.settings.sheet_bjdt,
            self.settings.bjdt_timeout
        )
        
        self.on_step_completed("Get BJDT ID", result.success)
        if result.success:
            self.on_log(f"BJDT completed: {result.message}", "SUCCESS")
        else:
            self.on_log(f"BJDT failed: {result.message}", "ERROR")
        
        return result
    
    def run_csv_only(self) -> CsvResult:
        """Run only Generate CSV step"""
        self.on_step_started("Generate CSV")
        self.on_log("Generating CSV files...", "INFO")
        
        result = generate_csv_files(
            self.settings.pengujian_path,
            self.settings.sheet_pengujian,
            self.settings.sheet_odoo,
            self.settings.sheet_bjdt,
            self.settings.output_dir
        )
        
        self.on_step_completed("Generate CSV", result.success)
        if result.success:
            self.on_log(f"CSV completed: {result.message}", "SUCCESS")
        else:
            self.on_log(f"CSV failed: {result.message}", "ERROR")
        
        return result
    
    def run_upload_only(self, csv_path: Optional[str] = None) -> DbResult:
        """
        Run only Upload DB step.
        
        Args:
            csv_path: Optional specific CSV path (otherwise uses today's)
        """
        self.on_step_started("Upload DB")
        self.on_log("Uploading to database...", "INFO")
        
        if csv_path:
            from modules.services import upload_csv_to_db
            result = upload_csv_to_db(
                csv_path,
                self.settings.db_host,
                self.settings.db_port,
                self.settings.db_user,
                self.settings.db_password,
                self.settings.db_name,
                self.settings.db_table
            )
        else:
            result = upload_today_csv(
                self.settings.output_dir,
                self.settings.db_host,
                self.settings.db_port,
                self.settings.db_user,
                self.settings.db_password,
                self.settings.db_name,
                self.settings.db_table
            )
        
        self.on_step_completed("Upload DB", result.success)
        if result.success:
            self.on_log(f"Upload completed: {result.message}", "SUCCESS")
        else:
            if result.duplicate:
                self.on_log("Upload failed: Duplicate data", "WARNING")
            else:
                self.on_log(f"Upload failed: {result.message}", "ERROR")
        
        return result
