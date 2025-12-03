import configparser
import logging
import ssl
import threading
import time
import os
import sys
import queue

import psycopg2
import requests

# Global logger variable for the daemon
daemon_logger = None

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ThreadSafeLogHandler(logging.Handler):
    """Thread-safe log handler that queues messages instead of direct GUI access"""
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        log_entry = self.format(record)
        try:
            # Put log message in thread-safe queue
            self.log_queue.put_nowait(log_entry)
        except queue.Full:
            # If queue is full, skip this log entry
            pass


def setup_daemon_logging(log_queue=None, log_level=logging.INFO, log_file=None):
    """
    Setup logging for daemon with thread-safe queue integration
    
    Args:
        log_queue: Thread-safe queue.Queue() for GUI logging (replaces direct widget access)
        log_level: Logging level (default: INFO)
        log_file: File path for logging to file (optional)
    
    Returns:
        Configured logger instance
    """
    global daemon_logger
    
    # Create logger
    daemon_logger = logging.getLogger('DaemonSink')
    daemon_logger.setLevel(log_level)
    
    # Clear existing handlers to prevent duplicates
    for handler in daemon_logger.handlers[:]:
        daemon_logger.removeHandler(handler)
    
    # Add thread-safe queue handler if queue provided
    if log_queue:
        queue_handler = ThreadSafeLogHandler(log_queue)
        queue_handler.setFormatter(
            logging.Formatter('%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s')
        )
        daemon_logger.addHandler(queue_handler)
    
    # Add file handler if file path provided
    if log_file:
        try:
            file_handler = logging.FileHandler(log_file)
            file_handler.setFormatter(
                logging.Formatter('%(asctime)s [%(levelname)s] [%(threadName)s] (%(module)s:%(lineno)d) %(message)s')
            )
            daemon_logger.addHandler(file_handler)
        except Exception as e:
            print(f"Warning: Could not setup file logging: {e}")
    
    # Add console handler as fallback
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(
        logging.Formatter('%(asctime)s [%(levelname)s] [%(threadName)s] %(message)s')
    )
    daemon_logger.addHandler(console_handler)
    
    return daemon_logger


def get_daemon_logger():
    """Get the daemon logger instance"""
    global daemon_logger
    if daemon_logger is None:
        daemon_logger = setup_daemon_logging()
    return daemon_logger


class threadSinkData(threading.Thread):
    """
    Thread-safe daemon for syncing local database with ERP database
    
    CRITICAL CHANGES:
    - Removed direct log_widget parameter (causes GUI freezing)
    - Added log_queue for thread-safe logging
    - Improved stop mechanism with threading.Event
    - Better error handling and graceful shutdown
    """

    def __init__(self, threadID, name, delayNya, log_queue=None, logger=None):
        """
        Initialize daemon thread
        
        Args:
            threadID: Thread identifier
            name: Thread name
            delayNya: Delay between sync cycles in seconds
            log_queue: Thread-safe queue.Queue() for logging (NOT a GUI widget)
            logger: Optional logger instance
        """
        threading.Thread.__init__(self)
        self.daemon = True  # Make this a daemon thread
        
        self.threadID = threadID
        self.name = name
        self.delayNya = delayNya
        
        # Setup logger - use provided logger or get global daemon logger
        if logger:
            self.logger = logger
        elif log_queue:
            # Setup logger with thread-safe queue
            self.logger = setup_daemon_logging(log_queue=log_queue, log_file="daemonSync.log")
        else:
            self.logger = get_daemon_logger()
        
        # Thread control
        self._stop_event = threading.Event()
        self.stop_flag = False  # Backward compatibility
        
        self.logger.info("=" * 60)
        self.logger.info("Inisialisasi Daemon Sinkronisasi")
        self.logger.info(f"Thread ID: {threadID}, Name: {name}, Delay: {delayNya}s")
        self.logger.info("=" * 60)
    
    def stop(self):
        """Set stop flag to gracefully stop the daemon"""
        self.logger.info("=" * 60)
        self.logger.info("STOP SIGNAL RECEIVED")
        self.logger.info("Daemon will stop after current cycle completes")
        self.logger.info("=" * 60)
        
        self._stop_event.set()
        self.stop_flag = True
    
    def is_stopped(self):
        """Check if stop was requested"""
        return self._stop_event.is_set()
    
    def sleep_interruptible(self, seconds):
        """
        Sleep that can be interrupted by stop signal
        
        Args:
            seconds: Number of seconds to sleep
        
        Returns:
            True if sleep completed, False if interrupted
        """
        # Break sleep into 0.5 second chunks for faster response to stop
        chunks = int(seconds * 2)
        for _ in range(chunks):
            if self.is_stopped():
                return False
            time.sleep(0.5)
        return True

    def run(self):
        """Main daemon execution loop"""
        siklusUmum = 1
        
        self.logger.info("=" * 60)
        self.logger.info("DAEMON STARTED")
        self.logger.info(f"Thread: {self.name} is now running")
        self.logger.info("=" * 60)
        
        while not self.is_stopped():
            try:
                self.logger.debug(f"Memulai siklus thread ThreadSinkron yang ke-{siklusUmum}")
                self.logger.info(f"╔{'═' * 58}╗")
                self.logger.info(f"║ Siklus Thread ke-{siklusUmum:3d} - {self.name:40s}║")
                self.logger.info(f"╚{'═' * 58}╝")
                
                counterCekData = 1
                data = None
                
                # Query Data Benda pada database lokal dengan kondisi field sinkron = 'B'
                while not data and not self.is_stopped():
                    self.logger.debug(f"Melakukan pemeriksaan data ke-{counterCekData}")
                    self.logger.info(f"Cek data ke-{counterCekData}")
                    
                    data = cekData()
                    self.logger.debug(f"Hasil query cekData(): {data}")
                    
                    if not data:
                        self.logger.warning("Tidak ada data yang harus disinkronkan")
                        self.logger.info(f"Sleep {self.delayNya}s sebelum cek data berikutnya...")
                        counterCekData += 1
                        
                        # Interruptible sleep
                        if not self.sleep_interruptible(self.delayNya):
                            self.logger.info("Sleep interrupted by stop signal")
                            break
                    else:
                        self.logger.debug(f"Data ditemukan: {data}")
                        self.logger.info(f"Data yang harus disinkronkan: Docket {data[5]}, Urut {data[6]}")

                if self.is_stopped():
                    break

                # Proses pengiriman data
                respon = 0
                counterPost = 1
                
                while respon != 200 and not self.is_stopped():
                    self.logger.info(f"Percobaan ke-{counterPost} mengirim data ke ERP")
                    self.logger.debug(f"Memanggil kirimDataPost() dengan data: {data}")
                    
                    responKode = kirimDataPost(data)
                    respon = responKode.status_code if responKode else 0
                    
                    self.logger.debug(f"Respon server: {respon}")
                    
                    if respon != 200:
                        self.logger.warning(f"Respon error: {respon}")
                        self.logger.warning("Data tidak berhasil dikirim ke Database ERP")
                        self.logger.info(f"Retry setelah {self.delayNya}s...")
                        counterPost += 1
                        
                        # Interruptible sleep
                        if not self.sleep_interruptible(self.delayNya):
                            self.logger.info("Sleep interrupted by stop signal")
                            break
                    else:
                        self.logger.info("Respon OK - Data berhasil disinkronkan dengan ERP")
                        self.logger.info("Memulai proses update database lokal...")

                if self.is_stopped():
                    break

                # Proses update database lokal
                countCekDb = 1
                cekdb = data[7] if data else None
                self.logger.debug(f"Status sinkron awal: {cekdb}")
                
                while cekdb == "B" and not self.is_stopped():
                    self.logger.info(f"Update database lokal ke-{countCekDb}")
                    self.logger.info(f"Set status='S' untuk Docket: {data[5]}, Urut: {data[6]}") # type: ignore
                    
                    sinkUpdateLokal(data[5], data[6]) # type: ignore
                    self.logger.debug(f"Update status untuk bjdt_id: {data[4]}") # type: ignore
                    
                    # Verifikasi update berhasil
                    self.logger.debug("Verifikasi update dengan cekSinkronLokal()")
                    cekdb = cekSinkronLokal(data[4]) # pyright: ignore[reportOptionalSubscript]
                    self.logger.debug(f"Status sinkron setelah update: {cekdb}")
                    
                    if cekdb == "S":
                        self.logger.info(f"Update berhasil - Status sekarang: {cekdb}")
                    else:
                        self.logger.warning(f"Update mungkin gagal - Status: {cekdb}")
                    
                    countCekDb += 1
                    
                    # Small delay before next check
                    if not self.sleep_interruptible(1):
                        break
                
                self.logger.info(f"{'─' * 60}")
                self.logger.info(f"Siklus {siklusUmum} selesai")
                self.logger.info(f"{'─' * 60}")
                    
            except Exception as e:
                self.logger.error(f"ERROR pada siklus {siklusUmum}: {str(e)}")
                self.logger.exception("Exception details:")
                
                # Wait before retry on error
                self.logger.info(f"Menunggu {self.delayNya}s sebelum retry...")
                if not self.sleep_interruptible(self.delayNya):
                    break

            if self.is_stopped():
                break

            # Interruptible sleep between cycles
            self.logger.info(f"Sleep {self.delayNya}s sebelum siklus berikutnya...")
            if not self.sleep_interruptible(self.delayNya):
                break
            
            siklusUmum += 1
        
        self.logger.info("=" * 60)
        self.logger.info("DAEMON STOPPED GRACEFULLY")
        self.logger.info(f"Total siklus yang dijalankan: {siklusUmum - 1}")
        self.logger.info("=" * 60)


def cekData():
    """
    Check for data that needs to be synchronized (sinkron = 'B')
    
    Returns:
        Tuple of data if found, None otherwise
    """
    logger = get_daemon_logger()
    try:
        logger.debug("Memulai proses cekData()")
        config = configparser.RawConfigParser()
        fileConfig = resource_path("config.cnf")
        config.read(fileConfig)
        
        datab = config.get("data", "database")
        hosted = config.get("data", "host")
        login = config.get("data", "user")
        passed = config.get("data", "password")
        
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        logger.debug(f"Koneksi database: host={hosted}, database={datab}, user={login}")
        
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        
        SQL = """ SELECT tgluji, nilaikn, beratbenda, tiperetak, idbendauji, nodocket, nourutbenda, sinkron, bebanmpa, kuattekan, umur 
                  FROM pengujian
                  WHERE sinkron = 'B' 
                  LIMIT 1; """
        
        logger.debug("Executing SQL query for cekData()")
        kursor.execute(SQL)
        daftar = kursor.fetchone()
        
        kursor.close()
        konekdb.close()
        
        logger.debug(f"Data ditemukan: {daftar is not None}")
        return daftar
        
    except Exception as e:
        logger.error(f"Error pada cekData(): {str(e)}")
        logger.exception("Exception details in cekData():")
        return None


def cekSinkronLokal(bjdt_id):
    """
    Check synchronization status for specific bjdt_id
    
    Args:
        bjdt_id: Benda uji ID
        
    Returns:
        Status sinkron ('B' or 'S') or None on error
    """
    logger = get_daemon_logger()
    try:
        logger.debug(f"Memulai proses cekSinkronLokal() untuk ID: {bjdt_id}")
        
        config = configparser.RawConfigParser()
        fileConfig = resource_path("config.cnf")
        config.read(fileConfig)
        
        datab = config.get("data", "database")
        hosted = config.get("data", "host")
        login = config.get("data", "user")
        passed = config.get("data", "password")
        
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        
        SQL = """ SELECT sinkron FROM pengujian WHERE idbendauji = %s ; """
        data = (bjdt_id,)
        
        logger.debug(f"Executing cekSinkronLokal query for ID: {bjdt_id}")
        kursor.execute(SQL, data)
        daftar = kursor.fetchone()
        
        kursor.close()
        konekdb.close()
        
        result = daftar[0] if daftar else None
        logger.debug(f"Status sinkron untuk ID {bjdt_id}: {result}")
        return result
        
    except Exception as e:
        logger.error(f"Error pada cekSinkronLokal(): {str(e)}")
        logger.exception("Exception details in cekSinkronLokal():")
        return None


def sinkUpdateLokal(nomer, urut):
    """
    Update local database to mark data as synchronized
    
    Args:
        nomer: Nomor docket
        urut: Nomor urut benda
    """
    logger = get_daemon_logger()
    try:
        logger.debug(f"Memulai proses sinkUpdateLokal() untuk nomer: {nomer}, urut: {urut}")
        
        config = configparser.RawConfigParser()
        fileConfig = resource_path("config.cnf")
        config.read(fileConfig)
        
        datab = config.get("data", "database")
        hosted = config.get("data", "host")
        login = config.get("data", "user")
        passed = config.get("data", "password")
        
        conn = f"dbname={datab} user={login} host={hosted} password={passed}"
        konekdb = psycopg2.connect(conn)
        konekdb.autocommit = True
        kursor = konekdb.cursor()
        
        SQL = """ UPDATE pengujian SET sinkron = 'S' WHERE noDocket = %s AND noUrutBenda = %s; """
        data = (nomer, urut)
        
        logger.debug(f"Executing update query for noDocket: {nomer}, noUrutBenda: {urut}")
        kursor.execute(SQL, data)
        
        rows_affected = kursor.rowcount
        kursor.close()
        konekdb.close()
        
        logger.info(f"Update berhasil ({rows_affected} rows) - noDocket: {nomer}, noUrutBenda: {urut}")

    except Exception as e:
        logger.error(f"Error pada sinkUpdateLokal(): {str(e)}")
        logger.exception("Exception details in sinkUpdateLokal():")


def kirimDataPost(bendaUji):
    """
    Send data to ERP webservice
    
    Args:
        bendaUji: Tuple containing test data
        
    Returns:
        Response object or None on error
    """
    logger = get_daemon_logger()
    try:
        logger.debug("Memulai proses kirimDataPost()")
        
        fileConfig = resource_path("config.cnf")
        config = configparser.RawConfigParser()
        config.read(fileConfig)
        
        urlKirim = config.get("webser", "webser_hasilUji")
        
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        
        logger.debug(f"URL target: {urlKirim}")
        
        datanya = {"params": {}}
        datanya["params"]["bjdt_tgl_test"] = str(bendaUji[0])
        datanya["params"]["bjdt_beban"] = float(bendaUji[1])
        datanya["params"]["bjdt_berat"] = float(bendaUji[2])
        datanya["params"]["bjdt_tipe_retak"] = bendaUji[3]
        datanya["params"]["bjdt_id"] = int(bendaUji[4])
        datanya["params"]["bjdt_beban_mpa"] = float(bendaUji[8])
        datanya["params"]["bjdt_beban_kg"] = str(bendaUji[9])
        datanya["params"]["bjdt_umur"] = int(bendaUji[10])

        logger.debug(f"Data yang akan dikirim: {datanya}")
        logger.info("Mengirim data ke webservice ERP...")
        
        response = requests.post(urlKirim, json=datanya, timeout=30)
        logger.info(f"Response status code: {response.status_code}")
        logger.debug(f"Response content: {response.text[:200]}")  # First 200 chars
        
        return response
        
    except requests.Timeout:
        logger.error("Request timeout - Server tidak merespon dalam waktu 30 detik")
        return None
    except requests.HTTPError as e:
        logger.error(f"HTTP Error: {e}")
        return None
    except requests.ConnectionError as e:
        logger.error(f"Connection Error - Gagal menyambungkan ke server: {e}")
        return None
    except Exception as e:
        logger.error(f"Error pada kirimDataPost(): {str(e)}")
        logger.exception("Exception details in kirimDataPost():")
        return None


# Main execution code for standalone testing
if __name__ == "__main__":
    print("=" * 60)
    print("Starting Daemon in Standalone Mode")
    print("=" * 60)
    
    # Setup logging for standalone execution
    setup_daemon_logging(log_file="daemonSync.log", log_level=logging.DEBUG)
    
    fileConfig = resource_path("config.cnf")
    config = configparser.RawConfigParser()
    config.read(fileConfig)
    tunda = config.get("daemon", "delay") if config.has_option("daemon", "delay") else '2'
    
    logger = get_daemon_logger()
    logger.info("Starting daemon in standalone mode")
    
    daemon = threadSinkData(1, "ThreadSinkron", float(tunda))
    daemon.start()
    
    try:
        # Keep main thread alive
        while daemon.is_alive():
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nKeyboard interrupt received - stopping daemon...")
        daemon.stop()
        daemon.join(timeout=10)
        print("Daemon stopped")
