import os
import subprocess
import requests
import threading
import logging
from colorama import Fore, Style, init
from datetime import datetime
import speedtest
import time
import ctypes
import sys
import psutil

import pyfiglet

# Initialize colorama
init(autoreset=True)
desktop_path="C:\\"


def is_onedrive_running():
    # OneDrive的进程名可能因版本而异，常见的是"OneDrive"或"OneDrive.exe"
    for proc in psutil.process_iter(['name']):
        if 'onedrive' in proc.info['name'].lower():
            return True
    return False

#find desktop
def get_desktop_path():
    try:

        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')

        if not os.path.exists(desktop_path):
            raise FileNotFoundError(f"Cant find desktop path: {desktop_path}")
        return desktop_path
    except KeyError as e:

        raise EnvironmentError(f"environment variable unavailable: {e}")
    except FileNotFoundError as e:

        raise e



# Logging configuration
def setup_logging():
    """Sets up logging to file on the desktop without color codes, and with color in the console."""
    global desktop_path
    if  is_onedrive_running():
        desktop_path="C:\\"
        print(f'Desktop path: {desktop_path}')
    else:
        try:
            desktop_path = get_desktop_path()
            print(f'Desktop path: {desktop_path}')
        except Exception as e:
            print(f"error: {e}")
    log_filename = os.path.join(desktop_path, f"quickfix-log-{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log")

    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    file_handler = logging.FileHandler(log_filename)
    stream_handler = logging.StreamHandler()

    file_format = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s', datefmt="%H:%M:%S")
    stream_format = logging.Formatter(f"{Fore.YELLOW}%(asctime)s{Fore.RESET} - %(levelname)s - %(message)s",
                                      datefmt="%H:%M:%S")

    file_handler.setFormatter(file_format)
    stream_handler.setFormatter(stream_format)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)

    logging.info("Logging setup complete. Log file is on the desktop.")


# Constants
APPDATA_PATH = os.path.join(os.getenv('APPDATA'), 'QuickFix')
SYSTEM32_PATH = "C:\\Windows\\System32"
BUGSPLAT_URL = "https://akiraghost.com/utils/BugSplat/"
VC_REDIST_URL = 'https://aka.ms/vs/17/release/vc_redist.x64.exe'
JAVA_URL = 'https://javadl.oracle.com/webapps/download/AutoDL?BundleId=249851_43d62d619be4e416215729597d70b8ac'
BUGSPLAT_FILES = ['BsSndRpt64.exe', 'BugSplat64.dll', 'BugSplatHD64.exe', 'BugSplatRc64.dll']

ASCII_ART = """
  ██████╗ ██╗   ██╗██╗ ██████╗██╗  ██╗███████╗██╗██╗  ██╗██╗
██╔═══██╗██║   ██║██║██╔════╝██║ ██╔╝██╔════╝██║╚██╗██╔╝██║
██║   ██║██║   ██║██║██║     █████╔╝ █████╗  ██║ ╚███╔╝ ██║
██║▄▄ ██║██║   ██║██║██║     ██╔═██╗ ██╔══╝  ██║ ██╔██╗ ╚═╝
╚██████╔╝╚██████╔╝██║╚██████╗██║  ██╗██║     ██║██╔╝ ██╗██╗
 ╚══▀▀═╝  ╚═════╝ ╚═╝ ╚═════╝╚═╝  ╚═╝╚═╝     ╚═╝╚═╝  ╚═╝╚═╝
"""
f = pyfiglet.Figlet(font='larry3d')
text = "QUICK FIX"


def clear_screen():
    """Clears the command line screen."""
    os.system('cls' if os.name == 'nt' else 'clear')


def delete_previous_logs():
    """Delete previous log files from the desktop."""
    #logging.info(f"{Fore.BLUE}Delete previous log files from the desktop.")
    # desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    for filename in os.listdir(desktop_path):
        if filename.startswith("quickfix-log-"):
            file_path = os.path.join(desktop_path, filename)
            try:
                os.remove(file_path)
                #logging.info(f"Deleted previous log file: {filename}")
            except OSError as e:
                #logging.error(f"Error deleting log file {filename}: {e}")
                pass


def user_consent():
    """Prompts the user for consent to proceed with the script's actions."""
    global all_test
    all_test = False
    disclaimer_text = (
        "By running this script, you acknowledge and consent to the following actions that the script will perform:\n"
        "- Terminate applications.\n"
        "- Install the latest Windows updates.\n"
        "- Update Java and Microsoft Visual C++ Redistributables.\n"
        "- Analyze your system.\n"
        "- Download additional files necessary for system repair and optimization.\n\n"
        f"These actions are intended to enhance and secure your system performance. {Fore.GREEN}Please type 'y' to continue, "
        f"{Fore.YELLOW}or type 'all' to enable all more time-consuming tests. "
        f"{Fore.RESET}If you do not consent, please close this window."
    )
    print(disclaimer_text)
    user_input = input("Do you agree to proceed? (y/all to continue): ").strip().lower()

    if user_input == 'all':
        all_test = True

    if user_input in ['yes', 'y', 'Y', 'all', '']:
        print("Consent received. Proceeding with the script...")
    else:
        print("Consent not granted. Exiting the script...")
        sys.exit()


def create_restore_point():
    """Creates a system restore point."""
    ps_script = '''
    $description = "QuickFix Restore Point"
    $restorePoint = @{
        Description = $description
        RestorePointType = "MODIFY_SETTINGS"
        EventType = "BEGIN_SYSTEM_CHANGE"
    }
    Checkpoint-Computer @restorePoint -ErrorAction Stop
    Write-Output "Restore point created successfully."
    '''
    try:
        result = subprocess.run(["powershell", "-Command", ps_script], capture_output=True, text=True, check=True)
        logging.info(f"Restore point created: {result.stdout.strip()}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to create restore point. Error: {e.stderr}")


def run_command(command):
    """Execute system command with suppressed output."""
    try:
        subprocess.run(command, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        logging.info(f"Command executed successfully: {' '.join(command)}")
    except subprocess.CalledProcessError as e:
        logging.error(f"Error running command: {e}")


def check_security():
    """Check the antivirus and firewall status using PowerShell."""
    ps_script = '''
    Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | 
    Select-Object displayName, productState | 
    ForEach-Object { "$($_.displayName):" + (" ON" * ($_.productState -band 0x10) + " OFF" * -not ($_.productState -band 0x10)) }

    Get-NetFirewallProfile -Profile Domain,Public,Private | 
    Select-Object Name,Enabled | 
    ForEach-Object { "$($_.Name):" + (" ON" * $_.Enabled + " OFF" * -not $_.Enabled) }
    '''
    try:
        result = subprocess.check_output(["powershell", "-Command", ps_script], text=True)
        format_security_output(result)
    except subprocess.CalledProcessError as e:
        logging.error(f"{Fore.RED}Failed to check security status: {e}")


def format_security_output(result):
    """Format the security output from PowerShell for better display."""
    logging.info(f"{Fore.CYAN}Security Check Results:")
    for line in result.strip().split('\n'):
        if line:
            logging.info(f"{Fore.GREEN}{line.strip()}")


def download_file(url, target):
    """Download a file from a URL to a target location."""
    try:
        response = requests.get(url)
        response.raise_for_status()
        with open(target, 'wb') as f:
            f.write(response.content)
        logging.info(f"{Fore.GREEN}Downloaded {os.path.basename(target)} successfully.")
    except requests.RequestException as e:
        logging.error(f"{Fore.RED}Failed to download {os.path.basename(target)}.")


def download_files_concurrently():
    """Download files using multiple threads."""
    logging.info(f"{Fore.BLUE}Start downloading additional files...")
    threads = []
    for file in BUGSPLAT_FILES:
        target_path = os.path.join(SYSTEM32_PATH, file)
        thread = threading.Thread(target=download_file, args=(f"{BUGSPLAT_URL}/{file}", target_path))
        threads.append(thread)
        thread.start()
        time.sleep(0.5)

    for thread in threads:
        thread.join()
    logging.info("All downloads completed.")


def admin_check():
    """Check if the script is running as an administrator, and restart with admin privileges if not."""
    if os.name == 'nt':
        try:
            is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        except:
            is_admin = False
        if not is_admin:
            print("This script needs to be run with administrative privileges.")
            logging.error(f"{Fore.RED}This script needs to be run with administrative privileges.")
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
            exit(0)


def prompt_for_restart():
    """Prompts the user for a system restart."""
    user_input = input(
        f"{Fore.CYAN}After the repair is complete, restarting the computer is necessary. Would you like to restart your system now? (y/n): ").strip().lower()
    if user_input in ['yes', 'y']:
        try:
            logging.info("Restarting the system...")
            subprocess.run(["shutdown", "/r", "/t", "0"], check=True)
        except subprocess.CalledProcessError as e:
            logging.error(f"Failed to restart the system. Error: {e}")
    else:
        logging.info("System restart skipped by the user.")


def check_and_install_updates():
    """Check for Windows updates and install them if available."""

    ps_script = '''
    $UpdateSession = New-Object -ComObject Microsoft.Update.Session
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $SearchResult = $UpdateSearcher.Search("IsInstalled=0")
    
    if ($SearchResult.Updates.Count -gt 0) {
        Write-Output "Updates found. Preparing to install..."
        $UpdatesToDownload = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($Update in $SearchResult.Updates) {
            if ($Update.EulaAccepted -eq $false) {
                $Update.AcceptEula()
            }
            $UpdatesToDownload.Add($Update) > $null
        }
        
        $Downloader = $UpdateSession.CreateUpdateDownloader()
        $Downloader.Updates = $UpdatesToDownload
        $DownloadResult = $Downloader.Download()
        
        $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($Update in $UpdatesToDownload) {
            if ($Update.IsDownloaded) {
                $UpdatesToInstall.Add($Update) > $null
            }
        }
        
        $Installer = $UpdateSession.CreateUpdateInstaller()
        $Installer.Updates = $UpdatesToInstall
        $InstallResult = $Installer.Install()
        if ($InstallResult.ResultCode -eq 2) {
            Write-Output "Updates installed successfully."
        } else {
            Write-Output "Failed to install some updates."
        }
    } else {
        Write-Output "No updates found. Your system is up to date."
    }
    '''
    try:
        logging.info(f"{Fore.BLUE}Check for Windows updates and install them if available.")

        result = subprocess.run(["powershell", "-Command", ps_script], capture_output=True, text=True, check=True)
        logging.info(f"{Fore.GREEN}{result.stdout.strip()}")
    except subprocess.CalledProcessError as e:
        logging.error(f"{Fore.RED}Failed to check or install updates. Error: {e.stderr.strip()}")


'''
def check_internet_speed():
    """Checks internet connection and measures download and upload speed."""
    logging.info("Checks internet connection and measures download and upload speed.")
    try:
        print("Checking internet connection and measuring speed...")
        st = speedtest.Speedtest()
        st.get_best_server()
        download_speed = st.download() / 1_000_000  # Convert to Mbps
        upload_speed = st.upload() / 1_000_000  # Convert to Mbps
        print(f"Download speed: {download_speed:.2f} Mbps")
        print(f"Upload speed: {upload_speed:.2f} Mbps")
    except speedtest.ConfigRetrievalError:
        print("Failed to retrieve speedtest configuration. Check your internet connection.")
    except Exception as e:
        print(f"An error occurred: {e}")
'''


def main():
    clear_screen()
    admin_check()
    delete_previous_logs()
    setup_logging()
    # print(ASCII_ART)
    print(f.renderText(text))
    user_consent()
    clear_screen()


    os.system('chcp 65001 >nul')
    create_restore_point()

    check_security()
    os.makedirs(APPDATA_PATH, exist_ok=True)
    run_command(['taskkill', '/F', '/IM', 'AnyDesk.exe'])

    download_files_concurrently()
    logging.info(f"{Fore.BLUE}Start downloading vc_redist.x64...")
    download_file(VC_REDIST_URL, os.path.join(APPDATA_PATH, 'vc_redist.x64.exe'))
    logging.info(f"{Fore.BLUE}Start downloading jreinstaller...")
    download_file(JAVA_URL, os.path.join(APPDATA_PATH, 'jreinstaller.exe'))
    logging.info(f"{Fore.BLUE}Start installing vc_redist.x64...")
    run_command([os.path.join(APPDATA_PATH, 'vc_redist.x64.exe'), '/install', '/quiet'])
    logging.info(f"{Fore.BLUE}Start installing jreinstaller...")
    run_command(
        [os.path.join(APPDATA_PATH, 'jreinstaller.exe'), 'INSTALL_SILENT=Enable', 'STATIC=Disable', 'WEB_JAVA=Disable',
         'WEB_ANALYTICS=Disable', 'REBOOT=Disable'])

    if all_test:
        #check_internet_speed()
        check_and_install_updates()
        logging.info(f"{Fore.BLUE}Start scanning system file...")
        run_command(['sfc', '/scannow'])
        logging.info(f"{Fore.BLUE}Start scanning and repairing system file...")
        run_command(['dism', '/online', '/cleanup-image', '/restorehealth'])

    print("\n")
    logging.info(
        f"{Fore.GREEN}All testing has been completed and the logs have been saved to the desktop, if there are any further questions, please provide your log files\n\n")

    prompt_for_restart()


if __name__ == "__main__":
    main()
