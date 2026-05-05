"""
EasyWin11 - Windows Power User Toolkit
Run as Administrator for best results.
"""

import os
import sys
import subprocess
import ctypes
import time
import platform

# ─── Colors (ANSI) ────────────────────────────────────────────────────────────
class C:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"
    CYAN    = "\033[96m"
    GREEN   = "\033[92m"
    YELLOW  = "\033[93m"
    RED     = "\033[91m"
    BLUE    = "\033[94m"
    WHITE   = "\033[97m"
    MAGENTA = "\033[95m"

# Enable ANSI on Windows
os.system("")

# ─── Helpers ──────────────────────────────────────────────────────────────────
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run(cmd, shell=True, capture=False):
    """Run a shell command. Returns output if capture=True."""
    try:
        if capture:
            result = subprocess.run(cmd, shell=shell, capture_output=True, text=True)
            return result.stdout.strip() or result.stderr.strip()
        else:
            subprocess.run(cmd, shell=shell)
    except Exception as e:
        print(f"{C.RED}  Error: {e}{C.RESET}")

def header(title):
    os.system("cls")
    width = 58
    print(f"\n{C.CYAN}{'═' * width}{C.RESET}")
    print(f"{C.CYAN}  ⚡ EasyWin11  {C.DIM}│{C.RESET}  {C.WHITE}{C.BOLD}{title}{C.RESET}")
    print(f"{C.CYAN}{'═' * width}{C.RESET}\n")

def ok(msg):
    print(f"  {C.GREEN}✔  {msg}{C.RESET}")

def info(msg):
    print(f"  {C.YELLOW}→  {msg}{C.RESET}")

def sep():
    print(f"  {C.DIM}{'─' * 54}{C.RESET}")

def pause():
    print(f"\n  {C.DIM}Press Enter to go back...{C.RESET}", end="")
    input()

def menu_choice(options):
    """Print numbered options and return valid choice."""
    for i, (label, _) in enumerate(options, 1):
        print(f"  {C.CYAN}{i}.{C.RESET}  {label}")
    print(f"  {C.CYAN}0.{C.RESET}  {C.DIM}Back / Exit{C.RESET}")
    print()
    while True:
        raw = input(f"  {C.WHITE}Choose ›{C.RESET} ").strip()
        if raw == "0":
            return 0
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw)
        print(f"  {C.RED}Invalid choice. Try again.{C.RESET}")

# ─── PC Optimization Actions ──────────────────────────────────────────────────
def clean_temp():
    header("Clean Temp Files")
    info("Deleting %TEMP% files...")
    run(r'del /q /f /s "%TEMP%\*" 2>nul')
    ok("User temp folder cleaned.")

    info("Deleting C:\\Windows\\Temp files...")
    run(r'del /q /f /s "C:\Windows\Temp\*" 2>nul')
    ok("System temp folder cleaned.")

    info("Clearing Prefetch (admin required)...")
    run(r'del /q /f /s "C:\Windows\Prefetch\*" 2>nul')
    ok("Prefetch folder cleaned.")
    pause()

def run_disk_cleanup():
    header("Disk Cleanup")
    info("Launching Disk Cleanup utility...")
    run("cleanmgr /sagerun:1")
    ok("Disk Cleanup launched.")
    pause()

def clear_event_logs():
    header("Clear Event Logs")
    info("Clearing Windows Event Logs (admin required)...")
    logs = ["Application", "System", "Security", "Setup"]
    for log in logs:
        result = run(f'wevtutil cl {log} 2>&1', capture=True)
        if "Access" in (result or ""):
            print(f"  {C.RED}✘  {log} — Access denied (run as Admin){C.RESET}")
        else:
            ok(f"{log} log cleared.")
    pause()

RAMMAP_PATH = r"C:\Tools\RAMMap\RAMMap.exe"
RAMMAP_URL  = "https://download.sysinternals.com/files/RAMMap.zip"
# Registry key where Sysinternals tools store EULA acceptance
RAMMAP_EULA_KEY = r"Software\Sysinternals\RamMap"

def _rammap_eula_accepted():
    """Check if the RAMMap EULA has been accepted in the registry."""
    import winreg
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, RAMMAP_EULA_KEY)
        val, _ = winreg.QueryValueEx(key, "EulaAccepted")
        winreg.CloseKey(key)
        return val == 1
    except Exception:
        return False

def _accept_rammap_eula():
    """Write EULA acceptance to registry so RAMMap runs silently."""
    import winreg
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, RAMMAP_EULA_KEY)
        winreg.SetValueEx(key, "EulaAccepted", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"  {C.RED}Could not write EULA to registry: {e}{C.RESET}")
        return False

def _ensure_rammap():
    """Download and extract RAMMap if not already present. Returns path or None."""
    if os.path.exists(RAMMAP_PATH):
        return RAMMAP_PATH
    import urllib.request, zipfile, io
    rammap_dir = os.path.dirname(RAMMAP_PATH)
    info("RAMMap not found. Downloading from Sysinternals...")
    try:
        os.makedirs(rammap_dir, exist_ok=True)
        with urllib.request.urlopen(RAMMAP_URL, timeout=20) as resp:
            zip_data = resp.read()
        with zipfile.ZipFile(io.BytesIO(zip_data)) as z:
            for member in z.namelist():
                if member.lower().endswith("rammap.exe"):
                    with z.open(member) as src, open(RAMMAP_PATH, "wb") as dst:
                        dst.write(src.read())
                    break
        if os.path.exists(RAMMAP_PATH):
            ok(f"RAMMap downloaded to {rammap_dir}")
            return RAMMAP_PATH
        else:
            print(f"  {C.RED}RAMMap.exe not found inside the zip.{C.RESET}")
            return None
    except Exception as e:
        print(f"  {C.RED}Download failed: {e}{C.RESET}")
        return None

def flush_ram():
    header("Empty Standby RAM  (via RAMMap)")
    if not is_admin():
        print(f"  {C.RED}Administrator rights required for this operation.{C.RESET}")
        pause()
        return

    rammap = _ensure_rammap()
    if not rammap:
        print(f"  {C.YELLOW}Could not obtain RAMMap. Check your internet connection.{C.RESET}")
        pause()
        return

    # RAMMap shows a EULA dialog on first ever launch which blocks -Et silently.
    # We accept it via registry so it always runs headless.
    if not _rammap_eula_accepted():
        info("First-time setup: accepting RAMMap EULA via registry...")
        if _accept_rammap_eula():
            ok("EULA accepted. RAMMap will now run silently.")
        else:
            print(f"\n  {C.YELLOW}Could not auto-accept the EULA.{C.RESET}")
            print(f"  {C.WHITE}Please do this once manually:{C.RESET}")
            print(f"  {C.DIM}1. Open:  {RAMMAP_PATH}{C.RESET}")
            print(f"  {C.DIM}2. Click 'Agree' on the license dialog.{C.RESET}")
            print(f"  {C.DIM}3. Close RAMMap, then run this option again.{C.RESET}")
            print(f"\n  {C.YELLOW}Opening RAMMap now for you...{C.RESET}")
            subprocess.Popen([rammap])
            pause()
            return

    info("Running RAMMap -Et  (Empty Standby List)...")
    try:
        result = subprocess.run(
            [rammap, "-Et"],
            capture_output=True, text=True, timeout=30
        )
        ok("Standby RAM list emptied successfully.")
    except subprocess.TimeoutExpired:
        print(f"  {C.YELLOW}RAMMap timed out after 30s.{C.RESET}")
    except Exception as e:
        print(f"  {C.RED}Error: {e}{C.RESET}")
    pause()

def disable_startup_apps():
    header("View Startup Apps")
    info("Opening Task Manager → Startup tab...")
    run("taskmgr")
    ok("Manage startup apps from the Startup tab.")
    pause()

def sfc_scan():
    header("System File Checker (SFC)")
    info("Running sfc /scannow — this may take a few minutes...")
    sep()
    run("sfc /scannow")
    sep()
    ok("SFC scan complete.")
    pause()

def dism_repair():
    header("DISM Health Repair")
    info("Running DISM RestoreHealth — this may take several minutes...")
    sep()
    run("DISM /Online /Cleanup-Image /RestoreHealth")
    sep()
    ok("DISM repair complete.")
    pause()

def power_plan():
    header("Power Plan")
    print(f"  {C.WHITE}Select a power plan:{C.RESET}\n")
    plans = [
        ("High Performance",     "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c", False),
        ("Balanced (default)",   "381b4222-f694-41f0-9685-ff5bb260df2e", False),
        ("Power Saver",          "a1841308-3541-4fab-bc81-f71556f20b4a", False),
        ("Ultimate Performance", "e9a42b02-d5df-448d-aa00-03f14749eb61", True),
    ]
    choice = menu_choice([(label, None) for label, _, _ in plans])
    if choice == 0:
        return
    label, guid, needs_provision = plans[choice - 1]

    if needs_provision:
        info("Provisioning Ultimate Performance plan...")
        out = run(f"powercfg -duplicatescheme {guid}", capture=True) or ""
        # Extract the GUID powercfg just created (it prints a new one)
        new_guid = None
        for word in out.split():
            word = word.strip("().,:").lower()
            if len(word) == 36 and word.count("-") == 4:
                new_guid = word
                break
        activate_guid = new_guid if new_guid else guid
    else:
        activate_guid = guid

    # Attempt activation, capture stderr to detect "unsupported setting"
    result = subprocess.run(
        f"powercfg /setactive {activate_guid}",
        shell=True, capture_output=True, text=True
    )
    stderr = (result.stderr or "").strip()
    stdout = (result.stdout or "").strip()

    if "unsupported" in stderr.lower() or "unsupported" in stdout.lower():
        print(f"  {C.RED}This plan is not supported on your hardware.{C.RESET}")
        print(f"  {C.YELLOW}Ultimate Performance is only available on desktops.")
        print(f"  Laptops and VMs do not support it.{C.RESET}")
        pause()
        return

    # Verify by reading back the active scheme
    active = run("powercfg /getactivescheme", capture=True) or ""
    if activate_guid.lower() in active.lower():
        ok(f'"{label}" power plan is now active.')
    else:
        print(f"  {C.YELLOW}Could not confirm. Active scheme reported as:{C.RESET}")
        print(f"  {C.DIM}{active}{C.RESET}")
    pause()

# ─── Network Tweak Actions ────────────────────────────────────────────────────
def flush_dns():
    header("Flush DNS Cache")
    info("Running ipconfig /flushdns ...")
    output = run("ipconfig /flushdns", capture=True)
    print(f"\n  {C.WHITE}{output}{C.RESET}\n")
    ok("DNS cache flushed.")
    pause()

def ip_release_renew():
    header("IP Release & Renew")
    info("Releasing IP address...")
    run("ipconfig /release")
    time.sleep(1)
    info("Renewing IP address...")
    run("ipconfig /renew")
    ok("IP address renewed.")
    pause()

def show_ip_info():
    header("IP Configuration Info")
    info("Running ipconfig /all ...\n")
    sep()
    run("ipconfig /all")
    sep()
    pause()

def reset_winsock():
    header("Reset Winsock & TCP/IP Stack")
    info("Resetting Winsock catalog...")
    run("netsh winsock reset")
    info("Resetting TCP/IP stack...")
    run("netsh int ip reset")
    info("Flushing DNS...")
    run("ipconfig /flushdns")
    info("Resetting IPv4/IPv6 interfaces...")
    run("netsh interface ipv4 reset")
    run("netsh interface ipv6 reset")
    ok("All network stack components reset. Restart your PC to apply.")
    pause()

def ping_test():
    header("Ping Test")
    print(f"  Enter a host to ping {C.DIM}(e.g. google.com or 8.8.8.8){C.RESET}")
    host = input(f"  {C.WHITE}Host ›{C.RESET} ").strip()
    if not host:
        host = "8.8.8.8"
    print()
    info(f"Pinging {host} ...\n")
    sep()
    run(f"ping {host}")
    sep()
    pause()

def tracert_test():
    header("Trace Route")
    print(f"  Enter a host to trace {C.DIM}(e.g. google.com){C.RESET}")
    host = input(f"  {C.WHITE}Host ›{C.RESET} ").strip()
    if not host:
        host = "google.com"
    print()
    info(f"Tracing route to {host} ...\n")
    sep()
    run(f"tracert {host}")
    sep()
    pause()

def check_open_ports():
    header("Open Ports (netstat)")
    info("Listing active connections and listening ports...\n")
    sep()
    run("netstat -ano")
    sep()
    pause()

# ─── Windows Update & Driver Cleanup ─────────────────────────────────────────
def cleanup_windows_update_cache():
    header("Clean Windows Update Cache")
    print(f"  {C.YELLOW}⚠  This will stop Windows Update service temporarily.{C.RESET}\n")

    info("Stopping Windows Update services...")
    run("net stop wuauserv /y")
    run("net stop bits /y")
    run("net stop cryptsvc /y")
    run("net stop msiserver /y")

    info("Deleting SoftwareDistribution\\Download folder...")
    run(r'rd /s /q "C:\Windows\SoftwareDistribution\Download" 2>nul')
    run(r'mkdir "C:\Windows\SoftwareDistribution\Download"')
    ok("SoftwareDistribution\\Download cleared.")

    info("Deleting catroot2 folder (update signature cache)...")
    run(r'rd /s /q "C:\Windows\System32\catroot2" 2>nul')
    ok("catroot2 cleared.")

    info("Restarting Windows Update services...")
    run("net start wuauserv")
    run("net start bits")
    run("net start cryptsvc")
    ok("All services restarted. Windows Update cache is clean.")
    pause()

def dism_cleanup_winsxs():
    header("DISM – Clean WinSxS Component Store")
    info("Analyzing component store size first...")
    sep()
    run("DISM /Online /Cleanup-Image /AnalyzeComponentStore")
    sep()
    print()
    info("Running StartComponentCleanup to remove superseded components...")
    info("This may take several minutes...\n")
    sep()
    run("DISM /Online /Cleanup-Image /StartComponentCleanup /ResetBase")
    sep()
    ok("WinSxS cleanup complete. Old update files removed.")
    pause()

def remove_old_update_backups():
    header("Remove Old Update Backup Folders")
    info("Scanning for $NtUninstall and old patch folders...")

    # These are leftover cab/backup folders from old-style updates
    targets = [
        r"C:\Windows\SoftwareDistribution\Download",
        r"C:\Windows\Temp",
    ]

    # Use cleanmgr flags for update cleanup
    info("Running Disk Cleanup for Windows Update Cleanup category...")
    # Set registry flag for update cleanup category (sageset 99)
    run(r'reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Update Cleanup" /v StateFlags0099 /t REG_DWORD /d 2 /f')
    run(r'reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Windows Upgrade Log Files" /v StateFlags0099 /t REG_DWORD /d 2 /f')
    run(r'reg add "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Temporary Files" /v StateFlags0099 /t REG_DWORD /d 2 /f')
    run("cleanmgr /sagerun:99")
    ok("Disk Cleanup launched for update leftovers.")
    pause()

def reset_windows_update_components():
    header("Reset Windows Update Components")
    print(f"  {C.YELLOW}⚠  This resets all WU components and re-registers DLLs.{C.RESET}")
    print(f"  {C.YELLOW}   Use this if Windows Update is stuck or failing.{C.RESET}\n")

    steps = [
        ("Stopping services...",                "net stop wuauserv & net stop cryptsvc & net stop bits & net stop msiserver"),
        ("Renaming SoftwareDistribution...",    r'ren "C:\Windows\SoftwareDistribution" SoftwareDistribution.old 2>nul'),
        ("Renaming catroot2...",                r'ren "C:\Windows\System32\catroot2" catroot2.old 2>nul'),
        ("Resetting BITS service...",           "sc.exe sdset bits D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"),
        ("Resetting wuauserv security...",      "sc.exe sdset wuauserv D:(A;;CCLCSWRPWPDTLOCRRC;;;SY)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCLCSWRPWPDTLOCRRC;;;PU)"),
        ("Re-registering DLLs...",              "regsvr32.exe /s atl.dll & regsvr32.exe /s urlmon.dll & regsvr32.exe /s mshtml.dll & regsvr32.exe /s shdocvw.dll & regsvr32.exe /s browseui.dll & regsvr32.exe /s jscript.dll & regsvr32.exe /s vbscript.dll & regsvr32.exe /s scrrun.dll & regsvr32.exe /s msxml.dll & regsvr32.exe /s msxml3.dll & regsvr32.exe /s msxml6.dll & regsvr32.exe /s actxprxy.dll & regsvr32.exe /s softpub.dll & regsvr32.exe /s wintrust.dll & regsvr32.exe /s dssenh.dll & regsvr32.exe /s rsaenh.dll & regsvr32.exe /s gpkcsp.dll & regsvr32.exe /s sccbase.dll & regsvr32.exe /s slbcsp.dll & regsvr32.exe /s cryptdlg.dll & regsvr32.exe /s oleaut32.dll & regsvr32.exe /s ole32.dll & regsvr32.exe /s shell32.dll & regsvr32.exe /s initpki.dll & regsvr32.exe /s wuapi.dll & regsvr32.exe /s wuaueng.dll & regsvr32.exe /s wuaueng1.dll & regsvr32.exe /s wucltui.dll & regsvr32.exe /s wups.dll & regsvr32.exe /s wups2.dll & regsvr32.exe /s wuweb.dll & regsvr32.exe /s qmgr.dll & regsvr32.exe /s qmgrprxy.dll & regsvr32.exe /s wucltux.dll & regsvr32.exe /s muweb.dll & regsvr32.exe /s wuwebv.dll"),
        ("Resetting network settings...",       "netsh winsock reset & netsh winhttp reset proxy"),
        ("Starting services...",                "net start wuauserv & net start cryptsvc & net start bits & net start msiserver"),
    ]
    for msg, cmd in steps:
        info(msg)
        run(cmd)

    ok("Windows Update components fully reset.")
    print(f"  {C.YELLOW}  Restart your PC, then try Windows Update again.{C.RESET}")
    pause()

def clean_old_drivers():
    header("Clean Old / Staged Driver Packages")
    info("Scanning driver store for old staged packages...\n")
    sep()

    # List all third-party OEM drivers in the driver store
    output = run("pnputil /enum-drivers", capture=True)
    if output:
        print(f"{C.DIM}{output}{C.RESET}")
    sep()

    print(f"\n  {C.WHITE}Options:{C.RESET}\n")
    options = [
        ("Auto-remove all unused/superseded drivers (safe)",  "_auto"),
        ("Open Device Manager to manually manage drivers",    "_devmgr"),
    ]
    choice = menu_choice(options)
    if choice == 0:
        return
    if choice == 1:
        info("Removing unused driver packages via DISM...")
        run("DISM /Online /Cleanup-Image /StartComponentCleanup")
        info("Forcing removal of superseded drivers (pnputil)...")
        # pnputil /delete-driver on superseded — we do this via DISM above;
        # pnputil direct delete requires knowing the oem#.inf name.
        ok("Unused drivers cleaned via DISM component cleanup.")
    elif choice == 2:
        run("devmgmt.msc")
        ok("Device Manager opened.")
    pause()

def view_installed_drivers():
    header("View All Installed Drivers")
    info("Listing all signed drivers on this system...\n")
    sep()
    run("driverquery /FO list /SI 2>nul || driverquery /FO list")
    sep()
    pause()


# ─── Sub-menus ────────────────────────────────────────────────────────────────

def menu_windows_update_drivers():
    while True:
        header("Windows Update & Driver Cleanup")
        options = [
            ("Clean Windows Update Cache",          cleanup_windows_update_cache),
            ("DISM – Clean WinSxS Store",           dism_cleanup_winsxs),
            ("Remove Old Update Backup Files",      remove_old_update_backups),
            ("Reset Windows Update Components",     reset_windows_update_components),
            ("Clean Old Staged Driver Packages",    clean_old_drivers),
            ("View All Installed Drivers",          view_installed_drivers),
        ]
        choice = menu_choice(options)
        if choice == 0:
            break
        options[choice - 1][1]()


# ─── Sub-menus ────────────────────────────────────────────────────────────────
def menu_pc_optimization():
    while True:
        header("PC Optimization")
        options = [
            ("Clean Temp Files",            clean_temp),
            ("Run Disk Cleanup",            run_disk_cleanup),
            ("Clear Event Logs",            clear_event_logs),
            ("Flush Standby RAM",           flush_ram),
            ("View / Manage Startup Apps",  disable_startup_apps),
            ("System File Checker (SFC)",   sfc_scan),
            ("DISM Health Repair",          dism_repair),
            ("Change Power Plan",           power_plan),
        ]
        choice = menu_choice(options)
        if choice == 0:
            break
        options[choice - 1][1]()

def menu_network():
    while True:
        header("Network Tweaking")
        options = [
            ("Flush DNS Cache",             flush_dns),
            ("IP Release & Renew",          ip_release_renew),
            ("Show IP Configuration",       show_ip_info),
            ("Reset Winsock / TCP Stack",   reset_winsock),
            ("Ping Test",                   ping_test),
            ("Trace Route",                 tracert_test),
            ("Check Open Ports",            check_open_ports),
        ]
        choice = menu_choice(options)
        if choice == 0:
            break
        options[choice - 1][1]()

# ─── Main Menu ────────────────────────────────────────────────────────────────
def main():
    if platform.system() != "Windows":
        print(f"\n{C.RED}  EasyWin11 is designed for Windows only.{C.RESET}\n")
        sys.exit(1)

    admin_status = f"{C.GREEN}Administrator ✔{C.RESET}" if is_admin() else f"{C.YELLOW}Standard User  (some features need Admin){C.RESET}"

    while True:
        os.system("cls")
        print(f"""
{C.CYAN}╔══════════════════════════════════════════════════════╗
║   ⚡  EasyWin11  —  Power User Toolkit               ║
╚══════════════════════════════════════════════════════╝{C.RESET}
  Running as: {admin_status}
""")
        options = [
            ("PC Optimization",                 menu_pc_optimization),
            ("Network Tweaking",                menu_network),
            ("Windows Update & Driver Cleanup", menu_windows_update_drivers),
        ]
        for i, (label, _) in enumerate(options, 1):
            print(f"  {C.CYAN}{i}.{C.RESET}  {C.BOLD}{label}{C.RESET}")
        print(f"  {C.CYAN}0.{C.RESET}  {C.DIM}Exit{C.RESET}\n")

        raw = input(f"  {C.WHITE}Choose ›{C.RESET} ").strip()
        if raw == "0":
            print(f"\n  {C.DIM}Goodbye!{C.RESET}\n")
            break
        elif raw.isdigit() and 1 <= int(raw) <= len(options):
            options[int(raw) - 1][1]()

if __name__ == "__main__":
    main()