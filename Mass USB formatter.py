import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.scrolledtext as st
import psutil
import win32api
import win32file
import win32con
import os
import shutil
import subprocess
import hashlib
import time
from threading import Thread, Event
import logging
from queue import Queue
from datetime import datetime

# Global variables and stop event
selected_drives = []
log_file_path = None
task_queue = Queue()
stop_event = Event()  # When set, all operations should stop.
original_size = {}    # For storing drive sizes before format

def setup_logging():
    global log_file_path
    # Clear existing handlers so basicConfig works as expected
    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)
    if log_file_path:
        logging.basicConfig(
            filename=log_file_path,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
    else:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

# Custom logging handler to write messages to a Tkinter ScrolledText widget.
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.text_widget.tag_config("error", foreground="red")

    def emit(self, record):
        msg = self.format(record) + "\n"
        if record.levelno >= logging.ERROR:
            self.text_widget.after(0, self.text_widget.insert, tk.END, msg, "error")
        else:
            self.text_widget.after(0, self.text_widget.insert, tk.END, msg)
        self.text_widget.after(0, self.text_widget.see, tk.END)

def setup_text_logging(text_widget):
    
    text_handler = TextHandler(text_widget)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    text_handler.setFormatter(formatter)
    logging.getLogger().addHandler(text_handler)

def list_usb_drives():
    drives = []
    for p in psutil.disk_partitions(all=False):
        if 'removable' in p.opts.lower() and os.path.exists(p.mountpoint):
            try:
                label = win32api.GetVolumeInformation(p.device)[0]
            except Exception as e:
                logging.error(f"Error getting volume label for {p.device}: {e}")
                label = "Unknown"
            usage = psutil.disk_usage(p.mountpoint)
            size_gb = round(usage.total / (1024 ** 3), 2)
            drives.append({
                'device': p.device.rstrip("\\"),  # e.g., "F:"
                'mount': p.mountpoint,
                'label': label,
                'size': size_gb
            })
    return drives

def update_drive_count():
    count = sum(1 for d in selected_drives if d['var'].get())
    drive_count_label.config(text=f"Selected Drives: {count}")

def refresh_drive_list():
    for widget in drive_check_frame.winfo_children():
        widget.destroy()
    selected_drives.clear()
    for idx, drive in enumerate(list_usb_drives()):
        var = tk.BooleanVar()
        var.trace("w", lambda *args: update_drive_count())
        cb = ttk.Checkbutton(
            drive_check_frame,
            text=f"{drive['device']} - {drive['label']} - {drive['size']}GB",
            variable=var
        )
        cb.grid(row=idx, column=0, sticky=tk.W)
        selected_drives.append({'drive': drive['device'], 'mount': drive['mount'], 'var': var})
    update_drive_count()

def select_all_drives():
    for drive in selected_drives:
        drive['var'].set(True)
    update_drive_count()

def wait_for_drive(drive_letter, timeout=60, retries=3):
    drive_letter = drive_letter.strip(":\\")
    for attempt in range(retries):
        start_time = time.time()
        path = f"{drive_letter}:\\"
        while time.time() - start_time < timeout:
            if stop_event.is_set():
                logging.info("Stop event detected in wait_for_drive. Exiting wait loop.")
                return False
            if os.path.exists(path):
                logging.info(f"Drive {drive_letter} is now available after {time.time() - start_time:.2f} seconds (attempt {attempt + 1})")
                return True
            time.sleep(0.5)
        logging.warning(f"Drive {drive_letter} not available after {timeout} seconds (attempt {attempt + 1})")
        if attempt < retries - 1:
            logging.info(f"Retrying to find drive {drive_letter} ({retries - attempt - 1} attempts left)")
            time.sleep(5)
    return False

def find_drive_after_format(original_letter, original_label):
    key = original_letter.strip(":\\")
    drives = list_usb_drives()
    for drive in drives:
        drive_letter = drive['device'].strip(":\\")
        if drive['label'] == original_label and drive['size'] == original_size.get(key, 0):
            logging.info(f"Drive originally at {key} found at {drive_letter} after formatting")
            return drive_letter  # Return without colon.
    logging.warning(f"Could not find drive originally at {key} after formatting")
    return key

def format_drive(drive_letter, fs_type, label, dry_run=False):
    try:
        drive_letter_stripped = drive_letter.strip(":\\")
        if len(drive_letter_stripped) != 1 or not drive_letter_stripped.isalpha():
            raise ValueError(f"Invalid drive letter: {drive_letter}")

        drive_path = f"{drive_letter_stripped}:\\"
        if dry_run:
            logging.info(f"[DRY RUN] Would format {drive_letter_stripped}: with FS={fs_type}, label={label or '(no label)'}")
            return True

        if not os.path.exists(drive_path):
            raise ValueError(f"Drive {drive_letter_stripped}: is not accessible or doesn't exist")

        label = label if label else "USB_DRIVE"
        label = label[:11]
        logging.info(f"Formatting {drive_letter_stripped}: with FS={fs_type} and label '{label}'")

        usage = psutil.disk_usage(drive_path)
        original_size[drive_letter_stripped] = round(usage.total / (1024 ** 3), 2)

        # Build the command. If fs_type is FAT and Cisco support is enabled, add the cluster size flag.
        cmd = ["format", f"{drive_letter_stripped}:", "/FS:" + fs_type, "/V:" + label]
        if fs_type.upper() == "FAT" and cisco_support_var.get():
            cmd.append("/A:32768")
        cmd.extend(["/Q", "/X", "/Y"])

        proc = subprocess.run(
            cmd,
            input="y\n",
            capture_output=True,
            text=True,
            shell=True,
            timeout=300
        )

        if proc.returncode != 0:
            logging.error(f"Format failed for {drive_letter_stripped}:\n{proc.stderr}")
            return False

        logging.info(proc.stdout)
        return True

    except subprocess.TimeoutExpired:
        logging.error(f"Formatting {drive_letter_stripped}: timed out")
        return False
    except Exception as e:
        logging.error(f"Formatting {drive_letter_stripped}: {str(e)}")
        return False

def is_system_drive(drive_letter):
    return os.environ['SystemDrive'].strip(":\\").upper() == drive_letter.strip(":\\").upper()

def get_all_files_in_folder(src_folder):
    file_list = []
    for root_dir, _, files in os.walk(src_folder):
        for file in files:
            file_list.append(os.path.join(root_dir, file))
    return file_list

def compute_checksum(filepath, algorithm="sha256"):
    try:
        hash_func = getattr(hashlib, algorithm.lower(), None)
        if hash_func is None:
            raise ValueError(f"Unsupported algorithm: {algorithm}")
        hash_obj = hash_func()
        with open(filepath, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                hash_obj.update(chunk)
        logging.info(f"Checksum calculated for {filepath} using {algorithm.upper()}")
        return hash_obj.hexdigest()
    except Exception as e:
        logging.error(f"Hashing {filepath}: {e}")
        return None

def copy_folder_to_drive(src_folder, dst_drive, verify_checksums, drive_letter, dry_run=False):
    files = get_all_files_in_folder(src_folder)
    total_files = len(files)
    if total_files == 0:
        logging.info(f"No files found in {src_folder} to copy to drive {drive_letter}")
        return True

    selected_checksum = checksum_type_var.get()

    source_checksums = {}
    if verify_checksums:
        logging.info("Computing checksums for source folder files.")
        for file_path in files:
            rel_path = os.path.relpath(file_path, src_folder)
            chksum = compute_checksum(file_path, selected_checksum)
            if chksum is None:
                logging.error(f"Failed to compute checksum for {file_path}.")
                messagebox.showerror("Checksum Error", f"Failed to compute checksum for {file_path}.")
                return False
            source_checksums[rel_path] = chksum

    logging.info(f"Starting copy of {total_files} files from {src_folder} to {dst_drive}")

    dest_checksums = {}
    for file_path in files:
        if stop_event.is_set():
            logging.info("Stop event detected in file copy loop. Aborting copy.")
            return False

        rel_path = os.path.relpath(file_path, src_folder)
        target_path = os.path.join(dst_drive, rel_path)
        logging.info(f"Copying file: {file_path} to {target_path}")

        if dry_run:
            logging.info(f"[DRY RUN] Would copy: {file_path} → {target_path}")
            continue

        try:
            os.makedirs(os.path.dirname(target_path), exist_ok=True)
            shutil.copy2(file_path, target_path)
            logging.info(f"Copied {file_path} to {target_path}")
        except PermissionError as pe:
            logging.error(f"Permission denied while copying {file_path} to {target_path}: {pe}")
            continue
        except Exception as e:
            logging.error(f"Error copying {file_path} to {target_path}: {e}")
            continue

        if verify_checksums:
            dest_chksum = compute_checksum(target_path, selected_checksum)
            dest_checksums[rel_path] = dest_chksum
            if source_checksums.get(rel_path) != dest_chksum:
                error_msg = (f"Checksum mismatch for {rel_path}.\n"
                             f"Source: {source_checksums.get(rel_path)}\n"
                             f"Destination: {dest_chksum}")
                logging.error(error_msg)
                messagebox.showerror("Checksum Error", error_msg)
        time.sleep(0.01)  # Allow log updates

    # Write checksum file regardless of copy success, with a warning if not all files were copied
    if verify_checksums:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        checksum_filename = f"checksum_{selected_checksum}_{timestamp}.txt"
        
        # Write checksum file to destination drive
        dst_checksum_file = os.path.join(dst_drive, checksum_filename)
        try:
            with open(dst_checksum_file, "w") as f:
                if set(source_checksums.keys()) != set(dest_checksums.keys()):
                    f.write("WARNING: Not all files copied successfully or verified correctly.\n")
                for rel_path, chksum in source_checksums.items():
                    f.write(f"{rel_path}: {chksum}\n")
            logging.info(f"Checksum file written to {dst_checksum_file}")
        except Exception as e:
            logging.error(f"Failed to write checksum file to drive: {e}")
            messagebox.showerror("File Write Error", f"Failed to write checksum file to drive: {e}")

        # Write checksum file to source folder
        src_checksum_file = os.path.join(src_folder, checksum_filename)
        try:
            with open(src_checksum_file, "w") as f:
                if set(source_checksums.keys()) != set(dest_checksums.keys()):
                    f.write("WARNING: Not all files copied successfully or verified correctly.\n")
                for rel_path, chksum in source_checksums.items():
                    f.write(f"{rel_path}: {chksum}\n")
            logging.info(f"Checksum file written to source folder: {src_checksum_file}")
        except Exception as e:
            logging.error(f"Failed to write checksum file to source folder: {e}")
            messagebox.showerror("File Write Error", f"Failed to write checksum file to source folder: {e}")

    return True

def process_drive(drive, fs_type, label, folder, verify_checksums, dry_run_enabled, copy_files):
    drive_letter = drive['drive']
    if stop_event.is_set():
        logging.info("Stop event detected before processing drive. Aborting.")
        task_queue.get()
        task_queue.task_done()
        return

    if is_system_drive(drive_letter):
        root.after(0, lambda: messagebox.showerror("System Drive", f"{drive_letter} is a system drive. Skipping."))
        task_queue.get()
        task_queue.task_done()
        return

    # If "Format drives" is selected, perform formatting first.
    if format_var.get():
        formatted = format_drive(drive_letter, fs_type, label, dry_run_enabled)
        if not formatted or stop_event.is_set():
            root.after(0, lambda: messagebox.showerror("Format Error", f"Failed to format {drive_letter} or operation stopped."))
            task_queue.get()
            task_queue.task_done()
            return
        root.after(0, refresh_drive_list)
        time.sleep(1)
        new_drive_letter = find_drive_after_format(drive_letter, label)
        drive_letter = new_drive_letter  # now drive_letter is like "F" without colon.
        mount_path = f"{drive_letter}:\\"
        if copy_files and not dry_run_enabled:
            if not wait_for_drive(drive_letter):
                logging.info(f"Skipping copy to {drive_letter} as drive is not available after format")
                task_queue.get()
                task_queue.task_done()
                return
    else:
        drive_letter = drive_letter.strip(":\\")
        mount_path = f"{drive_letter}:\\"

    if copy_files:
        logging.info(f"Starting file copy to {drive_letter}")
        copied = copy_folder_to_drive(folder, mount_path, verify_checksums, drive_letter, dry_run_enabled)
        if not copied:
            root.after(0, lambda: messagebox.showerror("Copy Error", f"Failed to copy files to {mount_path}."))
            task_queue.get()
            task_queue.task_done()
            return
        else:
            logging.info(f"Successfully copied files to {drive_letter}")

    task_queue.get()
    task_queue.task_done()

def eject_drive(drive_letter):
    try:
        drive = drive_letter.strip(":\\")
        handle = win32file.CreateFile(r"\\.\%s:" % drive,
                                      win32con.GENERIC_READ,
                                      win32con.FILE_SHARE_READ | win32con.FILE_SHARE_WRITE,
                                      None,
                                      win32con.OPEN_EXISTING,
                                      0,
                                      None)
        # Lock, dismount, and eject the drive.
        win32file.DeviceIoControl(handle, 0x00090018, None, None)  # FSCTL_LOCK_VOLUME
        win32file.DeviceIoControl(handle, 0x00090020, None, None)  # FSCTL_DISMOUNT_VOLUME
        win32file.DeviceIoControl(handle, 0x002D4808, None, None)  # IOCTL_STORAGE_EJECT_MEDIA
        handle.Close()
        logging.info(f"Ejected drive {drive_letter}")
    except Exception as e:
        logging.error(f"Failed to eject drive {drive_letter}: {e}")

def mass_eject_drives():
    status_label.config(text="Status: Ejecting Drives...")
    drives_to_eject = [d for d in selected_drives if d['var'].get()]
    if not drives_to_eject:
        messagebox.showwarning("No Drives", "Select at least one USB drive to eject.")
        status_label.config(text="Status: Idle")
        return
    for drive in drives_to_eject:
        eject_drive(drive['drive'])
    status_label.config(text="Status: Idle")

def format_and_copy():
    stop_event.clear()
    status_label.config(text="Status: Running...")
    if not dry_run_var.get():
        if not messagebox.askyesno("Confirm", "This will erase all selected drives if formatting is enabled. Continue?"):
            status_label.config(text="Status: Idle")
            return

    fs_type = fs_var.get()
    label = label_var.get().strip()[:11]
    folder = folder_var.get()
    verify_checksums = verify_var.get()
    dry_run_enabled = dry_run_var.get()

    if not fs_type:
        messagebox.showwarning("No Filesystem", "Select a filesystem format.")
        status_label.config(text="Status: Idle")
        return

    targets = [d for d in selected_drives if d['var'].get()]
    if not targets:
        messagebox.showwarning("No Drives", "Select at least one USB drive.")
        status_label.config(text="Status: Idle")
        return

    copy_files = folder and os.path.isdir(folder)
    if folder and not os.path.isdir(folder):
        messagebox.showerror("Folder Error", f"The folder '{folder}' does not exist or is not accessible.")
        status_label.config(text="Status: Idle")
        return

    logging.info("Starting operations on selected drives.")
    for drive in targets:
        task_queue.put(drive)
        thread = Thread(target=process_drive, args=(drive, fs_var.get(), label, folder, verify_checksums, dry_run_enabled, copy_files))
        thread.start()

    root.after(100, check_queue)

def check_queue():
    if not task_queue.empty():
        root.after(100, check_queue)
    else:
        msg = "Operations stopped by user." if stop_event.is_set() else ("Dry-run complete!" if dry_run_var.get() else "All selected drives processed successfully.")
        root.after(0, lambda: messagebox.showinfo("Done", msg))
        format_button.config(state="normal")
        status_label.config(text="Status: Idle")

def start_format_and_copy():
    format_button.config(state="disabled")
    stop_event.clear()
    format_and_copy()

def stop_all_operations():
    stop_event.set()
    logging.info("Stop button pressed. Cancelling all operations.")
    messagebox.showinfo("Stop", "Stop requested. Ongoing operations will be cancelled shortly.")
    status_label.config(text="Status: Idle")

def choose_folder():
    path = filedialog.askdirectory()
    if path:
        folder_var.set(path)

def choose_log_file():
    global log_file_path
    path = filedialog.asksaveasfilename(
        defaultextension=".log",
        filetypes=[("Log files", "*.log"), ("All files", "*.*")],
        title="Choose Log File Location"
    )
    if path:
        log_file_path = path
        log_file_var.set(path)
        setup_logging()
        logging.info("Log file location set to: " + path)

def update_cisco_option(event=None, *args):
    if fs_var.get().upper() == "FAT" and format_var.get():
        cisco_check.grid()  # Show the checkbutton
    else:
        cisco_check.grid_remove()  # Hide it

# --- GUI Setup ---
root = tk.Tk()
root.title("Mass USB Formatter & Copier")
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky="nsew")
# Configure rows and columns for dynamic resizing
for i in range(21):
    main_frame.rowconfigure(i, weight=0)
main_frame.rowconfigure(19, weight=1)  # Give the log area extra weight
main_frame.columnconfigure(0, weight=1)
# Allow an extra column for the mass eject button
main_frame.columnconfigure(2, weight=0)

ttk.Label(main_frame, text="USB Drives:").grid(row=0, column=0, sticky=tk.W)
drive_check_frame = ttk.Frame(main_frame)
drive_check_frame.grid(row=1, column=0, columnspan=2, sticky=tk.W)
# Make sure drive_count_label is defined before calling refresh_drive_list()
drive_count_label = ttk.Label(main_frame, text="Selected Drives: 0")
drive_count_label.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
refresh_drive_list()

ttk.Button(main_frame, text="Refresh", command=refresh_drive_list).grid(row=2, column=0, sticky=tk.W, pady=5)
ttk.Button(main_frame, text="Select All", command=select_all_drives).grid(row=2, column=1, sticky=tk.W, pady=5)

ttk.Label(main_frame, text="Filesystem:").grid(row=4, column=0, sticky=tk.W)
fs_var = tk.StringVar()
fs_combo = ttk.Combobox(main_frame, textvariable=fs_var, values=["FAT", "FAT32", "NTFS", "exFAT"], state="readonly")
fs_combo.grid(row=5, column=0, sticky=tk.W)
fs_combo.current(1)
fs_combo.bind("<<ComboboxSelected>>", update_cisco_option)

format_var = tk.BooleanVar()
format_check = ttk.Checkbutton(main_frame, text="Format drives before copying", variable=format_var)
format_check.grid(row=6, column=0, columnspan=2, sticky=tk.W)
format_var.trace("w", update_cisco_option)

cisco_support_var = tk.BooleanVar()
cisco_check = ttk.Checkbutton(main_frame, text="Enable Cisco Support (set cluster size to 32K. Experimental not recommended)", variable=cisco_support_var)
cisco_check.grid(row=7, column=0, columnspan=2, sticky=tk.W)
if fs_var.get().upper() != "FAT" or not format_var.get():
    cisco_check.grid_remove()

ttk.Label(main_frame, text="Volume Label (optional, max 11 chars):").grid(row=8, column=0, sticky=tk.W, pady=(10, 0))
label_var = tk.StringVar()
label_entry = ttk.Entry(main_frame, textvariable=label_var, width=20)
label_entry.grid(row=9, column=0, sticky=tk.W)

ttk.Label(main_frame, text="Folder to Copy (optional):").grid(row=10, column=0, sticky=tk.W, pady=(10, 0))
folder_var = tk.StringVar()
folder_entry = ttk.Entry(main_frame, textvariable=folder_var, width=40)
folder_entry.grid(row=11, column=0, sticky=tk.W)
ttk.Button(main_frame, text="Browse", command=choose_folder).grid(row=11, column=1, padx=5)

ttk.Label(main_frame, text="Log File Location:").grid(row=12, column=0, sticky=tk.W, pady=(10, 0))
log_file_var = tk.StringVar()
log_file_entry = ttk.Entry(main_frame, textvariable=log_file_var, width=40)
log_file_entry.grid(row=13, column=0, sticky=tk.W)
ttk.Button(main_frame, text="Browse", command=choose_log_file).grid(row=13, column=1, padx=5)

verify_var = tk.BooleanVar()
verify_check = ttk.Checkbutton(main_frame, text="Verify checksums after copy", variable=verify_var)
verify_check.grid(row=14, column=0, columnspan=2, sticky=tk.W)

ttk.Label(main_frame, text="Checksum Algorithm:").grid(row=15, column=0, sticky=tk.W, pady=(10, 0))
checksum_type_var = tk.StringVar(value="sha256")
checksum_combo = ttk.Combobox(main_frame, textvariable=checksum_type_var, values=["sha256", "md5", "sha1"], state="readonly")
checksum_combo.grid(row=16, column=0, sticky=tk.W)

dry_run_var = tk.BooleanVar()
dry_run_check = ttk.Checkbutton(main_frame, text="Dry-run mode (no formatting or copying)", variable=dry_run_var)
dry_run_check.grid(row=17, column=0, columnspan=2, sticky=tk.W, pady=5)

format_button = ttk.Button(main_frame, text="Execute", command=start_format_and_copy)
format_button.grid(row=18, column=0, pady=10)
stop_button = ttk.Button(main_frame, text="Stop", command=stop_all_operations)
stop_button.grid(row=18, column=1, pady=10)
mass_eject_button = ttk.Button(main_frame, text="Mass Eject", command=mass_eject_drives)
mass_eject_button.grid(row=18, column=2, padx=5, pady=10)

# ScrolledText for logging
log_text = st.ScrolledText(main_frame, width=80, height=15)
log_text.grid(row=19, column=0, columnspan=3, sticky="nsew", pady=10)

# Status label to show running state
status_label = ttk.Label(main_frame, text="Status: Idle")
status_label.grid(row=20, column=0, columnspan=3, sticky=tk.W, pady=(0, 10))

# IMPORTANT: Configure logging before adding the text handler
setup_logging()
setup_text_logging(log_text)

root.mainloop()
