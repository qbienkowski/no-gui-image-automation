# script3_test_applications.py

import os
import sys
import time
import psutil
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from pywinauto import Desktop
from openpyxl import Workbook
from openpyxl.styles import Alignment
import logging
import threading
import re
from typing import List, Optional
import yaml

class ApplicationTester:
    """Class to test launching and closing of applications."""
    
    def __init__(self, config_path: str):
        self.load_config(config_path)
        self.configure_logging()
        self.shortcuts: List[str] = []
        self.executables: List[str] = []
        self.results: List[dict] = []
        self.testing_paused = False
        self.testing_cancelled = False
        self.progress_window = None
        self.progress_label = None
        self.progress_bar = None
        self.log_text_widget = None
        self.text_handler = None
        self.pause_button = None
        self.cancel_button = None

    def load_config(self, config_path: str):
        """Load configuration from a YAML file."""
        try:
            with open(config_path, 'r') as f:
                self.config = yaml.safe_load(f)
            self.excluded_processes = [proc.lower() for proc in self.config.get('excluded_processes', [])]
        except FileNotFoundError:
            messagebox.showerror("Error", f"Configuration file not found: {config_path}")
            sys.exit(1)
        except yaml.YAMLError as e:
            messagebox.showerror("Error", f"Error parsing configuration file: {e}")
            sys.exit(1)

    def configure_logging(self):
        """Set up logging based on configuration."""
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(self.config.get('log_file', 'application_test.log'), mode='w'),
                logging.StreamHandler(sys.stdout)
            ]
        )

    def load_files(self, shortcuts_file_path: str, executables_file_path: str) -> bool:
        """Load shortcuts and executables from provided file paths."""
        try:
            if not os.path.exists(shortcuts_file_path):
                logging.error(f"Shortcuts file not found at {shortcuts_file_path}.")
                messagebox.showerror("Error", f"Shortcuts file not found at {shortcuts_file_path}.")
                return False

            if not os.path.exists(executables_file_path):
                logging.error(f"Executables file not found at {executables_file_path}.")
                messagebox.showerror("Error", f"Executables file not found at {executables_file_path}.")
                return False

            with open(shortcuts_file_path, 'r', encoding='utf-8') as f:
                shortcuts = [line.strip() for line in f if line.strip()]

            with open(executables_file_path, 'r', encoding='utf-8') as f:
                executables = [line.strip() for line in f if line.strip()]

            # Filter out folder entries and ensure lengths match
            filtered_shortcuts = []
            filtered_executables = []

            for shortcut, executable in zip(shortcuts, executables):
                if not shortcut.startswith("[Folder]"):
                    filtered_shortcuts.append(shortcut)
                    filtered_executables.append(executable)

            if len(filtered_shortcuts) != len(filtered_executables):
                logging.error("The number of shortcuts and executables does not match after filtering.")
                messagebox.showerror("Error", "The number of shortcuts and executables does not match after filtering.")
                return False

            self.shortcuts = filtered_shortcuts
            self.executables = filtered_executables
            logging.info("Loaded shortcuts and executables files successfully.")
            return True

        except Exception as e:
            logging.exception(f"Error loading files: {e}")
            messagebox.showerror("Error", f"An error occurred while loading files: {e}")
            return False

    def create_progress_window(self, total_apps: int):
        """Create a GUI window to display progress and logs."""
        self.progress_window = tk.Toplevel()
        self.progress_window.title("Application Testing Progress")
        self.progress_window.geometry("500x400+1200+50")
        self.progress_window.attributes('-topmost', True)
        self.progress_window.resizable(False, False)

        self.progress_label = tk.Label(self.progress_window, text="Starting tests...", font=("Helvetica", 10))
        self.progress_label.pack(padx=10, pady=5)

        self.progress_bar = ttk.Progressbar(self.progress_window, length=380, mode='determinate', maximum=total_apps)
        self.progress_bar.pack(padx=10, pady=5)

        # Add ScrolledText widget for log display
        self.log_text_widget = ScrolledText(self.progress_window, width=58, height=15, state='disabled', font=("Courier", 9))
        self.log_text_widget.pack(padx=10, pady=5)

        # Add pause and cancel buttons
        button_frame = tk.Frame(self.progress_window)
        button_frame.pack(pady=5)

        self.pause_button = tk.Button(button_frame, text="Pause", command=self.pause_testing)
        self.pause_button.pack(side=tk.LEFT, padx=5)

        self.cancel_button = tk.Button(button_frame, text="Cancel", command=self.cancel_testing)
        self.cancel_button.pack(side=tk.LEFT, padx=5)

    def pause_testing(self):
        """Toggle the paused state of the testing."""
        self.testing_paused = not self.testing_paused
        if self.testing_paused:
            self.pause_button.config(text="Resume")
            logging.info("Testing paused.")
        else:
            self.pause_button.config(text="Pause")
            logging.info("Testing resumed.")

    def cancel_testing(self):
        """Cancel the testing process."""
        if messagebox.askyesno("Cancel Testing", "Are you sure you want to cancel the testing process?"):
            self.testing_cancelled = True
            logging.info("Testing cancelled by user.")

    def update_progress_window(self, current_app_index: int, app_name: str):
        """Update the progress bar and label."""
        self.progress_label.config(text=f"Testing application {current_app_index}/{int(self.progress_bar['maximum'])}: {app_name}")
        self.progress_bar['value'] = current_app_index
        self.progress_window.update_idletasks()

    def log_to_text_widget(self, message: str):
        """Log messages to the GUI text widget."""
        if self.log_text_widget is not None:
            self.log_text_widget.configure(state='normal')
            self.log_text_widget.insert(tk.END, message + '\n')
            self.log_text_widget.see(tk.END)
            self.log_text_widget.configure(state='disabled')
            self.progress_window.update_idletasks()

    class TextWidgetHandler(logging.Handler):
        """Custom logging handler that writes to the text widget."""
        def __init__(self, app_tester):
            super().__init__()
            self.app_tester = app_tester

        def emit(self, record):
            msg = self.format(record)
            self.app_tester.log_to_text_widget(msg)

    def is_system_process(self, proc_name: str) -> bool:
        """Check if a process is a system process."""
        return proc_name.lower() in self.excluded_processes

    def handle_uac_prompt(self) -> bool:
        """Check for and handle UAC prompts."""
        try:
            for proc in psutil.process_iter(['pid', 'name']):
                if proc.info['name'].lower() == 'consent.exe':
                    logging.warning("Detected UAC prompt (Consent.exe is running).")
                    return True
            return False
        except Exception as e:
            logging.error(f"Error detecting UAC prompt: {e}")
            return False

    def kill_process_tree(self, pid: int, including_parent: bool = True) -> List[str]:
        """Terminate a process and its child processes."""
        terminated_executables = []
        try:
            parent = psutil.Process(pid)
            parent_name = parent.name()
            if self.is_system_process(parent_name):
                logging.warning(f"Skipping termination of system process: PID {parent.pid}, Name {parent_name}")
                return terminated_executables

            children = parent.children(recursive=True)
            for child in children:
                try:
                    child_name = child.name()
                    if self.is_system_process(child_name):
                        logging.warning(f"Skipping termination of system process: PID {child.pid}, Name {child_name}")
                        continue
                    logging.info(f"Terminating child process: PID {child.pid}, Name {child_name}")
                    child.terminate()
                    terminated_executables.append(child_name)
                except psutil.NoSuchProcess:
                    continue
                except psutil.AccessDenied:
                    logging.error(f"Access denied when attempting to terminate process PID {child.pid}")
                    continue

            gone, still_alive = psutil.wait_procs(children, timeout=5)
            for child in still_alive:
                try:
                    child_name = child.name()
                    if self.is_system_process(child_name):
                        logging.warning(f"Skipping killing of system process: PID {child.pid}, Name {child_name}")
                        continue
                    logging.info(f"Killing child process: PID {child.pid}, Name {child_name}")
                    child.kill()
                    terminated_executables.append(child_name)
                except psutil.NoSuchProcess:
                    continue
                except psutil.AccessDenied:
                    logging.error(f"Access denied when attempting to kill process PID {child.pid}")
                    continue

            if including_parent:
                try:
                    logging.info(f"Terminating parent process: PID {parent.pid}, Name {parent_name}")
                    parent.terminate()
                    parent.wait(5)
                    terminated_executables.append(parent_name)
                except psutil.NoSuchProcess:
                    pass
                except psutil.AccessDenied:
                    logging.error(f"Access denied when attempting to terminate parent process PID {parent.pid}")

        except psutil.NoSuchProcess:
            pass
        except Exception as e:
            logging.exception(f"Error terminating process tree: {e}")
        return terminated_executables

    def launch_and_test_application(self, shortcut_path: str, expected_exe_path: str, app_name: str) -> dict:
        """Launch and test an individual application."""
        shortcut_name = os.path.basename(shortcut_path)
        expected_executable_name = os.path.basename(expected_exe_path).lower()

        result = {
            'Name': app_name,
            'Shortcut Path': shortcut_name,
            'Expected Executable': expected_executable_name,
            'Status': 'Not Tested',
            'Remarks': '',
            'Associated Windows': '',
            'Terminated Executables': '',
            'Closed Windows': ''
        }

        # Mapping of expected executable names to actual executable names
        executable_aliases = {
            'cmd.exe': 'WindowsTerminal.exe',
            'powershell.exe': 'WindowsTerminal.exe',
            'wmplayer.exe': 'setup_wm.exe',
            # Add other mappings if necessary
        }

        actual_executable_name = executable_aliases.get(expected_executable_name, expected_executable_name)

        associated_window_titles = []
        terminated_executables = []
        closed_windows = []

        logging.info(f"Starting test for application: {app_name}")

        if not os.path.exists(shortcut_path):
            result['Status'] = 'Failed'
            result['Remarks'] = f'Shortcut not found: {shortcut_name}'
            logging.error(f"Shortcut not found: {shortcut_name}")
            return result

        try:
            # Record initial processes and windows
            processes_before = {p.pid for p in psutil.process_iter(['pid'])}
            windows_before = set(w.handle for w in Desktop(backend="uia").windows())
            logging.debug("Recorded initial processes and windows.")

            # Start the application
            os.startfile(shortcut_path)
            logging.info(f"Launched application using shortcut: {shortcut_name}")

            # Wait for the application window
            start_time = time.time()
            application_window_found = False
            main_app_pid = None

            max_wait_time = self.config.get('max_wait_time', 20)
            poll_interval = self.config.get('poll_interval', 2)
            pause_after_found = self.config.get('pause_after_found', 2)
            additional_wait_time = self.config.get('additional_wait_time', 5)

            while time.time() - start_time < max_wait_time:
                if self.testing_cancelled:
                    logging.info("Testing process was cancelled by the user.")
                    result['Status'] = 'Cancelled'
                    result['Remarks'] = 'Testing cancelled by user.'
                    return result

                while self.testing_paused:
                    time.sleep(1)

                if self.handle_uac_prompt():
                    logging.warning("Application triggered UAC prompt.")
                    result.update({
                        'Status': 'Manual Intervention Required',
                        'Remarks': 'Application triggered UAC prompt.',
                        'Associated Windows': 'User Account Control',
                        'Terminated Executables': '; '.join(set(terminated_executables)) or 'None',
                        'Closed Windows': '; '.join(set(closed_windows)) or 'None'
                    })
                    return result

                windows_after = set(w.handle for w in Desktop(backend="uia").windows())
                new_windows_handles = windows_after - windows_before

                # Try to find the expected window
                for handle in new_windows_handles:
                    try:
                        window = Desktop(backend="uia").window(handle=handle)
                        window_title = window.window_text()
                        process_id = window.process_id()
                        process = psutil.Process(process_id)
                        exe_path = process.exe()

                        if os.path.basename(exe_path).lower() == actual_executable_name.lower() or \
                           re.search(re.escape(app_name), window_title, re.IGNORECASE):
                            application_window_found = True
                            expected_window = window
                            main_app_pid = process_id
                            associated_window_titles.append(window_title)
                            logging.info(f"Detected application window: {window_title}")
                            time.sleep(pause_after_found)
                            break
                    except Exception:
                        continue

                if application_window_found:
                    break

                time.sleep(poll_interval)

            if not application_window_found:
                # Handle applications without windows
                logging.warning("No application window detected. Attempting to handle background processes.")
                processes_after = {p.pid for p in psutil.process_iter(['pid'])}
                new_pids = processes_after - processes_before

                for pid in new_pids:
                    try:
                        proc = psutil.Process(pid)
                        proc_name = proc.name()
                        if proc_name.lower() == actual_executable_name.lower():
                            main_app_pid = pid
                            logging.info(f"Detected background process: PID {pid}, Name {proc_name}")
                            break
                    except psutil.NoSuchProcess:
                        continue

                if main_app_pid is None:
                    result['Status'] = 'Failed'
                    result['Remarks'] = 'Application did not open any windows or detectable processes.'
                    logging.error(result['Remarks'])
                    return result

            # Additional wait time
            time.sleep(additional_wait_time)

            # Close application window if found
            if application_window_found:
                try:
                    expected_window.close()
                    closed_windows.append(expected_window.window_text())
                    logging.info(f"Closed application window: {expected_window.window_text()}")
                except Exception as e:
                    logging.error(f"Error closing application window: {e}")

            # Terminate the application processes
            if main_app_pid:
                terminated_execs = self.kill_process_tree(main_app_pid)
                terminated_executables.extend(terminated_execs)

            # Wait for processes to terminate
            time.sleep(2)

            # Check for residual processes
            processes_after = {p.pid for p in psutil.process_iter(['pid'])}
            new_pids = processes_after - processes_before

            for pid in new_pids:
                try:
                    proc = psutil.Process(pid)
                    proc_name = proc.name()
                    if self.is_system_process(proc_name):
                        continue
                    logging.warning(f"Residual process detected: PID {pid}, Name {proc_name}")
                    proc.terminate()
                    terminated_executables.append(proc_name)
                except psutil.NoSuchProcess:
                    continue
                except psutil.AccessDenied:
                    logging.error(f"Access denied when attempting to terminate process PID {pid}")

            # Check for residual windows
            windows_after = set(w.handle for w in Desktop(backend="uia").windows())
            new_windows = windows_after - windows_before

            for handle in new_windows:
                try:
                    window = Desktop(backend="uia").window(handle=handle)
                    window_title = window.window_text()
                    logging.warning(f"Residual window detected: {window_title}")
                    window.close()
                    closed_windows.append(window_title)
                except Exception:
                    continue

            # Update result
            result.update({
                'Status': 'Success',
                'Remarks': 'Application tested successfully.',
                'Associated Windows': '; '.join(associated_window_titles) or 'None',
                'Terminated Executables': '; '.join(set(terminated_executables)) or 'None',
                'Closed Windows': '; '.join(set(closed_windows)) or 'None'
            })

        except Exception as e:
            result['Status'] = 'Failed'
            result['Remarks'] = f'Exception occurred: {e}'
            logging.exception(f"Exception during testing: {e}")

        return result

    def save_results_to_excel(self, results: List[dict], file_path: str):
        """Save test results to an Excel file."""
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Application Test Results"

            headers = [
                'Name', 'Shortcut Path', 'Expected Executable', 'Associated Windows',
                'Terminated Executables', 'Closed Windows', 'Status', 'Remarks'
            ]
            ws.append(headers)

            for result in results:
                ws.append([
                    result.get('Name', ''),
                    result.get('Shortcut Path', ''),
                    result.get('Expected Executable', ''),
                    result.get('Associated Windows', ''),
                    result.get('Terminated Executables', ''),
                    result.get('Closed Windows', ''),
                    result.get('Status', ''),
                    result.get('Remarks', '')
                ])

            # Adjust column widths
            for col in ws.columns:
                max_length = max(len(str(cell.value)) for cell in col) + 2
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max_length

            # Set alignment
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(vertical='top', horizontal='left', wrap_text=True)

            wb.save(file_path)
            logging.info(f"Results saved to Excel file: {file_path}")

        except Exception as e:
            logging.error(f"Error saving results to Excel: {e}")

    def run_tests(self):
        """Run the application tests."""
        try:
            total_apps = len(self.shortcuts)
            self.create_progress_window(total_apps)

            # Set up custom logging handler after GUI is initialized
            self.text_handler = self.TextWidgetHandler(self)
            formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            self.text_handler.setFormatter(formatter)
            logging.getLogger().addHandler(self.text_handler)

            for index, (shortcut_path, exe_path) in enumerate(zip(self.shortcuts, self.executables), start=1):
                if self.testing_cancelled:
                    logging.info("Testing process was cancelled by the user.")
                    break

                while self.testing_paused:
                    time.sleep(1)

                app_name = os.path.basename(shortcut_path).replace('.lnk', '')
                logging.info(f"Testing application: {app_name}")

                self.update_progress_window(index, app_name)

                result = self.launch_and_test_application(shortcut_path, exe_path, app_name)
                self.results.append(result)
                time.sleep(1)

            # Close progress window
            self.progress_window.destroy()
            # Remove the custom logging handler
            logging.getLogger().removeHandler(self.text_handler)

            # Save results
            self.save_results_to_excel(self.results, self.config.get('excel_output', 'Application_Test_Results.xlsx'))
            logging.info("Testing completed.")
            messagebox.showinfo("Completed", f"Testing completed. Results saved to {self.config.get('excel_output', 'Application_Test_Results.xlsx')}")

        except Exception as e:
            logging.exception(f"Error during testing: {e}")
            messagebox.showerror("Error", f"An error occurred during testing: {e}")

    def main(self):
        """Main function to execute the application testing."""
        root = tk.Tk()
        root.withdraw()

        # Prompt user for files
        shortcuts_file = filedialog.askopenfilename(
            title="Select Shortcuts File",
            filetypes=(("Text files", "*.txt"),)
        )
        if not shortcuts_file:
            logging.error("No shortcuts file selected.")
            messagebox.showerror("Error", "No shortcuts file selected.")
            sys.exit(1)

        executables_file = filedialog.askopenfilename(
            title="Select Executables File",
            filetypes=(("Text files", "*.txt"),)
        )
        if not executables_file:
            logging.error("No executables file selected.")
            messagebox.showerror("Error", "No executables file selected.")
            sys.exit(1)

        # Load files
        if not self.load_files(shortcuts_file, executables_file):
            sys.exit(1)

        # Run tests in a separate thread to keep GUI responsive
        test_thread = threading.Thread(target=self.run_tests)
        test_thread.start()
        root.mainloop()

if __name__ == '__main__':
    tester = ApplicationTester('config.yaml')
    tester.main()
