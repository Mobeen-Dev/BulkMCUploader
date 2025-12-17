import glob
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
from typing import Callable, Optional
from threading import Thread
import customtkinter as ctk
import serial.tools.list_ports
import threading
import time
import os
import json
import sys
import subprocess
import queue
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import re

Callback = Callable[[str, str], None]

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


SETTINGS_FILE = resource_path("settings.json")
BOARDS_FILE = resource_path("custom_boards.json")
# Set the appearance and theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("dark-blue")

class ProgrammingResult:
    """Class to store programming results"""

    def __init__(self, port, success, message, duration, board_type, fqbn, baud_rate, logger):
        self.port = port
        self.success = success
        self.message = message
        self.duration = duration
        self.board_type = board_type
        self.fqbn = fqbn
        self.timestamp = datetime.now()

        # Start serial monitor thread (runs max 60 seconds)
        threading.Thread(
            target=self.monitor_serial_output,
            args=(logger, self.port, baud_rate),
            daemon=True
        ).start()

    @staticmethod
    def monitor_serial_output(logger, port, baud_rate=9600, timeout_seconds=120):
        """Continuously read serial data for up to 1 minute or until unplugged"""
        ser = None
        start_time = time.time()

        try:
            ser = serial.Serial(port, baud_rate, timeout=1)
            logger(f"📡 Started serial monitoring on {port} @ {baud_rate} baud", "MCU")

            while True:
                # Stop after timeout
                if time.time() - start_time > timeout_seconds:
                    logger(f"⏰ Timeout reached ({timeout_seconds}s). Stopping monitoring.", "MCU")
                    break

                try:
                    if ser.in_waiting:
                        line = ser.readline().decode(errors="ignore").strip()
                        if line:
                            logger(f"[{port}] {line}", "MCU")

                    time.sleep(0.05)

                except serial.SerialException as e:
                    logger(f"❌ Device on {port} disconnected: {e}", "ERROR")
                    break
                except Exception as e:
                    logger(f"⚠️ Unexpected error on {port}: {e}", "ERROR")
                    break

        except serial.SerialException as e:
            logger(f"❌ Could not open serial port {port}: {e}", "ERROR")

        finally:
            if ser and ser.is_open:
                try:
                    ser.close()
                    logger(f"🔌 Serial port {port} closed cleanly", "MCU")
                except Exception as e:
                    logger(f"⚠️ Error closing port {port}: {e}", "ERROR")

            logger(f"🛑 Monitoring thread for {port} exited", "MCU")

class ArduinoCLIProgrammer:
    """Arduino CLI based programmer for all supported boards"""

    def __init__(self, settings):
        self.settings = settings
        self.baud_rate = 9600
        self.timeout = settings.get("programming_timeout", 60)
        self.max_retries = settings.get("max_retries", 3)
        self.verify_after_program = settings.get("verify_programming", True)
        self.verbose = settings.get("verbose_output", False)

    def check_arduino_cli(self):
        """Check if Arduino CLI is installed and accessible"""
        try:
            result = subprocess.run(
                ["arduino-cli", "version"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )
            return result.returncode == 0, result.stdout.strip()
        except (subprocess.TimeoutExpired, FileNotFoundError):
            return False, "Arduino CLI not found"

    def install_core_if_needed(self, fqbn):
        """Install board core if not already installed"""
        try:
            # Extract core from FQBN (e.g., "esp32:esp32" from "esp32:esp32:esp32")
            core = ":".join(fqbn.split(":")[:2])

            # Check if core is installed
            result = subprocess.run(
                ["arduino-cli", "core", "list"],
                capture_output=True,
                text=True,
                timeout=30,
                creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            if core not in result.stdout:
                self.log_callback(f"Installing core: {core}", "INFO")
                install_result = subprocess.run(
                    ["arduino-cli", "core", "install", core],
                    capture_output=True,
                    text=True,
                    timeout=300,  # 5 minutes for core installation
                    creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
                )
                return install_result.returncode == 0, install_result.stderr

            return True, "Core already installed"

        except Exception as e:
            return False, str(e)

    def update_core_index(self):
        """Update Arduino CLI core index"""
        try:
            result = subprocess.run(
                ["arduino-cli", "core", "update-index"],
                capture_output=True,
                text=True,
                timeout=60,
                creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )
            return result.returncode == 0, result.stderr
        except Exception as e:
            return False, str(e)

    def compile_sketch(self, sketch_path, fqbn, output_dir=None):
        """Compile Arduino sketch"""
        try:
            cmd = ["arduino-cli", "compile", "--fqbn", fqbn]

            if output_dir:
                cmd.extend(["--output-dir", output_dir])

            if self.verbose:
                cmd.append("--verbose")
            else:
                cmd.append("--quiet")

            cmd.append(sketch_path)

            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=self.timeout, creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            return result.returncode == 0, result.stdout, result.stderr

        except subprocess.TimeoutExpired:
            return False, "", f"Compilation timed out after {self.timeout} seconds"
        except Exception as e:
            return False, "", str(e)

    def upload_compiled(self, compiled_path, fqbn, port):
        """Upload pre-compiled firmware to board"""
        try:
            cmd = [
                "arduino-cli",
                "upload",
                "--fqbn",
                fqbn,
                "--port",
                port,
                "--input-file",
                compiled_path,
            ]

            if self.verbose:
                cmd.append("--verbose")

            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=self.timeout, creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            return result.returncode == 0, result.stdout, result.stderr

        except subprocess.TimeoutExpired:
            return False, "", f"Upload timed out after {self.timeout} seconds"
        except Exception as e:
            return False, "", str(e)

    # below function must be changed it will used above two and conditionally compiled if no compiled binary not present if present just push
    def upload_sketch(self, sketch_path, fqbn, port):
        """Compile and upload Arduino sketch in one step"""
        try:
            cmd = ["arduino-cli", "upload", "--fqbn", fqbn, "--port", port]

            if self.verify_after_program:
                cmd.append("--verify")

            if self.verbose:
                cmd.append("--verbose")

            cmd.append(sketch_path)

            result = subprocess.run(
                cmd, capture_output=True, text=True, timeout=self.timeout, creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            return result.returncode == 0, result.stdout, result.stderr

        except subprocess.TimeoutExpired:
            return False, "", f"Upload timed out after {self.timeout} seconds"
        except Exception as e:
            return False, "", str(e)

    def detect_board(self, port):
        """Detect connected board type"""
        try:
            result = subprocess.run(
                ["arduino-cli", "board", "list"],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            if result.returncode == 0:
                lines = result.stdout.split("\n")
                for line in lines:
                    if port in line:
                        # Parse the line to extract FQBN
                        parts = line.split()
                        if len(parts) >= 3:
                            return True, parts[2]  # FQBN is typically the 3rd column

            return False, "Board not detected"

        except Exception as e:
            return False, str(e)

    def list_connected_boards(self):
        """List all connected boards"""
        try:
            result = subprocess.run(
                ["arduino-cli", "board", "list"],
                capture_output=True,
                text=True,
                timeout=15,
                creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
            )

            boards = []
            if result.returncode == 0:
                lines = result.stdout.split("\n")[1:]  # Skip header
                for line in lines:
                    if line.strip() and not line.startswith("Port"):
                        parts = line.split()
                        if len(parts) >= 3:
                            port = parts[0]
                            protocol = parts[1]
                            fqbn = parts[2] if parts[2] != "Unknown" else ""
                            board_name = (
                                " ".join(parts[3:]) if len(parts) > 3 else "Unknown"
                            )

                            boards.append(
                                {
                                    "port": port,
                                    "protocol": protocol,
                                    "fqbn": fqbn,
                                    "board_name": board_name,
                                }
                            )

            return True, boards

        except Exception as e:
            return False, str(e)

    def compile_only_mode(self):
        """Compile sketch and save binaries for production use"""
        if not self.board_var.get() or not self.file_var.get():
            messagebox.showerror("Error", "Select board and sketch file first!")
            return

        # Create output directory
        sketch_name = os.path.splitext(os.path.basename(self.file_var.get()))[0]
        output_dir = os.path.join(
            os.path.dirname(self.file_var.get()), f"{sketch_name}_compiled"
        )
        os.makedirs(output_dir, exist_ok=True)

        # Compile sketch
        board_config = self.supported_boards[self.board_var.get()]
        fqbn = board_config["fqbn"]

        success, stdout, stderr = self.programmer.compile_sketch(
            self.file_var.get(), fqbn, output_dir
        )

        if success:
            # Find generated .bin file
            bin_files = glob.glob(os.path.join(output_dir, "*.bin"))
            if bin_files:
                self.compiled_firmware_path = bin_files[0]  # Store for production use
                self.log_message(
                    f"✅ Compiled successfully: {os.path.basename(bin_files[0])}",
                    "SUCCESS",
                )
                self.log_message(f"📁 Saved to: {output_dir}", "INFO")
            else:
                self.log_message(
                    "⚠️ Compilation succeeded but no .bin file found", "WARNING"
                )
        else:
            self.log_message(f"❌ Compilation failed: {stderr}", "ERROR")

    def program_board(self, firmware_path, fqbn, port):
        """Main method to program a board"""
        # Ensure core is installed
        core_success, core_message = self.install_core_if_needed(fqbn)
        if not core_success:
            return False, f"Failed to install core: {core_message}"

        # Determine file type and programming method
        file_ext = os.path.splitext(firmware_path)[1].lower()

        if file_ext in [".ino", ".pde"]:
            # Arduino sketch - compile and upload
            success, stdout, stderr = self.upload_sketch(firmware_path, fqbn, port)
        elif file_ext in [".hex", ".bin"]:
            # Pre-compiled firmware - upload directly
            success, stdout, stderr = self.upload_compiled(firmware_path, fqbn, port)
        else:
            return False, f"Unsupported file type: {file_ext}"

        if success:
            return True, "Programming completed successfully"
        else:
            error_msg = stderr if stderr else stdout
            return False, f"Programming failed: {error_msg}"

    def set_log_callback(self, callback):
        """Set callback function for logging"""
        self.log_callback = callback


class MCUProgrammerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Arduino CLI Production Programmer")
        self.root.geometry("1000x900")

        try:
            self.root.iconbitmap(resource_path("microchip.ico"))
        except:
            pass

        # Initialize variables
        self.saved_data = {}
        self.programming_active = False
        self.stop_programming = False
        self.programming_queue = queue.Queue()
        self.result_queue = queue.Queue()
        self.executor = None

        # Supported boards with Arduino CLI FQBN (Fully Qualified Board Name)
        self.supported_boards = {
            "Custom": {"fqbn": "- Developer Options -", "core": "---"},
            "Arduino Uno": {"fqbn": "arduino:avr:uno", "core": "arduino:avr"},
            "Arduino Nano (Old Bootloader)": {
                "fqbn": "arduino:avr:nano:cpu=atmega328old",
                "core": "arduino:avr",
            },
            "Arduino Nano (New Bootloader)": {
                "fqbn": "arduino:avr:nano:cpu=atmega328",
                "core": "arduino:avr",
            },
            "Arduino Mega": {
                "fqbn": "arduino:avr:mega:cpu=atmega2560",
                "core": "arduino:avr",
            },
            "Arduino Leonardo": {"fqbn": "arduino:avr:leonardo", "core": "arduino:avr"},
            "Arduino Micro": {"fqbn": "arduino:avr:micro", "core": "arduino:avr"},
            "ESP32 Dev Module": {"fqbn": "esp32:esp32:esp32", "core": "esp32:esp32"},
            "ESP32-S2": {"fqbn": "esp32:esp32:esp32s2", "core": "esp32:esp32"},
            "ESP32-S3": {"fqbn": "esp32:esp32:esp32s3", "core": "esp32:esp32"},
            "ESP32-C3": {"fqbn": "esp32:esp32:esp32c3", "core": "esp32:esp32"},
            "ESP8266 NodeMCU": {
                "fqbn": "esp8266:esp8266:nodemcuv2",
                "core": "esp8266:esp8266",
            },
            "ESP8266 Wemos D1": {
                "fqbn": "esp8266:esp8266:d1_mini",
                "core": "esp8266:esp8266",
            },
            "Raspberry Pi Pico": {
                "fqbn": "rp2040:rp2040:rpipico",
                "core": "rp2040:rp2040",
            },
            "Adafruit Feather M0": {
                "fqbn": "adafruit:samd:adafruit_feather_m0",
                "core": "adafruit:samd",
            },
            "Arduino MKR WiFi 1010": {
                "fqbn": "arduino:samd:mkrwifi1010",
                "core": "arduino:samd",
            },
            "Teensy 3.2": {"fqbn": "teensy:avr:teensy31", "core": "teensy:avr"},
            "Teensy 4.0": {"fqbn": "teensy:avr:teensy40", "core": "teensy:avr"},
        }
        if os.path.exists(BOARDS_FILE):
            try:
                with open(BOARDS_FILE, "r") as f:
                    custom_boards = json.load(f)

                self.supported_boards.update(custom_boards)  # merge into existing

            except Exception as e:
                self.log_message(f"⚠️ Failed to load custom boards: {str(e)}", "WARNING")

        # Statistics
        self.stats = {
            "total_programmed": 0,
            "successful": 0,
            "failed": 0,
            "start_time": None,
            "session_results": [],
        }

        # Default settings
        self.default_settings = {
            "programming_timeout": 60,
            "max_retries": 3,
            "verify_programming": True,
            "verbose_output": False,
            "max_parallel_jobs": 4,
            "port_scan_interval": 5,
            "auto_detect_devices": True,
            "log_detailed_output": False,
            "max_devices_per_hub": 2,
            "hub_programming_delay": 2,
            "auto_install_cores": True,
            "update_index_on_start": True,
            "programming_mode": "Binary",  # Default to Binary mode
            "execution_mode": "Continuous",  # Default to Continuous mode
        }

        # Create Arduino CLI programmer
        self.programmer = None

        # Load settings and create UI
        self.load_settings()
        self.create_widgets()

        # Initialize Arduino CLI
        self.initialize_arduino_cli()

        # Start background tasks
        self.start_port_monitoring()
        self.start_result_processor()

    def initialize_arduino_cli(self):
        """Initialize Arduino CLI programmer and check installation"""
        self.programmer = ArduinoCLIProgrammer(self.saved_data)
        self.programmer.set_log_callback(self.log_message)

        # Check if Arduino CLI is installed
        cli_available, version_info = self.programmer.check_arduino_cli()

        if cli_available:
            self.log_message(f"✅ Arduino CLI detected: {version_info}", "SUCCESS")

            # Update core index if enabled
            if self.saved_data.get("update_index_on_start", True):
                self.log_message("🔄 Updating Arduino CLI core index...", "INFO")
                threading.Thread(
                    target=self.update_core_index_async, daemon=True
                ).start()
        else:
            self.log_message(
                "❌ Arduino CLI not found! Please install Arduino CLI", "ERROR"
            )
            self.log_message(
                "📥 Download from: https://arduino.github.io/arduino-cli/", "INFO"
            )

            # Disable programming controls
            self.start_button.configure(state="disabled")

    def update_core_index_async(self):
        """Update core index in background"""
        success, message = self.programmer.update_core_index()
        if success:
            self.log_message("✅ Core index updated successfully", "SUCCESS")
        else:
            self.log_message(f"⚠️ Core index update warning: {message}", "WARNING")

    def compile_only_mode(self):
        """Compile sketch and save binaries for production use"""
        if not self.board_var.get():
            messagebox.showerror("Error", "Please select a target board!")
            return

        if not self.file_var.get() or not os.path.exists(self.file_var.get()):
            messagebox.showerror("Error", "Please select a valid sketch file!")
            return

        if not self.programmer:
            messagebox.showerror("Error", "Arduino CLI not initialized!")
            return

        # Check if it's a sketch file
        file_ext = os.path.splitext(self.file_var.get())[1].lower()
        if file_ext not in [".ino", ".pde"]:
            messagebox.showerror(
                "Error", "Please select an Arduino sketch (.ino) file for compilation!"
            )
            return

        self.log_message("🔨 Starting compilation...", "INFO")

        def compile_async():
            try:
                # Create output directory
                sketch_name = os.path.splitext(os.path.basename(self.file_var.get()))[0]
                sketch_dir = os.path.dirname(self.file_var.get())
                output_dir = os.path.join(sketch_dir, f"{sketch_name}_compiled")
                os.makedirs(output_dir, exist_ok=True)

                # Get board configuration
                board_config = self.supported_boards[self.board_var.get()]
                fqbn = board_config["fqbn"]

                self.log_message(f"📋 Board: {self.board_var.get()} ({fqbn})", "INFO")
                self.log_message(f"📁 Output: {output_dir}", "INFO")

                # Ensure core is installed
                core_success, core_message = self.programmer.install_core_if_needed(
                    fqbn
                )
                if not core_success:
                    self.log_message(
                        f"❌ Core installation failed: {core_message}", "ERROR"
                    )
                    return

                # Compile sketch
                success, stdout, stderr = self.programmer.compile_sketch(
                    self.file_var.get(), fqbn, output_dir
                )

                if success:
                    # Find generated .bin files
                    bin_files = glob.glob(os.path.join(output_dir, "*.bin"))

                    if bin_files:
                        self.log_message(f"✅ Compilation successful!", "SUCCESS")
                        for bin_file in bin_files:
                            file_size = os.path.getsize(bin_file) / 1024  # Size in KB
                            self.log_message(
                                f"📦 Generated: {os.path.basename(bin_file)} ({file_size:.1f} KB)",
                                "SUCCESS",
                            )

                        # Store the main binary path for potential use
                        main_bin = next(
                            (f for f in bin_files if sketch_name in f), bin_files[0]
                        )
                        self.compiled_firmware_path = main_bin

                        self.log_message(
                            f"💾 Ready for production programming!", "SUCCESS"
                        )
                        self.log_message(
                            f"📂 Compiled files saved to: {os.path.basename(output_dir)}",
                            "INFO",
                        )

                        # Show success message
                        messagebox.showinfo(
                            "Compilation Complete",
                            f"Sketch compiled successfully!\n\n"
                            f"Binary files saved to:\n{output_dir}\n\n"
                            f"You can now use these .bin files for fast production programming.",
                        )

                    else:
                        self.log_message(
                            "⚠️ Compilation succeeded but no .bin files found", "WARNING"
                        )
                        self.log_message(
                            f"📂 Check output directory: {output_dir}", "INFO"
                        )

                else:
                    self.log_message(f"❌ Compilation failed!", "ERROR")
                    if stderr:
                        # Show first few lines of error
                        error_lines = stderr.split("\n")[:5]
                        for line in error_lines:
                            if line.strip():
                                self.log_message(f"   {line}", "ERROR")

            except Exception as e:
                self.log_message(f"❌ Compilation error: {str(e)}", "ERROR")

        # Run compilation in background thread
        threading.Thread(target=compile_async, daemon=True).start()

    def create_widgets(self):
        # Main container with padding
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(padx=10, pady=10, fill="both", expand=True)

        # Configuration section
        config_frame = ctk.CTkFrame(main_frame)
        config_frame.pack(padx=10, pady=(10, 5), fill="x")

        # Row 0: Board selection
        board_label = ctk.CTkLabel(config_frame, text="Target Board:")
        board_label.grid(row=0, column=0, sticky="w", pady=5, padx=10)

        self.board_var = ctk.StringVar()
        self.board_combo = ctk.CTkComboBox(
            config_frame,
            variable=self.board_var,
            values=list(self.supported_boards.keys()),
            width=250,
            command=self.on_board_changed,
        )
        self.board_combo.grid(row=0, column=1, pady=5, padx=10, sticky="w")

        # Board info display
        self.board_info_label = ctk.CTkLabel(config_frame, text="", text_color="gray")
        self.board_info_label.grid(row=0, column=2, pady=5, padx=10, sticky="w")

        # Auto-detect button
        detect_button = ctk.CTkButton(
            config_frame,
            text="🔍 Auto-Detect",
            command=self.auto_detect_boards,
            width=100,
        )
        detect_button.grid(row=0, column=3, pady=5, padx=10)

        # Row 1: Port selection
        port_label = ctk.CTkLabel(config_frame, text="Port:")
        port_label.grid(row=1, column=0, sticky="w", pady=5, padx=10)

        self.port_var = ctk.StringVar()
        self.port_combo = ctk.CTkComboBox(
            config_frame,
            variable=self.port_var,
            values=["All"] + self.get_ports(),
            width=200,
        )
        self.port_combo.set("All")
        self.port_combo.grid(row=1, column=1, pady=5, padx=10, sticky="w")

        refresh_button = ctk.CTkButton(
            config_frame, text="🔄 Refresh", command=self.refresh_ports, width=80
        )
        refresh_button.grid(row=1, column=2, pady=5, padx=10, sticky="w")

        # Connected devices display
        self.connected_label = ctk.CTkLabel(
            config_frame, text="Connected: 0 devices", text_color="blue"
        )
        self.connected_label.grid(row=1, column=3, pady=5, padx=10)

        # Row 2: Firmware file selection
        file_label = ctk.CTkLabel(config_frame, text="Firmware/Sketch:")
        file_label.grid(row=2, column=0, sticky="w", pady=5, padx=10)

        self.file_var = ctk.StringVar()
        self.file_entry = ctk.CTkEntry(
            config_frame, textvariable=self.file_var, width=400
        )
        self.file_entry.grid(
            row=2, column=1, columnspan=2, pady=5, padx=10, sticky="ew"
        )

        browse_button = ctk.CTkButton(
            config_frame, text="📁 Browse", command=self.browse_file, width=80
        )
        browse_button.grid(row=2, column=3, pady=5, padx=10, sticky="w")

        # Configure grid weights
        config_frame.grid_columnconfigure(1, weight=1)

        # Control buttons section
        control_frame = ctk.CTkFrame(main_frame)
        control_frame.pack(padx=10, pady=5, fill="x")

        # Programming controls
        self.start_button = ctk.CTkButton(
            control_frame,
            text=" ▶ Start Programming",
            command=self.start_programming,
            fg_color="green",
            hover_color="#006600",
            font=("Arial", 14, "bold"),
            height=40,
        )
        self.start_button.pack(side="left", padx=10, pady=10)

        self.stop_button = ctk.CTkButton(
            control_frame,
            text=" ⏹ Stop",
            command=self.stop_programming_action,
            fg_color="red4",
            hover_color="#cc0000",
            font=("Arial", 14, "bold"),
            height=40,
            state="disabled",
        )
        self.stop_button.pack(side="left", padx=10, pady=10)

        self.mode_var = "Continuous"
        # # Mode selection
        # self.mode_var = ctk.StringVar(value="Continuous")
        # mode_frame = ctk.CTkFrame(control_frame)
        # mode_frame.pack(side="left", padx=20, pady=10)
        #
        # ctk.CTkLabel(mode_frame, text="Mode:").pack(side="left", padx=5)
        # mode_radio1 = ctk.CTkRadioButton(mode_frame, text="Continuous", variable=self.mode_var, value="Continuous")
        # mode_radio1.pack(side="left", padx=5)
        # mode_radio2 = ctk.CTkRadioButton(mode_frame, text="Single Batch", variable=self.mode_var, value="Single")
        # mode_radio2.pack(side="left", padx=5)
        #
        # Settings and tools
        settings_button = ctk.CTkButton(
            control_frame, text="⚙️ Settings", command=self.open_settings, height=40
        )
        settings_button.pack(side="right", padx=10, pady=10)

        self.programming_mode = "Binary"
        # # Add mode selection
        # self.programming_mode = ctk.StringVar(value="Binary")
        # mode_frame = ctk.CTkFrame(control_frame)
        # mode_frame.pack(side="left", padx=20, pady=10)
        #
        # ctk.CTkLabel(mode_frame, text="Mode:").pack(side="left", padx=5)
        #
        # binary_only_radio = ctk.CTkRadioButton(mode_frame, text="Binary Only",
        #                                        variable=self.programming_mode, value="Binary")
        # binary_only_radio.pack(side="left", padx=5)
        #
        # compile_upload_radio = ctk.CTkRadioButton(mode_frame, text="Compile+Upload",
        #                                           variable=self.programming_mode, value="Compile+Upload")
        # compile_upload_radio.pack(side="left", padx=5)

        self.programming_mode = ctk.StringVar(
            value=self.saved_data.get("programming_mode", "Binary")
        )
        self.mode_var = ctk.StringVar(
            value=self.saved_data.get("execution_mode", "Continuous")
        )

        # Mode display labels (read-only display of current settings)
        mode_display_frame = ctk.CTkFrame(control_frame)
        mode_display_frame.pack(side="left", padx=20, pady=10)

        self.mode_display_label = ctk.CTkLabel(
            mode_display_frame,
            text=f"{self.mode_var.get()} | {self.programming_mode.get()}",
            font=("Arial", 15),
        )
        self.mode_display_label.pack(padx=10, pady=10)

        # Add compile button
        compile_button = ctk.CTkButton(
            control_frame,
            text="🔨 Compile Only",
            command=self.compile_only_mode,
            height=40,
        )
        compile_button.pack(side="left", padx=10, pady=10)

        test_button = ctk.CTkButton(
            control_frame, text="🔧 Test CLI", command=self.test_arduino_cli, height=40
        )
        test_button.pack(side="right", padx=10, pady=10)

        # Statistics frame
        stats_frame = ctk.CTkFrame(main_frame)
        stats_frame.pack(padx=10, pady=5, fill="x")

        stats_title = ctk.CTkLabel(
            stats_frame, text="📊 Programming Statistics", font=("Arial", 14, "bold")
        )
        stats_title.pack(pady=(10, 5))

        # Stats display
        stats_content_frame = ctk.CTkFrame(stats_frame)
        stats_content_frame.pack(padx=10, pady=(0, 10), fill="x")

        self.total_label = ctk.CTkLabel(stats_content_frame, text="Total: 0")
        self.total_label.grid(row=0, column=0, padx=10, pady=5)

        self.success_label = ctk.CTkLabel(
            stats_content_frame, text="Success: 0", text_color="green"
        )
        self.success_label.grid(row=0, column=1, padx=10, pady=5)

        self.failed_label = ctk.CTkLabel(
            stats_content_frame, text="Failed: 0", text_color="red"
        )
        self.failed_label.grid(row=0, column=2, padx=10, pady=5)

        self.rate_label = ctk.CTkLabel(stats_content_frame, text="Success Rate: 0%")
        self.rate_label.grid(row=0, column=3, padx=10, pady=5)

        self.time_label = ctk.CTkLabel(stats_content_frame, text="Runtime: 00:00:00")
        self.time_label.grid(row=0, column=4, padx=10, pady=5)

        # Throughput display
        self.throughput_label = ctk.CTkLabel(
            stats_content_frame, text="Throughput: 0/min"
        )
        self.throughput_label.grid(row=0, column=5, padx=10, pady=5)

        # Configure stats grid
        for i in range(6):
            stats_content_frame.grid_columnconfigure(i, weight=1)

        # # Terminal/Log section
        terminal_frame = ctk.CTkFrame(main_frame)
        terminal_frame.pack(padx=10, pady=5, fill="both", expand=True)

        terminal_title = ctk.CTkLabel(
            terminal_frame, text="💻 Programming Log", font=("Arial", 14, "bold")
        )
        terminal_title.pack(pady=(10, 5))

        # Terminal text area with horizontal scrollbar
        self.terminal = ctk.CTkTextbox(
            terminal_frame,
            width=850,
            height=250,
            font=("Consolas", 18),  # Reduced font size slightly for better readability
            wrap="none",  # ← KEY FIX: Disable word wrapping to enable horizontal scroll
        )
        self.terminal.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        # Terminal controls
        terminal_controls = ctk.CTkFrame(terminal_frame)
        terminal_controls.pack(padx=10, pady=(0, 10), fill="x")

        clear_button = ctk.CTkButton(
            terminal_controls, text="🗑️ Clear Log", command=self.clear_terminal
        )
        clear_button.pack(side="left", padx=5)

        save_log_button = ctk.CTkButton(
            terminal_controls, text="💾 Save Log", command=self.save_log
        )
        save_log_button.pack(side="left", padx=5)

        export_stats_button = ctk.CTkButton(
            terminal_controls, text="📊 Export Stats", command=self.export_statistics
        )
        export_stats_button.pack(side="left", padx=5)

        # Auto-scroll checkbox
        self.auto_scroll_var = ctk.BooleanVar(value=True)
        auto_scroll_check = ctk.CTkCheckBox(
            terminal_controls, text="Auto-scroll", variable=self.auto_scroll_var
        )
        auto_scroll_check.pack(side="right", padx=5)

        # Word wrap toggle checkbox (optional - gives users control)
        self.word_wrap_var = ctk.BooleanVar(value=False)  # Default to no wrap
        word_wrap_check = ctk.CTkCheckBox(
            terminal_controls,
            text="Word wrap",
            variable=self.word_wrap_var,
            command=self.toggle_word_wrap,
        )
        word_wrap_check.pack(side="right", padx=5)

        # Initialize with welcome message
        self.log_message("🚀 Arduino CLI Production Programmer initialized", "INFO")
        self.log_message(
            "Select target board, port, and firmware file to begin", "INFO"
        )

        # ============================================================================
        # ADD THIS NEW METHOD TO HANDLE WORD WRAP TOGGLE:
        # ============================================================================

    def toggle_word_wrap(self):
        """Toggle word wrapping in terminal"""
        if self.word_wrap_var.get():
            self.terminal.configure(wrap="word")  # Enable word wrapping
            self.log_message("📝 Word wrap enabled", "INFO")
        else:
            self.terminal.configure(
                wrap="none"
            )  # Disable word wrapping (horizontal scroll)
            self.log_message(
                "📝 Word wrap disabled - horizontal scroll enabled", "INFO"
            )

    def log_message(self, message, level="INFO"):
        """Add a message to the terminal log with thread safety"""

        def _log():
            timestamp = time.strftime("%H:%M:%S")

            # Format message with consistent spacing for better readability
            if level == "SUCCESS":
                log_entry = f"[{timestamp}] ✅ {message}\n"
            elif level == "MCU":
                log_entry = f" ──────➜ 💎 {message}\n"
            elif level == "ERROR":
                log_entry = f"[{timestamp}] ❌ {message}\n"
            elif level == "WARNING":
                log_entry = f"[{timestamp}] ⚠️  {message}\n"
            elif level == "INFO":
                log_entry = f"[{timestamp}] ℹ️  {message}\n"
            else:
                log_entry = f"[{timestamp}] {message}\n"

            # Insert into terminal
            self.terminal.insert("end", log_entry)

            # Auto-scroll if enabled
            if self.auto_scroll_var.get():
                self.terminal.see("end")

        # Ensure GUI updates happen on main thread
        self.root.after(0, _log)

    # ============================================================================
    # ALTERNATIVE: MORE ADVANCED TERMINAL WITH BETTER HORIZONTAL SCROLL
    # ============================================================================
    # If you want even better control, you can replace the CTkTextbox with this:

    def create_advanced_terminal(self, terminal_frame):
        """Create terminal with better horizontal scrolling"""

        terminal_title = ctk.CTkLabel(
            terminal_frame, text="💻 Programming Log", font=("Arial", 14, "bold")
        )
        terminal_title.pack(pady=(10, 5))

        # Create frame for terminal with scrollbars
        terminal_container = ctk.CTkFrame(terminal_frame)
        terminal_container.pack(padx=10, pady=(0, 10), fill="both", expand=True)

        # Terminal text area with no wrapping
        self.terminal = ctk.CTkTextbox(
            terminal_container,
            font=("Consolas", 24),
            wrap="none",  # No word wrapping
            state="normal",
        )
        self.terminal.pack(fill="both", expand=True, padx=5, pady=5)

        # Configure text widget for better horizontal scrolling
        # Note: This accesses the underlying tkinter Text widget
        try:
            # Get the underlying Text widget from CTkTextbox
            text_widget = self.terminal._textbox
            text_widget.configure(
                wrap="none",  # Ensure no wrapping
                font=("Consolas", 16),
                tabs=("2c",),  # Set tab stops for alignment
            )
        except AttributeError:
            # If internal structure changes, fallback to basic configuration
            pass

        return self.terminal

    def detect_nested_hubs(self):
        """Detect if we have nested USB hubs"""
        ports_info = self.get_ports_with_hub_info()
        print(ports_info)
        max_hub_level = 0
        for port_info in ports_info:
            hub_level = port_info.get("hub_level", 0)
            max_hub_level = max(max_hub_level, hub_level)

        if max_hub_level >= 2:
            self.log_message(
                f"⚠️ Nested USB hubs detected (Level {max_hub_level})", "WARNING"
            )
            self.log_message("🔧 Reducing parallel operations for stability", "INFO")

            # Auto-adjust settings for nested hubs
            self.saved_data.update(
                {
                    "max_devices_per_hub": 1,
                    "hub_programming_delay": 2.0,
                    "max_parallel_jobs": min(
                        2, self.saved_data.get("max_parallel_jobs", 4)
                    ),
                }
            )

        return max_hub_level

    def get_ports(self):
        """Returns a list of available serial ports."""
        ports = serial.tools.list_ports.comports()
        return [port.device for port in ports]

    def refresh_ports(self):
        """Refresh the list of available ports and update connected devices count."""
        self.detect_nested_hubs()
        current_ports = ["All"] + self.get_ports()
        self.port_combo.configure(values=current_ports)

        # Update connected devices count
        device_count = len(current_ports) - 1  # Subtract 1 for "All" option
        self.connected_label.configure(text=f"Connected: {device_count} devices")

        self.log_message(f"🔄 Refreshed ports: {device_count} devices found", "INFO")

    def auto_detect_boards(self):
        """Auto-detect connected boards using Arduino CLI"""
        if not self.programmer:
            messagebox.showerror("Error", "Arduino CLI not initialized!")
            return

        self.log_message("🔍 Auto-detecting connected boards...", "INFO")

        def detect_async():
            success, boards = self.programmer.list_connected_boards()

            if success and boards:
                self.log_message(f"✅ Found {len(boards)} connected boards:", "SUCCESS")

                detected_fqbns = set()
                for board in boards:
                    port = board["port"]
                    fqbn = board["fqbn"]
                    board_name = board["board_name"]

                    if fqbn:
                        detected_fqbns.add(fqbn)
                        self.log_message(f"  📋 {port}: {board_name} ({fqbn})", "INFO")
                    else:
                        self.log_message(f"  ❓ {port}: Unknown board", "WARNING")

                # Auto-select the first detected board type
                if detected_fqbns:
                    first_fqbn = list(detected_fqbns)[0]
                    for board_name, config in self.supported_boards.items():
                        if config["fqbn"] == first_fqbn:
                            self.board_combo.set(board_name)
                            self.on_board_changed(board_name)
                            break

            elif success:
                self.log_message("⚠️ No boards detected. Check connections.", "WARNING")
            else:
                self.log_message(f"❌ Detection failed: {boards}", "ERROR")

        # Run detection in background thread
        threading.Thread(target=detect_async, daemon=True).start()

    def open_custom_board_adder(self):
        """Open a small window to add a custom board with FQBN and core"""
        custom_modal = ctk.CTkToplevel(self.root)
        custom_modal.title("Add Custom Board")
        custom_modal.geometry("500x320")
        custom_modal.transient(self.root)
        custom_modal.grab_set()

        # Frame inside modal
        frame = ctk.CTkFrame(custom_modal)
        frame.pack(padx=20, pady=20, fill="both", expand=True)

        # Board name
        ctk.CTkLabel(frame, text="Board Name:").pack(anchor="w")
        board_name_entry = ctk.CTkEntry(frame, placeholder_text="e.g., Custom Uno")
        board_name_entry.pack(fill="x", pady=5)

        # FQBN
        ctk.CTkLabel(frame, text="FQBN:").pack(anchor="w")
        fqbn_entry = ctk.CTkEntry(frame, placeholder_text="e.g., arduino:avr:uno")
        fqbn_entry.pack(fill="x", pady=5)

        # Core
        ctk.CTkLabel(frame, text="Core:").pack(anchor="w")
        core_entry = ctk.CTkEntry(frame, placeholder_text="e.g., arduino:avr")
        core_entry.pack(fill="x", pady=5)

        # Save action
        def save_custom_board():
            board_name = board_name_entry.get().strip()
            fqbn = fqbn_entry.get().strip()
            core = core_entry.get().strip()

            if not board_name or not fqbn or not core:
                messagebox.showerror("Error", "Please fill in all fields.")
                return

            # Check for duplicate
            if board_name in self.supported_boards:
                messagebox.showerror(
                    "Duplicate", f"Board '{board_name}' already exists."
                )
                return

            # Add to in-memory supported boards
            new_board_data = {"fqbn": fqbn, "core": core}
            self.supported_boards[board_name] = new_board_data

            # Update dropdown
            self.board_combo.configure(values=list(self.supported_boards.keys()))

            # Save to custom_added_board.json
            try:
                custom_boards_path = BOARDS_FILE

                # Load existing custom boards if file exists
                if os.path.exists(custom_boards_path):
                    with open(custom_boards_path, "r") as f:
                        custom_data = json.load(f)
                else:
                    custom_data = {}

                # Add new board
                custom_data[board_name] = new_board_data

                # Save back to file
                with open(custom_boards_path, "w") as f:
                    json.dump(custom_data, f, indent=2)

                messagebox.showinfo(
                    "Saved", f"✅ Custom board '{board_name}' saved successfully!"
                )
                custom_modal.destroy()

            except Exception as e:
                messagebox.showerror("Error", f"Failed to save board: {str(e)}")

        # Save Button
        save_button = ctk.CTkButton(
            frame, text="➕ Add Board", command=save_custom_board
        )
        save_button.pack(pady=15)

    def on_board_changed(self, selection):
        """Handle board selection change."""
        if selection in self.supported_boards:
            if selection == "Custom":
                self.open_custom_board_adder()
            board_config = self.supported_boards[selection]
            fqbn = board_config["fqbn"]
            core = board_config["core"]
            self.board_info_label.configure(text=f"FQBN: {fqbn}   CORE: {core}")
            self.log_message(f"📋 Selected board: {selection} ({fqbn})", "INFO")

    def browse_file(self):
        if self.programming_mode.get() == "Binary":
            # Binary mode - browse for .bin/.hex files
            filetypes = [
                ("Binary files", "*.bin"),
                ("Intel HEX files", "*.hex"),
                ("All firmware", "*.bin *.hex"),
                ("All files", "*.*"),
            ]
        else:
            # Source mode - browse for .ino files
            filetypes = [("Arduino files", "*.ino *.pde"), ("All files", "*.*")]

        file_path = filedialog.askopenfilename(filetypes=filetypes)
        if file_path:
            self.file_var.set(file_path)
            file_type = self.detect_file_type(file_path)
            self.log_message(
                f"📁 Selected {file_type}: {os.path.basename(file_path)}", "SUCCESS"
            )

    def detect_file_type(self, file_path):
        """Detect the type of firmware file"""
        ext = os.path.splitext(file_path)[1].lower()

        if ext in [".ino", ".pde"]:
            return "Arduino Sketch"
        elif ext == ".hex":
            return "Intel HEX firmware"
        elif ext == ".bin":
            return "Binary firmware"
        elif ext == ".elf":
            return "ELF executable"
        else:
            return "Unknown file type"

    def test_arduino_cli(self):
        """Test Arduino CLI installation and list available cores"""
        if not self.programmer:
            self.log_message("❌ Arduino CLI programmer not initialized", "ERROR")
            return

        self.log_message("🔧 Testing Arduino CLI installation...", "INFO")

        def test_async():
            # Test basic CLI
            cli_available, version_info = self.programmer.check_arduino_cli()
            if cli_available:
                self.log_message(f"✅ Arduino CLI: {version_info}", "SUCCESS")
            else:
                self.log_message(f"❌ Arduino CLI: {version_info}", "ERROR")
                return

            # List installed cores
            try:
                result = subprocess.run(
                    ["arduino-cli", "core", "list"],
                    capture_output=True,
                    text=True,
                    timeout=30,
                    creationflags=subprocess.CREATE_NO_WINDOW  # suppress console window
                )

                if result.returncode == 0:
                    self.log_message("📦 Installed cores:", "INFO")
                    lines = result.stdout.strip().split("\n")[1:]  # Skip header
                    if lines and lines[0].strip():
                        for line in lines:
                            if line.strip():
                                self.log_message(f"  • {line}", "INFO")
                    else:
                        self.log_message("  No cores installed", "WARNING")
                else:
                    self.log_message(
                        f"❌ Failed to list cores: {result.stderr}", "ERROR"
                    )

            except Exception as e:
                self.log_message(f"❌ Core list error: {str(e)}", "ERROR")

            # Test board detection
            self.log_message("🔍 Testing board detection...", "INFO")
            success, boards = self.programmer.list_connected_boards()
            if success:
                if boards:
                    self.log_message(f"✅ Detected {len(boards)} boards", "SUCCESS")
                    for board in boards:
                        status = "✅" if board["fqbn"] else "❓"
                        self.log_message(
                            f"  {status} {board['port']}: {board['board_name']}", "INFO"
                        )
                else:
                    self.log_message("⚠️ No boards currently connected", "WARNING")
            else:
                self.log_message(f"❌ Board detection failed: {boards}", "ERROR")

        # Run tests in background
        threading.Thread(target=test_async, daemon=True).start()

    def start_programming(self):
        """Start the programming process."""
        # Validate inputs
        if not self.board_var.get():
            messagebox.showerror("Error", "Please select a target board!")
            return

        if not self.file_var.get() or not os.path.exists(self.file_var.get()):
            messagebox.showerror(
                "Error", "Please select a valid firmware file or sketch!"
            )
            return

        if not self.programmer:
            messagebox.showerror("Error", "Arduino CLI not initialized!")
            return

        # Check if Arduino CLI is working
        cli_available, version_info = self.programmer.check_arduino_cli()
        if not cli_available:
            messagebox.showerror(
                "Error",
                "Arduino CLI not available!\nPlease install Arduino CLI and restart the application.",
            )
            return

        # Reset statistics
        self.stats = {
            "total_programmed": 0,
            "successful": 0,
            "failed": 0,
            "start_time": time.time(),
            "session_results": [],
        }

        # Update UI
        self.start_button.configure(state="disabled")
        self.stop_button.configure(state="normal")
        self.stop_programming = False
        self.programming_active = True

        # Create thread pool executor
        max_workers = self.saved_data.get("max_parallel_jobs", 4)
        self.executor = ThreadPoolExecutor(max_workers=max_workers)

        # Start programming manager thread
        programming_thread = threading.Thread(
            target=self.programming_manager, daemon=True
        )
        programming_thread.start()

        # Start statistics update timer
        self.update_statistics()

        board_config = self.supported_boards[self.board_var.get()]
        self.log_message("🟢 Programming session started", "SUCCESS")
        self.log_message(
            f"📋 Board: {self.board_var.get()} ({board_config['fqbn']})", "INFO"
        )
        self.log_message(f"🔌 Port: {self.port_var.get()}", "INFO")
        self.log_message(f"📁 File: {os.path.basename(self.file_var.get())}", "INFO")
        self.log_message(f"⚡ Max parallel jobs: {max_workers}", "INFO")
        self.log_message(f"🔄 Mode: {self.mode_var.get()}", "INFO")

    def stop_programming_action(self):
        """Stop the programming process."""
        self.stop_programming = True
        self.programming_active = False

        if self.executor:
            self.executor.shutdown(wait=False)

        self.start_button.configure(state="normal")
        self.stop_button.configure(state="disabled")
        self.log_message("🔴 Programming session stopped by user", "WARNING")

    def programming_manager(self):
        """Main programming manager thread"""
        try:
            board_config = self.supported_boards[self.board_var.get()]
            fqbn = board_config["fqbn"]

            # Install core if needed
            self.log_message(
                f"🔧 Ensuring core {board_config['core']} is installed...", "INFO"
            )
            core_success, core_message = self.programmer.install_core_if_needed(fqbn)
            if not core_success:
                self.log_message(
                    f"❌ Core installation failed: {core_message}", "ERROR"
                )
                return
            else:
                self.log_message(f"✅ Core ready: {core_message}", "SUCCESS")

            if self.port_var.get() == "All":
                # Auto-detection mode
                self.auto_detect_and_program(fqbn)
            else:
                # Single port mode
                self.program_single_port(fqbn, self.port_var.get())

        except Exception as e:
            self.log_message(f"❌ Programming manager error: {str(e)}", "ERROR")
        finally:
            # Clean up
            self.programming_active = False
            if self.executor:
                self.executor.shutdown(wait=True)
            self.root.after(0, lambda: self.start_button.configure(state="normal"))
            self.root.after(0, lambda: self.stop_button.configure(state="disabled"))

    def auto_detect_and_program(self, fqbn):
        """Auto-detect new devices and program them with USB hub optimization"""
        known_ports = set()
        pending_futures = {}
        hub_load_balancing = {}
        programmed_ports = set()  # Track programmed ports in single batch mode

        self.log_message(
            "🔍 Auto-detection mode: Monitoring for new devices...", "INFO"
        )
        self.log_message(
            "📋 USB Hub Support: Devices connected through hubs will be detected",
            "INFO",
        )

        while not self.stop_programming:
            try:
                # Get current ports with hub information
                current_ports_info = self.get_ports_with_hub_info()
                current_ports = set(
                    port_info["port"] for port_info in current_ports_info
                )
                new_ports = current_ports - known_ports

                # Remove disconnected ports from tracking
                disconnected_ports = known_ports - current_ports
                for port in disconnected_ports:
                    if port in pending_futures:
                        self.log_message(f"🔌 Device disconnected: {port}", "WARNING")
                        del pending_futures[port]
                    # Update hub load tracking
                    for hub_id in list(hub_load_balancing.keys()):
                        if port in hub_load_balancing[hub_id]:
                            hub_load_balancing[hub_id].remove(port)

                # Program new devices
                for port_info in current_ports_info:
                    port = port_info["port"]
                    hub_id = port_info.get("hub_id", "direct")

                    # Skip if already programmed in single batch mode
                    if self.mode_var.get() == "Single" and port in programmed_ports:
                        continue

                    if (
                        port in new_ports
                        and not self.stop_programming
                        and port not in pending_futures
                    ):
                        # Check hub capacity
                        if self.check_hub_capacity(hub_id, hub_load_balancing):
                            self.log_message(
                                f"🔌 New device detected: {port} (Hub: {hub_id})",
                                "INFO",
                            )

                            # Track hub load
                            if hub_id not in hub_load_balancing:
                                hub_load_balancing[hub_id] = set()
                            hub_load_balancing[hub_id].add(port)

                            # Submit programming job
                            future = self.executor.submit(
                                self.program_device_with_retry, fqbn, port, hub_id
                            )
                            pending_futures[port] = future
                        else:
                            self.log_message(
                                f"⏳ Delaying {port} - USB hub {hub_id} at capacity",
                                "WARNING",
                            )

                # Check completed futures
                completed_ports = []
                for port, future in pending_futures.items():
                    if future.done():
                        try:
                            result = future.result()
                            self.result_queue.put(result)
                            completed_ports.append(port)

                            # Track programmed ports for single batch mode
                            if self.mode_var.get() == "Single":
                                programmed_ports.add(port)

                        except Exception as e:
                            self.log_message(
                                f"❌ Programming future error for {port}: {str(e)}",
                                "ERROR",
                            )
                            completed_ports.append(port)

                # Remove completed futures and update hub tracking
                for port in completed_ports:
                    if port in pending_futures:
                        del pending_futures[port]
                    # Update hub load tracking
                    for hub_id in hub_load_balancing:
                        if port in hub_load_balancing[hub_id]:
                            hub_load_balancing[hub_id].remove(port)

                known_ports = current_ports
                time.sleep(self.saved_data.get("port_scan_interval", 2))

            except Exception as e:
                self.log_message(f"❌ Auto-detection error: {str(e)}", "ERROR")
                time.sleep(1)

    def program_single_port(self, fqbn, port):
        """Program a single specific port"""
        self.log_message(f"🎯 Single port mode: Programming {port}", "INFO")

        try:
            future = self.executor.submit(
                self.program_device_with_retry, fqbn, port, "direct"
            )
            result = future.result()
            self.result_queue.put(result)
        except Exception as e:
            self.log_message(f"❌ Single port programming error: {str(e)}", "ERROR")

    def program_device_with_retry(self, fqbn, port, hub_id="direct"):
        """Program a device with retry logic and hub-aware delays"""
        firmware_path = self.file_var.get()
        board_type = self.board_var.get()
        max_retries = self.saved_data.get("max_retries", 3)

        start_time = time.time()

        # Add small delay for USB hub devices to prevent conflicts
        if hub_id != "direct":
            hub_delay = self.saved_data.get("hub_programming_delay", 0.5)
            time.sleep(hub_delay)

        for attempt in range(max_retries + 1):
            try:
                if attempt > 0:
                    self.log_message(
                        f"🔄 Retry {attempt}/{max_retries} for {port} (Hub: {hub_id})",
                        "WARNING",
                    )
                    # Longer delay between retries for hub devices
                    delay = 2 if hub_id != "direct" else 1
                    time.sleep(delay)

                # Attempt programming using Arduino CLI
                success, message = self.programmer.program_board(
                    firmware_path, fqbn, port
                )
                duration = time.time() - start_time

                result = ProgrammingResult(
                    port, success, message, duration, board_type, fqbn, self.saved_data["baud_rate"], self.log_message
                )

                if success:
                    return result
                else:
                    if attempt == max_retries:
                        # Final failure
                        return result
                    else:
                        # Log retry reason
                        self.log_message(
                            f"⚠️ {port} failed (attempt {attempt + 1}): {message}",
                            "WARNING",
                        )

            except Exception as e:
                duration = time.time() - start_time
                error_msg = f"Programming exception: {str(e)}"

                if attempt == max_retries:
                    return ProgrammingResult(
                        port, False, error_msg, duration, board_type, fqbn, self.saved_data["baud_rate"], self.log_message
                    )
                else:
                    self.log_message(
                        f"⚠️ {port} exception (attempt {attempt + 1}): {error_msg}",
                        "WARNING",
                    )

        # Should not reach here, but safety fallback
        duration = time.time() - start_time
        return ProgrammingResult(
            port, False, "Max retries exceeded", duration, board_type, fqbn, self.saved_data["baud_rate"], self.log_message
        )

    def get_ports_with_hub_info(self):
        """Get ports with USB hub information for better load balancing"""
        ports_info = []
        ports = serial.tools.list_ports.comports()

        for port in ports:
            port_info = {
                "port": port.device,
                "description": port.description,
                "hwid": port.hwid,
                "hub_id": "direct",  # Default to direct connection
                "hub_path": "direct",  # Full path for nested hubs
            }

            # Try to extract USB hub information from hardware ID
            if port.hwid:
                # Look for USB hub patterns in hardware ID
                import re

                location_match = re.search(r"LOCATION=(\d+-[\d.]+)", port.hwid)
                if location_match:
                    # Extract hub hierarchy from USB location
                    location = location_match.group(1)
                    port_info["hub_path"] = location
                    # Use the hub part as hub_id (e.g., "1-1.4" from "1-1.4.1:1.0")
                    hub_parts = location.split(":")[0].split(".")
                    if len(hub_parts) > 1:
                        port_info["hub_id"] = ".".join(hub_parts[:-1])
                        port_info["hub_level"] = len(hub_parts) - 1
                    else:
                        port_info["hub_id"] = "direct"
                        port_info["hub_level"] = 0

                # Alternative: Check for USB hub in description
                elif "USB" in port.hwid and "VID" in port.hwid:
                    # Extract VID:PID to group devices
                    vid_pid_match = re.search(
                        r"VID:PID=([0-9A-F]{4}:[0-9A-F]{4})", port.hwid
                    )
                    if vid_pid_match:
                        vid_pid = vid_pid_match.group(1)
                        # Use first part of location or VID:PID as grouping
                        port_info["hub_id"] = f"hub_{vid_pid.split(':')[0]}"

            ports_info.append(port_info)

        return ports_info

    def check_hub_capacity(self, hub_id, hub_load_balancing):
        """Check if USB hub can handle another simultaneous programming operation"""
        max_devices_per_hub = self.saved_data.get("max_devices_per_hub", 2)

        if hub_id == "direct":
            return True  # Direct connections don't have hub limitations

        current_load = len(hub_load_balancing.get(hub_id, set()))
        return current_load < max_devices_per_hub

    def start_result_processor(self):
        """Start background thread to process programming results"""

        def process_results():
            while True:
                try:
                    result = self.result_queue.get(timeout=1)
                    self.process_programming_result(result)
                except queue.Empty:
                    if not self.programming_active:
                        # Check if there are any remaining results
                        try:
                            while True:
                                result = self.result_queue.get_nowait()
                                self.process_programming_result(result)
                        except queue.Empty:
                            pass
                        time.sleep(1)
                except Exception as e:
                    self.log_message(f"❌ Result processor error: {str(e)}", "ERROR")

        result_thread = threading.Thread(target=process_results, daemon=True)
        result_thread.start()

    def process_programming_result(self, result):
        """Process a single programming result"""
        # Update statistics
        self.stats["total_programmed"] += 1
        self.stats["session_results"].append(result)

        if result.success:
            self.stats["successful"] += 1
            self.log_message(
                f"✅ {result.port} programmed successfully ({result.duration:.2f}s) [{result.fqbn}]",
                "SUCCESS",
            )
        else:
            self.stats["failed"] += 1
            self.log_message(
                f"❌ {result.port} failed: {result.message} ({result.duration:.2f}s)",
                "ERROR",
            )

        # Log detailed message if enabled
        if self.saved_data.get("log_detailed_output", False):
            self.log_message(f"  📊 Details: {result.message}", "INFO")

    def start_port_monitoring(self):
        """Start monitoring for new ports (for auto-detection)"""

        def monitor_ports():
            known_ports = set()
            while True:
                try:
                    current_ports = set(self.get_ports())
                    new_ports = current_ports - known_ports
                    removed_ports = known_ports - current_ports

                    # Only log if programming is not active to avoid spam
                    if not self.programming_active:
                        if new_ports:
                            for port in new_ports:
                                self.log_message(f"🔌 Device connected: {port}", "INFO")

                        if removed_ports:
                            for port in removed_ports:
                                self.log_message(
                                    f"🔌 Device disconnected: {port}", "INFO"
                                )

                        # Update connected devices count
                        if new_ports or removed_ports:
                            self.root.after(0, self.refresh_ports)

                    known_ports = current_ports
                    time.sleep(self.saved_data.get("port_scan_interval", 2))
                except Exception as e:
                    if self.programming_active:
                        self.log_message(
                            f"⚠️ Port monitoring error: {str(e)}", "WARNING"
                        )
                    time.sleep(5)

        monitor_thread = threading.Thread(target=monitor_ports, daemon=True)
        monitor_thread.start()

    def update_statistics(self):
        """Update the statistics display"""
        if self.stats["start_time"]:
            # Update labels
            self.total_label.configure(text=f"Total: {self.stats['total_programmed']}")
            self.success_label.configure(text=f"Success: {self.stats['successful']}")
            self.failed_label.configure(text=f"Failed: {self.stats['failed']}")

            # Calculate success rate
            if self.stats["total_programmed"] > 0:
                rate = (self.stats["successful"] / self.stats["total_programmed"]) * 100
                self.rate_label.configure(text=f"Success Rate: {rate:.1f}%")

            # Calculate runtime
            runtime = time.time() - self.stats["start_time"]
            hours = int(runtime // 3600)
            minutes = int((runtime % 3600) // 60)
            seconds = int(runtime % 60)
            self.time_label.configure(
                text=f"Runtime: {hours:02d}:{minutes:02d}:{seconds:02d}"
            )

            # Calculate throughput (devices per minute)
            if runtime > 0:
                throughput = (self.stats["total_programmed"] / runtime) * 60
                self.throughput_label.configure(
                    text=f"Throughput: {throughput:.1f}/min"
                )

        # Schedule next update
        if self.programming_active:
            self.root.after(1000, self.update_statistics)

    def log_message(self, message, level="INFO"):
        """Add a message to the terminal log with thread safety"""

        def _log():
            timestamp = time.strftime("%H:%M:%S")
            log_entry = f"[{timestamp}] {message}\n"
            if level == "MCU":
                log_entry = f" ──────➜ 💎 {message}\n"
            # Insert into terminal
            self.terminal.insert("end", log_entry)

            # Auto-scroll if enabled
            if self.auto_scroll_var.get():
                self.terminal.see("end")

        # Ensure GUI updates happen on main thread
        self.root.after(0, _log)

    def clear_terminal(self):
        """Clear the terminal log"""
        self.terminal.delete("1.0", "end")
        self.log_message("🗑️ Terminal log cleared", "INFO")

    def save_log(self):
        """Save the terminal log to a file"""
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        default_filename = f"arduino_cli_log_{timestamp}.txt"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialname=default_filename,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if file_path:
            try:
                with open(file_path, "w") as f:
                    f.write(self.terminal.get("1.0", "end"))
                self.log_message(
                    f"💾 Log saved to {os.path.basename(file_path)}", "SUCCESS"
                )
            except Exception as e:
                self.log_message(f"❌ Failed to save log: {str(e)}", "ERROR")

    def export_statistics(self):
        """Export detailed statistics to CSV file"""
        if not self.stats["session_results"]:
            messagebox.showinfo("No Data", "No programming results to export.")
            return

        timestamp = time.strftime("%Y%m%d_%H%M%S")
        default_filename = f"arduino_cli_stats_{timestamp}.csv"

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialname=default_filename,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )

        if file_path:
            try:
                import csv

                with open(file_path, "w", newline="") as f:
                    writer = csv.writer(f)

                    # Write header
                    writer.writerow(
                        [
                            "Timestamp",
                            "Port",
                            "Board Type",
                            "FQBN",
                            "Success",
                            "Duration (s)",
                            "Message",
                        ]
                    )

                    # Write results
                    for result in self.stats["session_results"]:
                        writer.writerow(
                            [
                                result.timestamp.strftime("%Y-%m-%d %H:%M:%S"),
                                result.port,
                                result.board_type,
                                result.fqbn,
                                "Success" if result.success else "Failed",
                                f"{result.duration:.2f}",
                                result.message,
                            ]
                        )

                self.log_message(
                    f"📊 Statistics exported to {os.path.basename(file_path)}",
                    "SUCCESS",
                )

            except Exception as e:
                self.log_message(f"❌ Failed to export statistics: {str(e)}", "ERROR")

    def update_mode_display(self):
        """Update the mode display label in the main UI"""
        mode_text = f"Mode: {self.mode_var.get()} | {self.programming_mode.get()}"
        self.mode_display_label.configure(text=mode_text)

    def open_settings(self):
        """Open settings modal with Arduino CLI specific options"""
        settings_modal = ctk.CTkToplevel(self.root)
        settings_modal.title("Arduino CLI Settings")
        settings_modal.geometry("650x650")
        settings_modal.transient(self.root)
        settings_modal.grab_set()

        # Settings content
        settings_frame = ctk.CTkScrollableFrame(settings_modal)
        settings_frame.pack(padx=20, pady=20, fill="both", expand=True)

        title_label = ctk.CTkLabel(
            settings_frame,
            text="⚙️ Arduino CLI Production Settings",
            font=("Arial", 16, "bold"),
        )
        title_label.pack(pady=(10, 20))

        # Programming Mode Settings
        prog_mode_label = ctk.CTkLabel(
            settings_frame, text="Programming Mode Settings", font=("Arial", 14, "bold")
        )
        prog_mode_label.pack(anchor="w", pady=(10, 5))

        # Programming Mode Selection (Binary vs Compile+Upload)
        prog_mode_frame = ctk.CTkFrame(settings_frame)
        prog_mode_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(prog_mode_frame, text="Programming Mode:").pack(
            side="left", padx=10
        )

        # Create variable for settings window
        programming_mode_var = ctk.StringVar(
            value=self.saved_data.get("programming_mode", "Binary")
        )

        prog_mode_inner_frame = ctk.CTkFrame(prog_mode_frame)
        prog_mode_inner_frame.pack(side="right", padx=10)

        binary_radio = ctk.CTkRadioButton(
            prog_mode_inner_frame,
            text="Binary Only",
            variable=programming_mode_var,
            value="Binary",
        )
        binary_radio.pack(side="left", padx=5)

        compile_radio = ctk.CTkRadioButton(
            prog_mode_inner_frame,
            text="Compile+Upload",
            variable=programming_mode_var,
            value="Compile+Upload",
        )
        compile_radio.pack(side="left", padx=5)

        # Execution Mode Selection (Continuous vs Single Batch)
        exec_mode_frame = ctk.CTkFrame(settings_frame)
        exec_mode_frame.pack(fill="x", pady=5)

        ctk.CTkLabel(exec_mode_frame, text="Execution Mode:").pack(side="left", padx=10)

        # Create variable for settings window
        execution_mode_var = ctk.StringVar(
            value=self.saved_data.get("execution_mode", "Continuous")
        )

        exec_mode_inner_frame = ctk.CTkFrame(exec_mode_frame)
        exec_mode_inner_frame.pack(side="right", padx=10)

        continuous_radio = ctk.CTkRadioButton(
            exec_mode_inner_frame,
            text="Continuous",
            variable=execution_mode_var,
            value="Continuous",
        )
        continuous_radio.pack(side="left", padx=5)

        single_radio = ctk.CTkRadioButton(
            exec_mode_inner_frame,
            text="Single Batch",
            variable=execution_mode_var,
            value="Single",
        )
        single_radio.pack(side="left", padx=5)

        # Arduino CLI settings
        cli_label = ctk.CTkLabel(
            settings_frame, text="Arduino CLI Settings", font=("Arial", 14, "bold")
        )
        cli_label.pack(anchor="w", pady=(10, 5))

        # Timeout setting
        timeout_frame = ctk.CTkFrame(settings_frame)
        timeout_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(timeout_frame, text="Programming Timeout (seconds):").pack(
            side="left", padx=10
        )
        timeout_var = ctk.StringVar(
            value=str(self.saved_data.get("programming_timeout", 60))
        )
        timeout_entry = ctk.CTkEntry(timeout_frame, textvariable=timeout_var, width=100)
        timeout_entry.pack(side="right", padx=10)

        # Baude Rate
        baud_rate  = ctk.CTkFrame(settings_frame)
        baud_rate .pack(fill="x", pady=5)
        ctk.CTkLabel(baud_rate , text="Baud Rate: ").pack(side="left", padx=10)
        baud_var = ctk.StringVar(value=str(self.saved_data.get("baud_rate", 9600)))
        baud_entry = ctk.CTkEntry(baud_rate , textvariable=baud_var, width=100)
        baud_entry.pack(side="right", padx=10)

        # Max retries
        retry_frame = ctk.CTkFrame(settings_frame)
        retry_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(retry_frame, text="Max Retry Attempts:").pack(side="left", padx=10)
        retry_var = ctk.StringVar(value=str(self.saved_data.get("max_retries", 3)))
        retry_entry = ctk.CTkEntry(retry_frame, textvariable=retry_var, width=100)
        retry_entry.pack(side="right", padx=10)

        # Verification checkbox
        verify_var = ctk.BooleanVar(
            value=self.saved_data.get("verify_programming", True)
        )
        verify_check = ctk.CTkCheckBox(
            settings_frame, text="Verify programming after upload", variable=verify_var
        )
        verify_check.pack(anchor="w", pady=5)

        # Verbose output checkbox
        verbose_var = ctk.BooleanVar(value=self.saved_data.get("verbose_output", False))
        verbose_check = ctk.CTkCheckBox(
            settings_frame, text="Verbose Arduino CLI output", variable=verbose_var
        )
        verbose_check.pack(anchor="w", pady=5)

        # Auto-install cores checkbox
        auto_cores_var = ctk.BooleanVar(
            value=self.saved_data.get("auto_install_cores", True)
        )
        auto_cores_check = ctk.CTkCheckBox(
            settings_frame,
            text="Auto-install required board cores",
            variable=auto_cores_var,
        )
        auto_cores_check.pack(anchor="w", pady=5)

        # Update index on start checkbox
        update_index_var = ctk.BooleanVar(
            value=self.saved_data.get("update_index_on_start", True)
        )
        update_index_check = ctk.CTkCheckBox(
            settings_frame,
            text="Update core index on startup",
            variable=update_index_var,
        )
        update_index_check.pack(anchor="w", pady=5)

        # Performance settings
        perf_label = ctk.CTkLabel(
            settings_frame, text="Performance Settings", font=("Arial", 14, "bold")
        )
        perf_label.pack(anchor="w", pady=(20, 5))

        # Max parallel jobs
        parallel_frame = ctk.CTkFrame(settings_frame)
        parallel_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(parallel_frame, text="Max Parallel Programming Jobs:").pack(
            side="left", padx=10
        )
        parallel_var = ctk.StringVar(
            value=str(self.saved_data.get("max_parallel_jobs", 4))
        )
        parallel_entry = ctk.CTkEntry(
            parallel_frame, textvariable=parallel_var, width=100
        )
        parallel_entry.pack(side="right", padx=10)

        # Port scan interval
        scan_frame = ctk.CTkFrame(settings_frame)
        scan_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(scan_frame, text="Port Scan Interval (seconds):").pack(
            side="left", padx=10
        )
        scan_var = ctk.StringVar(
            value=str(self.saved_data.get("port_scan_interval", 2))
        )
        scan_entry = ctk.CTkEntry(scan_frame, textvariable=scan_var, width=100)
        scan_entry.pack(side="right", padx=10)

        # USB Hub settings
        hub_label = ctk.CTkLabel(
            settings_frame, text="USB Hub Settings", font=("Arial", 14, "bold")
        )
        hub_label.pack(anchor="w", pady=(20, 5))

        # Max devices per hub
        hub_devices_frame = ctk.CTkFrame(settings_frame)
        hub_devices_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(
            hub_devices_frame, text="Max simultaneous devices per USB hub:"
        ).pack(side="left", padx=10)
        hub_devices_var = ctk.StringVar(
            value=str(self.saved_data.get("max_devices_per_hub", 2))
        )
        hub_devices_entry = ctk.CTkEntry(
            hub_devices_frame, textvariable=hub_devices_var, width=100
        )
        hub_devices_entry.pack(side="right", padx=10)

        # Hub programming delay
        hub_delay_frame = ctk.CTkFrame(settings_frame)
        hub_delay_frame.pack(fill="x", pady=5)
        ctk.CTkLabel(hub_delay_frame, text="Hub programming delay (seconds):").pack(
            side="left", padx=10
        )
        hub_delay_var = ctk.StringVar(
            value=str(self.saved_data.get("hub_programming_delay", 0.5))
        )
        hub_delay_entry = ctk.CTkEntry(
            hub_delay_frame, textvariable=hub_delay_var, width=100
        )
        hub_delay_entry.pack(side="right", padx=10)

        # Logging settings
        log_label = ctk.CTkLabel(
            settings_frame, text="Logging Settings", font=("Arial", 14, "bold")
        )
        log_label.pack(anchor="w", pady=(20, 5))

        detailed_var = ctk.BooleanVar(
            value=self.saved_data.get("log_detailed_output", False)
        )
        detailed_check = ctk.CTkCheckBox(
            settings_frame,
            text="Log detailed programming output",
            variable=detailed_var,
        )
        detailed_check.pack(anchor="w", pady=5)

        def save_settings():
            try:
                self.saved_data.update(
                    {
                        "programming_mode": programming_mode_var.get(),
                        "execution_mode": execution_mode_var.get(),
                        "programming_timeout": int(timeout_var.get()),
                        "baud_rate": int(baud_var.get()),
                        "max_retries": int(retry_var.get()),
                        "verify_programming": verify_var.get(),
                        "verbose_output": verbose_var.get(),
                        "auto_install_cores": auto_cores_var.get(),
                        "update_index_on_start": update_index_var.get(),
                        "max_parallel_jobs": int(parallel_var.get()),
                        "port_scan_interval": float(scan_var.get()),
                        "max_devices_per_hub": int(hub_devices_var.get()),
                        "hub_programming_delay": float(hub_delay_var.get()),
                        "log_detailed_output": detailed_var.get(),
                    }
                )
                self.save_settings()
                self.programming_mode.set(programming_mode_var.get())
                self.mode_var.set(execution_mode_var.get())
                self.update_mode_display()

                # Reinitialize programmer with new settings
                if self.programmer:
                    self.programmer = ArduinoCLIProgrammer(self.saved_data)
                    self.programmer.set_log_callback(self.log_message)

                self.log_message("⚙️ Settings saved successfully", "SUCCESS")
                settings_modal.destroy()
            except ValueError as e:
                messagebox.showerror(
                    "Invalid Input", "Please enter valid numeric values."
                )

        # Buttons
        button_frame = ctk.CTkFrame(settings_modal)
        button_frame.pack(fill="x", padx=20, pady=10)

        save_button = ctk.CTkButton(
            button_frame, text="💾 Save Settings", command=save_settings
        )
        save_button.pack(side="left", padx=10)

        cancel_button = ctk.CTkButton(
            button_frame, text="❌ Cancel", command=settings_modal.destroy
        )
        cancel_button.pack(side="right", padx=10)

        # Reset to defaults
        def reset_defaults():
            timeout_var.set(str(self.default_settings["programming_timeout"]))
            retry_var.set(str(self.default_settings["max_retries"]))
            verify_var.set(self.default_settings["verify_programming"])
            verbose_var.set(self.default_settings["verbose_output"])
            auto_cores_var.set(self.default_settings["auto_install_cores"])
            update_index_var.set(self.default_settings["update_index_on_start"])
            parallel_var.set(str(self.default_settings["max_parallel_jobs"]))
            scan_var.set(str(self.default_settings["port_scan_interval"]))
            hub_devices_var.set(str(self.default_settings["max_devices_per_hub"]))
            hub_delay_var.set(str(self.default_settings["hub_programming_delay"]))
            detailed_var.set(self.default_settings["log_detailed_output"])

        reset_button = ctk.CTkButton(
            button_frame, text="🔄 Reset to Defaults", command=reset_defaults
        )
        reset_button.pack(side="left", padx=10)

    def load_settings(self):
        """Load settings from file with defaults"""
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    loaded_data = json.load(f)
                # Merge with defaults
                self.saved_data = self.default_settings.copy()
                self.saved_data.update(loaded_data)
            except Exception as e:
                self.saved_data = self.default_settings.copy()
        else:
            self.saved_data = self.default_settings.copy()

        prog_mode = self.saved_data.get("programming_mode", "Binary")
        exec_mode = self.saved_data.get("execution_mode", "Continuous")
        self.log_message(
            f"📋 Loaded settings: {exec_mode} mode, {prog_mode} programming", "INFO"
        )

    def save_settings(self):
        """Save settings to file"""
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(self.saved_data, f, indent=4)
        except Exception as e:
            self.log_message(f"❌ Failed to save settings: {str(e)}", "ERROR")

    def on_closing(self):
        """Handle application closing"""
        if self.programming_active:
            if messagebox.askyesno(
                "Confirm Exit", "Programming is active. Are you sure you want to exit?"
            ):
                self.stop_programming = True
                if self.executor:
                    self.executor.shutdown(wait=False)
            else:
                return

        self.save_settings()
        self.root.destroy()


if __name__ == "__main__":
    root = ctk.CTk()
    app = MCUProgrammerApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
