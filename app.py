import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
import customtkinter as ctk
import serial.tools.list_ports
import threading
import time
import os
import json
import sys


def resource_path(relative_path):
  try:
    base_path = sys._MEIPASS
  except Exception:
    base_path = os.path.abspath(".")
  return os.path.join(base_path, relative_path)


SETTINGS_FILE = resource_path("settings.json")

# Set the appearance and theme
ctk.set_appearance_mode("Dark")  # "Dark", "Light", or "System"
ctk.set_default_color_theme("dark-blue")  # Available themes: "blue", "green", "dark-blue"


class MCUProgrammerApp:
  def __init__(self, root):
    self.root = root
    self.root.title("MCU Production Programmer")
    self.root.geometry("800x900")
    
    # Try to set icon if available
    try:
      self.root.iconbitmap(resource_path("digilog.ico"))
    except:
      pass
    
    # Initialize variables
    self.saved_data = {}
    self.programming_thread = None
    self.stop_programming = False
    
    # Supported boards with their programming parameters
    self.supported_boards = {
      "Arduino Uno": {"protocol": "avrdude", "mcu": "atmega328p", "baud": "115200"},
      "Arduino Nano": {"protocol": "avrdude", "mcu": "atmega328p", "baud": "115200"},
      "Arduino Mega": {"protocol": "avrdude", "mcu": "atmega2560", "baud": "115200"},
      "ESP32": {"protocol": "esptool", "chip": "esp32", "baud": "921600"},
      "ESP8266": {"protocol": "esptool", "chip": "esp8266", "baud": "921600"},
      "ESP32-S2": {"protocol": "esptool", "chip": "esp32s2", "baud": "921600"},
      "ESP32-S3": {"protocol": "esptool", "chip": "esp32s3", "baud": "921600"},
      "ESP32-C3": {"protocol": "esptool", "chip": "esp32c3", "baud": "921600"},
      "Raspberry Pi Pico": {"protocol": "picotool", "chip": "rp2040", "baud": "115200"},
      "STM32F103": {"protocol": "stm32flash", "chip": "stm32f103", "baud": "115200"},
      "STM32F407": {"protocol": "stm32flash", "chip": "stm32f407", "baud": "115200"},
      "Teensy 3.2": {"protocol": "teensy_loader", "chip": "mk20dx256", "baud": "115200"},
      "Teensy 4.0": {"protocol": "teensy_loader", "chip": "imxrt1062", "baud": "115200"}
    }
    
    # Statistics
    self.stats = {
      "total_programmed": 0,
      "successful": 0,
      "failed": 0,
      "start_time": None
    }
    
    # Load settings and create UI
    self.load_settings()
    self.create_widgets()
    
    # Start port monitoring
    self.start_port_monitoring()
  
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
      width=200,
      command=self.on_board_changed
    )
    self.board_combo.grid(row=0, column=1, pady=5, padx=10, sticky="w")
    
    # Board info display
    self.board_info_label = ctk.CTkLabel(config_frame, text="", text_color="gray")
    self.board_info_label.grid(row=0, column=2, pady=5, padx=10, sticky="w")
    
    # Row 1: Port selection
    port_label = ctk.CTkLabel(config_frame, text="Port:")
    port_label.grid(row=1, column=0, sticky="w", pady=5, padx=10)
    
    self.port_var = ctk.StringVar()
    self.port_combo = ctk.CTkComboBox(
      config_frame,
      variable=self.port_var,
      values=["All"] + self.get_ports(),
      width=200
    )
    self.port_combo.set("All")
    self.port_combo.grid(row=1, column=1, pady=5, padx=10, sticky="w")
    
    refresh_button = ctk.CTkButton(config_frame, text="🔄 Refresh", command=self.refresh_ports, width=80)
    refresh_button.grid(row=1, column=2, pady=5, padx=10, sticky="w")
    
    # Row 2: Firmware file selection
    file_label = ctk.CTkLabel(config_frame, text="Firmware File:")
    file_label.grid(row=2, column=0, sticky="w", pady=5, padx=10)
    
    self.file_var = ctk.StringVar()
    self.file_entry = ctk.CTkEntry(config_frame, textvariable=self.file_var, width=300)
    self.file_entry.grid(row=2, column=1, pady=5, padx=10, sticky="ew")
    
    browse_button = ctk.CTkButton(config_frame, text="📁 Browse", command=self.browse_file, width=80)
    browse_button.grid(row=2, column=2, pady=5, padx=10, sticky="w")
    
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
      height=40
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
      state="disabled"
    )
    self.stop_button.pack(side="left", padx=10, pady=10)
    
    # Settings button
    settings_button = ctk.CTkButton(
      control_frame,
      text="⚙️ Settings",
      command=self.open_settings,
      height=40
    )
    settings_button.pack(side="right", padx=10, pady=10)
    
    # Statistics frame
    stats_frame = ctk.CTkFrame(main_frame)
    stats_frame.pack(padx=10, pady=5, fill="x")
    
    stats_title = ctk.CTkLabel(stats_frame, text="📊 Programming Statistics", font=("Arial", 14, "bold"))
    stats_title.pack(pady=(10, 5))
    
    # Stats display
    stats_content_frame = ctk.CTkFrame(stats_frame)
    stats_content_frame.pack(padx=10, pady=(0, 10), fill="x")
    
    self.total_label = ctk.CTkLabel(stats_content_frame, text="Total: 0")
    self.total_label.grid(row=0, column=0, padx=10, pady=5)
    
    self.success_label = ctk.CTkLabel(stats_content_frame, text="Success: 0", text_color="green")
    self.success_label.grid(row=0, column=1, padx=10, pady=5)
    
    self.failed_label = ctk.CTkLabel(stats_content_frame, text="Failed: 0", text_color="red")
    self.failed_label.grid(row=0, column=2, padx=10, pady=5)
    
    self.rate_label = ctk.CTkLabel(stats_content_frame, text="Success Rate: 0%")
    self.rate_label.grid(row=0, column=3, padx=10, pady=5)
    
    self.time_label = ctk.CTkLabel(stats_content_frame, text="Runtime: 00:00:00")
    self.time_label.grid(row=0, column=4, padx=10, pady=5)
    
    # Configure stats grid
    for i in range(5):
      stats_content_frame.grid_columnconfigure(i, weight=1)
    
    # Terminal/Log section
    terminal_frame = ctk.CTkFrame(main_frame)
    terminal_frame.pack(padx=10, pady=5, fill="both", expand=True)
    
    terminal_title = ctk.CTkLabel(terminal_frame, text="💻 Programming Log", font=("Arial", 14, "bold"))
    terminal_title.pack(pady=(10, 5))
    
    # Terminal text area with scrollbar
    self.terminal = ctk.CTkTextbox(
      terminal_frame,
      width=750,
      height=300,
      font=("Consolas", 11)
    )
    self.terminal.pack(padx=10, pady=(0, 10), fill="both", expand=True)
    
    # Terminal controls
    terminal_controls = ctk.CTkFrame(terminal_frame)
    terminal_controls.pack(padx=10, pady=(0, 10), fill="x")
    
    clear_button = ctk.CTkButton(terminal_controls, text="🗑️ Clear Log", command=self.clear_terminal)
    clear_button.pack(side="left", padx=5)
    
    save_log_button = ctk.CTkButton(terminal_controls, text="💾 Save Log", command=self.save_log)
    save_log_button.pack(side="left", padx=5)
    
    # Auto-scroll checkbox
    self.auto_scroll_var = ctk.BooleanVar(value=True)
    auto_scroll_check = ctk.CTkCheckBox(terminal_controls, text="Auto-scroll", variable=self.auto_scroll_var)
    auto_scroll_check.pack(side="right", padx=5)
    
    # Initialize with welcome message
    self.log_message("🚀 MCU Production Programmer initialized", "INFO")
    self.log_message("Select target board, port, and firmware file to begin", "INFO")
  
  def get_ports(self):
    """Returns a list of available serial ports."""
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]
  
  def refresh_ports(self):
    """Refresh the list of available ports."""
    current_ports = ["All"] + self.get_ports()
    self.port_combo.configure(values=current_ports)
    self.log_message(f"🔄 Refreshed ports: {len(current_ports) - 1} ports found", "INFO")
  
  def on_board_changed(self, selection):
    """Handle board selection change."""
    if selection in self.supported_boards:
      board_info = self.supported_boards[selection]
      info_text = f"Protocol: {board_info['protocol']}, Baud: {board_info['baud']}"
      self.board_info_label.configure(text=info_text)
      self.log_message(f"📋 Selected board: {selection}", "INFO")
  
  def browse_file(self):
    """Open file dialog to select firmware file."""
    filetypes = [
      ("Firmware files", "*.hex *.bin *.elf *.uf2"),
      ("HEX files", "*.hex"),
      ("Binary files", "*.bin"),
      ("ELF files", "*.elf"),
      ("UF2 files", "*.uf2"),
      ("All files", "*.*")
    ]
    
    file_path = filedialog.askopenfilename(filetypes=filetypes)
    if file_path:
      self.file_var.set(file_path)
      self.log_message(f"📁 Selected firmware: {os.path.basename(file_path)}", "INFO")
  
  def start_programming(self):
    """Start the programming process."""
    # Validate inputs
    if not self.board_var.get():
      messagebox.showerror("Error", "Please select a target board!")
      return
    
    if not self.file_var.get() or not os.path.exists(self.file_var.get()):
      messagebox.showerror("Error", "Please select a valid firmware file!")
      return
    
    # Reset statistics
    self.stats = {
      "total_programmed": 0,
      "successful": 0,
      "failed": 0,
      "start_time": time.time()
    }
    
    # Update UI
    self.start_button.configure(state="disabled")
    self.stop_button.configure(state="normal")
    self.stop_programming = False
    
    # Start programming thread
    self.programming_thread = threading.Thread(target=self.programming_worker, daemon=True)
    self.programming_thread.start()
    
    # Start statistics update timer
    self.update_statistics()
    
    self.log_message("🟢 Programming session started", "SUCCESS")
    self.log_message(f"📋 Board: {self.board_var.get()}", "INFO")
    self.log_message(f"🔌 Port: {self.port_var.get()}", "INFO")
    self.log_message(f"📁 Firmware: {os.path.basename(self.file_var.get())}", "INFO")
  
  def stop_programming_action(self):
    """Stop the programming process."""
    self.stop_programming = True
    self.start_button.configure(state="normal")
    self.stop_button.configure(state="disabled")
    self.log_message("🔴 Programming session stopped by user", "WARNING")
  
  def programming_worker(self):
    """Main programming worker thread."""
    try:
      while not self.stop_programming:
        # This is where the actual programming logic will go
        # For now, we'll simulate the process
        
        if self.port_var.get() == "All":
          # Monitor for new devices
          available_ports = self.get_ports()
          for port in available_ports:
            if not self.stop_programming:
              success = self.program_device(port)
              self.update_stats(success)
              time.sleep(1)  # Brief pause between devices
        else:
          # Program specific port
          success = self.program_device(self.port_var.get())
          self.update_stats(success)
          break  # Single device mode
        
        time.sleep(0.5)  # Small delay in continuous mode
    
    except Exception as e:
      self.log_message(f"❌ Programming worker error: {str(e)}", "ERROR")
    finally:
      # Re-enable start button when done
      self.root.after(0, lambda: self.start_button.configure(state="normal"))
      self.root.after(0, lambda: self.stop_button.configure(state="disabled"))
  
  def program_device(self, port):
    """Program a single device (placeholder for actual implementation)."""
    try:
      self.log_message(f"🔄 Programming device on {port}...", "INFO")
      
      # Placeholder for actual programming logic
      # This will be replaced with real programming code
      time.sleep(2)  # Simulate programming time
      
      # Simulate success/failure (90% success rate for demo)
      import random
      success = random.random() > 0.1
      
      if success:
        self.log_message(f"✅ Device on {port} programmed successfully", "SUCCESS")
      else:
        self.log_message(f"❌ Failed to program device on {port}", "ERROR")
      
      return success
    
    except Exception as e:
      self.log_message(f"❌ Error programming {port}: {str(e)}", "ERROR")
      return False
  
  def update_stats(self, success):
    """Update programming statistics."""
    self.stats["total_programmed"] += 1
    if success:
      self.stats["successful"] += 1
    else:
      self.stats["failed"] += 1
  
  def update_statistics(self):
    """Update the statistics display."""
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
      self.time_label.configure(text=f"Runtime: {hours:02d}:{minutes:02d}:{seconds:02d}")
    
    # Schedule next update
    if not self.stop_programming and self.programming_thread and self.programming_thread.is_alive():
      self.root.after(1000, self.update_statistics)
  
  def start_port_monitoring(self):
    """Start monitoring for new ports (for auto-detection)."""
    
    def monitor_ports():
      known_ports = set()
      while True:
        try:
          current_ports = set(self.get_ports())
          new_ports = current_ports - known_ports
          
          if new_ports and self.port_var.get() == "All":
            for port in new_ports:
              self.log_message(f"🔌 New device detected on {port}", "INFO")
          
          known_ports = current_ports
          time.sleep(2)  # Check every 2 seconds
        except:
          pass
    
    monitor_thread = threading.Thread(target=monitor_ports, daemon=True)
    monitor_thread.start()
  
  def log_message(self, message, level="INFO"):
    """Add a message to the terminal log."""
    timestamp = time.strftime("%H:%M:%S")
    
    # Color coding based on level
    colors = {
      "INFO": "white",
      "SUCCESS": "green",
      "WARNING": "orange",
      "ERROR": "red"
    }
    
    log_entry = f"[{timestamp}] {message}\n"
    
    # Insert into terminal
    self.terminal.insert("end", log_entry)
    
    # Auto-scroll if enabled
    if self.auto_scroll_var.get():
      self.terminal.see("end")
  
  def clear_terminal(self):
    """Clear the terminal log."""
    self.terminal.delete("1.0", "end")
    self.log_message("🗑️ Terminal log cleared", "INFO")
  
  def save_log(self):
    """Save the terminal log to a file."""
    file_path = filedialog.asksaveasfilename(
      defaultextension=".txt",
      filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
    )
    if file_path:
      try:
        with open(file_path, "w") as f:
          f.write(self.terminal.get("1.0", "end"))
        self.log_message(f"💾 Log saved to {os.path.basename(file_path)}", "SUCCESS")
      except Exception as e:
        self.log_message(f"❌ Failed to save log: {str(e)}", "ERROR")
  
  def open_settings(self):
    """Open settings modal."""
    # Create settings modal
    settings_modal = ctk.CTkToplevel(self.root)
    settings_modal.title("Settings")
    settings_modal.geometry("500x400")
    settings_modal.transient(self.root)
    settings_modal.grab_set()
    
    # Settings content
    settings_frame = ctk.CTkFrame(settings_modal)
    settings_frame.pack(padx=20, pady=20, fill="both", expand=True)
    
    title_label = ctk.CTkLabel(settings_frame, text="⚙️ Programming Settings", font=("Arial", 16, "bold"))
    title_label.pack(pady=(10, 20))
    
    # Placeholder for settings
    info_label = ctk.CTkLabel(settings_frame,
                              text="Settings will be implemented here:\n\n• Programming timeouts\n• Retry attempts\n• Verification settings\n• Custom board configurations")
    info_label.pack(pady=20)
    
    # Close button
    close_button = ctk.CTkButton(settings_modal, text="Close", command=settings_modal.destroy)
    close_button.pack(pady=10)
  
  def load_settings(self):
    """Load settings from file."""
    if os.path.exists(SETTINGS_FILE):
      try:
        with open(SETTINGS_FILE, "r") as f:
          self.saved_data = json.load(f)
      except:
        self.saved_data = {}
    else:
      self.saved_data = {}
  
  def save_settings(self):
    """Save settings to file."""
    try:
      with open(SETTINGS_FILE, "w") as f:
        json.dump(self.saved_data, f, indent=4)
    except Exception as e:
      self.log_message(f"❌ Failed to save settings: {str(e)}", "ERROR")
  
  def on_closing(self):
    """Handle application closing."""
    if self.programming_thread and self.programming_thread.is_alive():
      self.stop_programming = True
      self.programming_thread.join(timeout=2)
    
    self.save_settings()
    self.root.destroy()


if __name__ == "__main__":
  root = ctk.CTk()
  app = MCUProgrammerApp(root)
  root.protocol("WM_DELETE_WINDOW", app.on_closing)
  root.mainloop()