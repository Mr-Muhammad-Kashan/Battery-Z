# Import the 'os' module to provide access to operating system-dependent functionality, such as file and directory operations, environment variable access, and path manipulation, ensuring cross-platform compatibility for file system interactions.
import os

# Import the 'struct' module to enable handling of binary data, allowing conversion between Python values and C-style binary data structures, critical for low-level operations like parsing hardware data or file formats.
import struct

# Import the 'sys' module to provide access to system-specific parameters and functions, such as command-line arguments, Python interpreter details, and runtime environment information, enabling dynamic program control.
import sys

# Import the 'subprocess' module to allow spawning new processes, executing external commands, and capturing their output, essential for interacting with system utilities like powercfg for battery information retrieval.
import subprocess

# Import the 'datetime' module to provide classes for manipulating dates and times, enabling precise timestamp handling for logging, scheduling, or time-based calculations in the application.
import datetime

# Import the 'time' module to provide time-related functions, such as accessing the current system time, measuring elapsed time, or introducing delays, critical for time-sensitive operations like caching.
import time

# Import the 're' module to provide regular expression matching operations, enabling pattern-based text processing for parsing system reports or extracting specific data from strings.
import re

# Import the 'random' module to provide functions for generating pseudo-random numbers, useful for creating unique identifiers or simulating random behavior in the application.
import random

# Re-import the 'struct' module (duplicate import, potentially redundant), which may indicate an oversight in code organization; included for completeness to handle binary data operations.
import struct

# Import the 'csv' module to provide functionality for reading and writing CSV files, enabling structured data storage and retrieval for logging or data export purposes.
import csv

# Import 'BeautifulSoup' from the 'bs4' module to provide robust HTML and XML parsing capabilities, enabling extraction of data from structured documents like powercfg battery reports.
from bs4 import BeautifulSoup

# Import the 'winreg' module to provide access to the Windows Registry, allowing retrieval and modification of system configuration data, such as battery-related settings or hardware information.
import winreg

# Import the 'psutil' module to provide cross-platform access to system resource information, such as CPU, memory, disk, and battery status, critical for monitoring system and battery metrics.
import psutil

# Import the 'wmi' module to provide access to Windows Management Instrumentation, enabling querying of detailed hardware and system information, such as battery and system model details.
import wmi

# Import the 'webbrowser' module to provide functionality for opening web pages in the default browser, useful for displaying reports or external resources to the user.
import webbrowser

# Import the 'ctypes' module to provide C-compatible data types and access to Windows API functions, enabling low-level system interactions for tasks like power management or UI customization.
import ctypes

# Import the 'wintypes' submodule from 'ctypes' to provide Windows-specific data types, such as handles and structures, ensuring compatibility with Windows API calls for system operations.
from ctypes import wintypes

# Import the 'tempfile' module to provide functionality for creating temporary files and directories, useful for generating temporary battery reports or caching data during runtime.
import tempfile

# Import the 'pandas' module (as 'pd') to provide high-performance data manipulation and analysis tools, enabling efficient handling of tabular data for battery statistics or log processing.
import pandas as pd

# Import the 'pyplot' submodule from 'matplotlib' (as 'plt') to provide plotting capabilities, enabling visualization of battery data through graphs and charts for user-friendly analysis.
import matplotlib.pyplot as plt

# Import 'FigureCanvasQTAgg' from 'matplotlib.backends.backend_qt5agg' as 'FigureCanvas' to provide a Qt5-compatible canvas for embedding matplotlib plots in PyQt5 GUI applications, ensuring seamless integration of visualizations.
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas

# Re-import the 'csv' module (duplicate import, potentially redundant), which may indicate an oversight in code organization; included for completeness to handle CSV file operations.
import csv

# Import the 'logging' module to provide a flexible framework for generating log messages, enabling detailed tracking of application events, errors, and debugging information.
import logging

# Import the 'json' module to provide functionality for encoding and decoding JSON data, critical for storing and retrieving structured data in cache and log files.
import json

# Import the 'platform' module to provide access to platform-specific information, such as the operating system version or architecture, ensuring compatibility and tailored behavior across systems.
import platform

# Re-import 'BeautifulSoup' from the 'bs4' module (duplicate import, potentially redundant), which may indicate an oversight; included for completeness to handle HTML/XML parsing.
from bs4 import BeautifulSoup

# Import 'QSystemTrayIcon' and 'QMenu' from 'PyQt5.QtWidgets' to provide system tray functionality, enabling the application to run in the background with a tray icon and context menu for user interaction.
from PyQt5.QtWidgets import QSystemTrayIcon, QMenu

# Import multiple Qt widgets from 'PyQt5.QtWidgets' to provide GUI components, including the main application window, labels, progress bars, buttons, layouts, and dialogs, enabling a rich user interface for the application.
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QProgressBar, QPushButton,
    QVBoxLayout, QWidget, QComboBox, QMessageBox, QSplashScreen,
    QMenuBar, QAction, QDialog, QGraphicsDropShadowEffect, QInputDialog,
    QFrame, QLineEdit, QHBoxLayout, QSizePolicy
)
# Import core Qt classes from 'PyQt5.QtCore' to provide essential functionality, such as Qt constants, timers, animations, and geometric classes, enabling advanced GUI behavior and event handling.
from PyQt5.QtCore import Qt, QTimer, QPropertyAnimation, QEasingCurve, QPoint, QRectF

# Import GUI-related classes from 'PyQt5.QtGui' to provide fonts, colors, images, icons, and painting tools, enabling customized visual styling and rendering in the application's user interface.
from PyQt5.QtGui import QFont, QColor, QPixmap, QIcon, QPainter, QBrush, QLinearGradient, QPen

# Import the 'QtCore' module from 'PyQt5' to provide core non-GUI functionality, such as signals and slots, ensuring robust event-driven programming for the application's logic.
from PyQt5 import QtCore

# Attempt to import 'win32api' and 'win32con' from the 'pywin32' package to provide access to Windows API functions and constants, enabling low-level system interactions like device enumeration; set a flag to track availability.
try:
    import win32api
    import win32con
    WIN32_AVAILABLE = True
# Handle the case where 'win32api' or 'win32con' is not available, setting a flag to False and assigning None to the modules to prevent errors in non-Windows environments or if 'pywin32' is not installed.
except ImportError:
    WIN32_AVAILABLE = False
    win32api = None
    win32con = None

# Enable high-DPI scaling for the Qt application by setting the 'AA_EnableHighDpiScaling' attribute, ensuring proper rendering and resolution independence on high-resolution displays.
QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)

# Enable high-DPI pixmap support by setting the 'AA_UseHighDpiPixmaps' attribute, ensuring that images and icons are scaled appropriately for high-resolution displays in the Qt application.
QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)

# Define the path for the application's data directory in the user's AppData folder, using 'os.path.join' to construct a platform-independent path for storing configuration, logs, and cache files.
appdata_path = os.path.join(os.getenv('APPDATA'), 'BatteryZ')

# Check if the AppData directory for BatteryZ exists; if not, create it using 'os.makedirs' to ensure a dedicated location for storing application data, preventing file access errors.
if not os.path.exists(appdata_path):
    os.makedirs(appdata_path)

# Define the path for the application's log file within the AppData directory, using 'os.path.join' to construct a platform-independent path for consistent logging.
log_file = os.path.join(appdata_path, 'battery_z.log')

# Import 'RotatingFileHandler' from 'logging.handlers' to provide a rotating log file mechanism, ensuring log files do not grow indefinitely by limiting size and maintaining backup files.
from logging.handlers import RotatingFileHandler

# Create a 'RotatingFileHandler' instance for the log file, configuring it to limit the log file size to 5 MB and keep up to 3 backup files, ensuring efficient log management.
log_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3)

# Configure the logging system using 'logging.basicConfig' to set the logging level to INFO, define a structured log format with timestamp, level, and message, and attach the rotating file handler for persistent logging.
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[log_handler]
)

# Define a function 'resource_path' to handle retrieval of resource file paths, supporting both development and bundled (e.g., PyInstaller) environments for accessing assets like icons or configuration files.
def resource_path(relative_path):
    # Check if the '_MEIPASS' attribute exists in 'sys', indicating the application is running as a bundled executable (e.g., via PyInstaller), and return the resource path relative to the temporary extraction directory.
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    # If not bundled, return the resource path relative to the current working directory, ensuring resources are accessible during development.
    return os.path.join(os.path.abspath("."), relative_path)

# Define the 'BatteryIntelligence' class to encapsulate battery monitoring and management logic, providing a structured interface for handling battery data, caching, and logging.
class BatteryIntelligence:
    # Initialize the 'BatteryIntelligence' class with optional parameters for cache and charge log file names, setting default values to 'battery_cache.json' and 'charge_log.json' for data persistence.
    def __init__(self, cache_file="battery_cache.json", charge_log_file="charge_log.json"):
        # Store the cache file name as an instance variable, specifying the file used for storing battery metadata to reduce redundant system queries.
        self.cache_file = cache_file
        # Store the charge log file name as an instance variable, specifying the file used for recording battery charge events over time.
        self.charge_log_file = charge_log_file
        # Load the cache data by calling the 'load_cache' method and store it in the 'cache' instance variable, initializing the cache for battery metadata.
        self.cache = self.load_cache()
        # Load the charge log data by calling the 'load_charge_log' method and store it in the 'charge_log' instance variable, initializing the charge event history.
        self.charge_log = self.load_charge_log()
        # Initialize an empty list for 'device_cache' to store cached device information, reducing redundant queries to hardware interfaces.
        self.device_cache = []
        # Initialize 'device_cache_time' to 0, tracking the timestamp of the last device cache update to manage cache validity.
        self.device_cache_time = 0
        # Define a dictionary of default battery values to use as fallbacks when actual data is unavailable, ensuring robust operation with sensible defaults for key battery metrics.
        self.default_values = {
            "manufacturer": "Unknown",
            "serial_number": "N/A",
            "chemistry": "Lithium-Ion",
            "design_capacity": 50000,  # mWh
            "full_charge_capacity": 45000,  # mWh
            "cycle_count": 0,
            "total_cycles": 600,  # Default for Li-Ion
            "usage_days": 30,
            "last_updated": 0,
            "battery_name": "Standard Battery"
        }
        # Define a nested dictionary 'BATTERY_DATABASE' mapping manufacturers to their respective models and expected total cycle counts, providing a reference for estimating battery lifespan based on manufacturer and model.
        self.BATTERY_DATABASE = {
            "Lenovo": {"Default": 600, "ThinkPad": 600, "Legion": 500, "IdeaPad": 450, "Yoga": 550},
            "Dell": {"Default": 500, "XPS": 600, "Inspiron": 450, "Latitude": 550, "Alienware": 500},
            "HP": {"Default": 500, "Spectre": 550, "Envy": 500, "Pavilion": 450, "Omen": 500},
            "ASUS": {"Default": 500, "ZenBook": 550, "VivoBook": 450, "ROG": 500, "TUF": 500},
            "Acer": {"Default": 500, "Nitro": 500, "Swift": 550, "Aspire": 450, "Predator": 500},
            "Microsoft": {"Default": 500, "Surface": 550},
            "Apple": {"Default": 1000, "MacBook": 1000},
            "Toshiba": {"Default": 450, "Satellite": 450},
            "Samsung": {"Default": 500, "Galaxy Book": 550},
            "Sony": {"Default": 500, "VAIO": 550},
            "Unknown": {"Default": 600}
        }

    # Define the 'load_cache' method to retrieve cached battery data from a JSON file, ensuring efficient access to previously gathered hardware information.
    def load_cache(self):
        # Check if the cache file exists using 'os.path.exists' to determine whether cached data is available for loading.
        if os.path.exists(self.cache_file):
            # Attempt to open and read the cache file, handling potential errors gracefully to ensure robust operation.
            try:
                # Open the cache file in read mode using a context manager to ensure proper file handling and closure.
                with open(self.cache_file, 'r') as f:
                    # Load the JSON data from the cache file into a Python dictionary using 'json.load' for structured data access.
                    cache = json.load(f)
                    # Check if the cache is still valid by comparing the current time with the 'last_updated' timestamp, ensuring the cache is less than 24 hours old (86400 seconds).
                    if time.time() - cache.get("last_updated", 0) < 86400:  # 24 hours
                        # Return the loaded cache if it is still valid, avoiding redundant system queries.
                        return cache
            # Handle any exceptions that occur during cache loading, logging the error to ensure traceability and returning an empty dictionary as a fallback.
            except Exception as e:
                logging.error(f"Failed to load cache: {e}")
        # Return an empty dictionary if the cache file does not exist or is invalid, ensuring the application can continue without cached data.
        return {}

    # Define the 'save_cache' method to persist the current cache data to a JSON file, ensuring battery metadata is stored for future use.
    def save_cache(self):
        # Attempt to save the cache data, handling potential errors to ensure robust file operations.
        try:
            # Update the 'last_updated' field in the cache with the current timestamp using 'time.time' to track when the cache was last modified.
            self.cache["last_updated"] = time.time()
            # Open the cache file in write mode using a context manager to ensure proper file handling and closure.
            with open(self.cache_file, 'w') as f:
                # Write the cache dictionary to the file in JSON format using 'json.dump' for persistent storage.
                json.dump(self.cache, f)
        # Handle any exceptions that occur during cache saving, logging the error to ensure traceability and prevent application crashes.
        except Exception as e:
            logging.error(f"Failed to save cache: {e}")

    # Define the 'load_charge_log' method to retrieve the charge log data from a JSON file, filtering entries to keep only those from the last 90 days.
    def load_charge_log(self):
        # Check if the charge log file exists using 'os.path.exists' to determine whether log data is available for loading.
        if os.path.exists(self.charge_log_file):
            # Attempt to open and read the charge log file, handling potential errors gracefully to ensure robust operation.
            try:
                # Open the charge log file in read mode using a context manager to ensure proper file handling and closure.
                with open(self.charge_log_file, 'r') as f:
                    # Load the JSON data from the charge log file into a Python list using 'json.load' for structured data access.
                    charge_log = json.load(f)
                    # Calculate the timestamp cutoff for retaining log entries, keeping only those from the last 90 days (90 * 24 * 3600 seconds).
                    cutoff = time.time() - (90 * 24 * 3600)  # 90 days
                    # Filter the charge log to include only entries with timestamps newer than the cutoff, ensuring the log remains relevant and manageable.
                    return [entry for entry in charge_log if entry["timestamp"] > cutoff]
            # Handle any exceptions that occur during charge log loading, logging the error to ensure traceability and returning an empty list as a fallback.
            except Exception as e:
                logging.error(f"Failed to load charge log: {e}")
        # Return an empty list if the charge log file does not exist or is invalid, ensuring the application can continue without log data.
        return []

    # Define the 'save_charge_log' method to persist the current charge log data to a JSON file, ensuring charge event history is stored for future analysis.
    def save_charge_log(self):
        # Attempt to save the charge log data, handling potential errors to ensure robust file operations.
        try:
            # Open the charge log file in write mode using a context manager to ensure proper file handling and closure.
            with open(self.charge_log_file, 'w') as f:
                # Write the charge log list to the file in JSON format using 'json.dump' for persistent storage.
                json.dump(self.charge_log, f)
        # Handle any exceptions that occur during charge log saving, logging the error to ensure traceability and prevent application crashes.
        except Exception as e:
            logging.error(f"Failed to save charge log: {e}")

    # Define the 'log_charge_event' method to record significant battery charge events, such as changes in charge status or significant percentage changes, for tracking battery usage patterns.
    def log_charge_event(self, battery):
        # Check if the battery object is None, exiting early to prevent errors from invalid battery data.
        if not battery:
            return
        # Retrieve the current battery percentage from the 'psutil' battery object, representing the current charge level as a percentage.
        current_percent = battery.percent
        # Retrieve the power plugged status from the 'psutil' battery object, indicating whether the device is currently charging.
        is_charging = battery.power_plugged
        # Get the current timestamp using 'time.time' to record when the charge event occurred.
        timestamp = time.time()
        # Check if the charge log is empty, initializing it with the first entry if no previous data exists.
        if not self.charge_log:
            # Append a new charge event to the charge log, including the timestamp, current percentage, charging status, and an initial cycle contribution of 0.
            self.charge_log.append({
                "timestamp": timestamp,
                "percent": current_percent,
                "is_charging": is_charging,
                "cycle_contribution": 0
            })
        # If the charge log is not empty, compare the current event with the last logged event to determine if logging is necessary.
        else:
            # Retrieve the last charge log entry to compare with the current event for detecting significant changes.
            last_entry = self.charge_log[-1]
            # Calculate the percentage change since the last logged event to determine if the change is significant (e.g., >= 5%).
            percent_change = current_percent - last_entry["percent"]
            # Check if the charging status has changed or the percentage change is significant (>= 5%), warranting a new log entry.
            if (last_entry["is_charging"] != is_charging or
                abs(percent_change) >= 5):  # Log only significant changes
                # Check if the battery is discharging and the percentage has increased significantly (e.g., >= 80%), indicating a possible full charge cycle.
                if percent_change > 0 and not is_charging:
                    # If the percentage change is at least 80% of a full cycle, log the event with a cycle contribution proportional to the percentage change.
                    if percent_change >= 0.8 * 100:
                        self.charge_log.append({
                            "timestamp": timestamp,
                            "percent": current_percent,
                            "is_charging": is_charging,
                            "cycle_contribution": percent_change / 100
                        })
                    # Otherwise, log the event without a cycle contribution, as the change is not significant enough to count as a cycle.
                    else:
                        self.charge_log.append({
                            "timestamp": timestamp,
                            "percent": current_percent,
                            "is_charging": is_charging,
                            "cycle_contribution": 0
                        })
                # If the battery is charging and the percentage has decreased, log the event without a cycle contribution, as discharging during charging does not contribute to cycle count.
                elif percent_change < 0 and is_charging:
                    self.charge_log.append({
                        "timestamp": timestamp,
                        "percent": current_percent,
                        "is_charging": is_charging,
                        "cycle_contribution": 0
                    })
        # Limit the charge log to the most recent 10,000 entries to prevent excessive growth, maintaining performance and storage efficiency.
        if len(self.charge_log) > 10000:
            self.charge_log = self.charge_log[-10000:]
        # Save the updated charge log to the JSON file by calling 'save_charge_log' to persist the new event data.
        self.save_charge_log()

    # Define the 'get_hardware_info' method to gather and cache detailed battery hardware information from multiple sources, such as battery data, powercfg reports, and system queries.
    def get_hardware_info(self, battery_data, powercfg_data, report_path):
        # Attempt to collect and cache hardware information, handling potential errors to ensure robust operation.
        try:
            # Estimate the number of usage days by calling 'get_usage_days' and store it in the cache, falling back to the default value if unavailable.
            self.cache["usage_days"] = self.get_usage_days() or self.default_values["usage_days"]
            # If powercfg data is available, update the cache with manufacturer, serial number, chemistry, design capacity, full charge capacity, and cycle count, using defaults where necessary.
            if powercfg_data:
                self.cache.update({
                    "manufacturer": powercfg_data.get("manufacturer", self.default_values["manufacturer"]),
                    "serial_number": powercfg_data.get("serial_number", self.default_values["serial_number"]),
                    "chemistry": powercfg_data.get("chemistry", self.default_values["chemistry"]),
                    "design_capacity": powercfg_data.get("design_capacity", self.default_values["design_capacity"]),
                    "full_charge_capacity": powercfg_data.get("full_charge_capacity", self.default_values["full_charge_capacity"]),
                    "cycle_count": powercfg_data.get("cycle_count", self.default_values["cycle_count"])
                })
            # If the manufacturer is not set in the cache, infer it from battery data using the 'infer_manufacturer' method.
            if not self.cache.get("manufacturer"):
                self.cache["manufacturer"] = self.infer_manufacturer(battery_data)
            # If the serial number is not set in the cache, generate a unique serial number using the 'generate_serial_number' method.
            if not self.cache.get("serial_number"):
                self.cache["serial_number"] = self.generate_serial_number()
            # If the chemistry is not set in the cache, infer it from battery data using the 'infer_chemistry' method.
            if not self.cache.get("chemistry"):
                self.cache["chemistry"] = self.infer_chemistry(battery_data)
            # If the design capacity is not set in the cache, estimate it from battery data using the 'estimate_capacity' method.
            if not self.cache.get("design_capacity"):
                self.cache["design_capacity"] = self.estimate_capacity(battery_data)
            # If the full charge capacity is not set in the cache, estimate it based on the design capacity using the 'estimate_full_capacity' method.
            if not self.cache.get("full_charge_capacity"):
                self.cache["full_charge_capacity"] = self.estimate_full_capacity(self.cache["design_capacity"])
            # Retrieve the cycle count, estimation status, and source by calling 'get_cycle_count', passing battery data and the powercfg report path.
            cycle_count, is_estimated, source = self.get_cycle_count(battery_data, report_path)
            # Store the retrieved cycle count in the cache for future reference.
            self.cache["cycle_count"] = cycle_count
            # Store the source of the cycle count (e.g., powercfg, WMI) in the cache for traceability.
            self.cache["cycle_source"] = source
            # Store whether the cycle count was estimated in the cache to indicate data reliability.
            self.cache["is_estimated"] = is_estimated
            # Estimate the total expected cycle count based on the manufacturer and battery data using the 'estimate_total_cycles' method.
            self.cache["total_cycles"] = self.estimate_total_cycles(self.cache["manufacturer"], battery_data)
            # Infer the battery model by calling 'infer_battery_model' with battery data, powercfg data, and the report path, and store it in the cache.
            self.cache["battery_name"] = self.infer_battery_model(battery_data, powercfg_data, report_path)
            # Save the updated cache to the JSON file by calling 'save_cache' to persist the collected hardware information.
            self.save_cache()
            # Return the updated cache dictionary containing all collected and inferred battery hardware information.
            return self.cache
        # Handle any exceptions that occur during hardware information retrieval, logging the error and falling back to default values to ensure robust operation.
        except Exception as e:
            logging.error(f"Error in get_hardware_info: {e}")
            # Update the cache with default values for all key battery metrics to ensure the application can continue functioning despite errors.
            self.cache.update({
                "manufacturer": self.default_values["manufacturer"],
                "serial_number": self.default_values["serial_number"],
                "chemistry": self.default_values["chemistry"],
                "design_capacity": self.default_values["design_capacity"],
                "full_charge_capacity": self.default_values["full_charge_capacity"],
                "cycle_count": self.default_values["cycle_count"],
                "total_cycles": self.default_values["total_cycles"],
                "usage_days": self.default_values["usage_days"],
                "cycle_source": "Default",
                "is_estimated": True,
                "battery_name": self.default_values["battery_name"]
            })
            # Save the cache with default values to the JSON file to ensure persistence despite the error.
            self.save_cache()
            # Return the cache dictionary with default values to allow the application to continue processing.
            return self.cache

    # Define the 'infer_battery_model' method to determine the battery model by querying multiple sources, such as WMI, powercfg reports, and system information, with fallbacks for robustness.
    def infer_battery_model(self, battery_data, powercfg_data, report_path):
        """Infer battery model from multiple sources."""
        # Attempt to retrieve the battery model using WMI, which provides detailed hardware information through the Win32_Battery class.
        try:
            # Create a WMI connection object to query system hardware information, enabling access to battery details.
            c = wmi.WMI()
            # Iterate through all battery instances reported by WMI to find a valid battery name or description.
            for battery in c.Win32_Battery():
                # Check if the battery's Name property is non-empty and not a generic value like "unknown" or "battery", returning it if valid.
                if battery.Name and battery.Name.strip() and battery.Name.lower() not in ["unknown", "battery"]:
                    return battery.Name.strip()
                # If the Name is invalid, check the Description property as a fallback, returning it if valid and non-generic.
                if battery.Description and battery.Description.strip() and battery.Description.lower() not in ["unknown", "battery"]:
                    return battery.Description.strip()
        # Handle any exceptions that occur during WMI queries, logging a warning to ensure traceability and continuing to the next method.
        except Exception as e:
            logging.warning(f"WMI battery model failed: {e}")

        # Attempt to retrieve the battery model from the powercfg report file, which contains detailed battery information in HTML format.
        try:
            # Check if the powercfg report file exists using 'os.path.exists' to ensure it can be read for model information.
            if os.path.exists(report_path):
                # Open the report file in read mode with UTF-8 encoding to handle special characters, using a context manager for proper file handling.
                with open(report_path, 'r', encoding='utf-8') as f:
                    # Read the entire content of the report file into a string for parsing.
                    content = f.read()
                    # Use a regular expression to search for the battery model name within XML-like tags in the powercfg report.
                    model_match = re.search(r'<Name>(.*?)</Name>', content)
                    # If a valid model name is found and it is non-empty and not generic, return it after stripping whitespace.
                    if model_match and model_match.group(1).strip() and model_match.group(1).lower() not in ["unknown", "battery"]:
                        return model_match.group(1).strip()
        # Handle any exceptions that occur during powercfg report parsing, logging a warning to ensure traceability and continuing to the next method.
        except Exception as e:
            logging.warning(f"Powercfg battery model failed: {e}")

        # Attempt to construct a battery model name from the manufacturer and system model as a fallback, using WMI to query system information.
        try:
            # Retrieve the manufacturer from the cache, falling back to 'infer_manufacturer' if not set, to ensure a valid manufacturer name.
            manufacturer = self.cache.get("manufacturer", self.infer_manufacturer(battery_data))
            # Initialize the system model as "Unknown" as a default value in case system information is unavailable.
            system_model = "Unknown"
            # Query WMI for the system model using the Win32_ComputerSystem class, which provides details about the computer's hardware.
            for system in c.Win32_ComputerSystem():
                # Retrieve the system model, stripping whitespace and falling back to "Unknown" if empty.
                system_model = system.Model.strip() or "Unknown"
            # If both manufacturer and system model are non-generic, construct and return a combined battery model name.
            if manufacturer != "Unknown" and system_model != "Unknown":
                return f"{manufacturer} {system_model} Battery"
        # Handle any exceptions that occur during system model retrieval, logging a warning to ensure traceability and proceeding to the fallback.
        except Exception as e:
            logging.warning(f"System model battery inference failed: {e}")

        # Return the default battery name from 'default_values' as a final fallback if all other methods fail to provide a valid model name.
        return self.default_values["battery_name"]

    # Define the 'find_battery_devices' method to enumerate battery devices on the system using Windows API calls, returning a list of device paths for further querying.
    def find_battery_devices(self):
        # Initialize a list with a default fallback device path "\\.\Battery0" to ensure at least one device is returned if enumeration fails.
        devices = [r"\\.\Battery0"]  # Default fallback
        # Check if the 'win32api' and 'win32con' modules are unavailable, indicating a non-Windows environment or missing 'pywin32' installation.
        if not WIN32_AVAILABLE:
            # Log a warning to indicate that device enumeration is not possible due to missing Windows API modules, relying on the default device path.
            logging.warning("win32api/win32con not available, using default battery device")
            # Return the default device list to allow the application to continue functioning.
            return devices
        # Attempt to enumerate battery devices using the Windows Setup API, handling potential errors to ensure robust operation.
        try:
            # Call 'SetupDiGetClassDevs' to retrieve a handle to a device information set for present battery devices with a device interface, using 'DIGCF_PRESENT' and 'DIGCF_DEVICEINTERFACE' flags.
            h_dev = win32api.SetupDiGetClassDevs(None, "Battery", None, win32con.DIGCF_PRESENT | win32con.DIGCF_DEVICEINTERFACE)
            # Check if a valid device information set handle was obtained, ensuring the API call succeeded.
            if h_dev != -1:
                # Initialize an empty list to store discovered battery device paths, replacing the default fallback.
                devices = []
                # Initialize an index to iterate through device information entries in the device information set.
                index = 0
                # Enter a loop to enumerate all battery devices until no more devices are found or an error occurs.
                while True:
                    # Attempt to retrieve device information for the current index, handling potential errors gracefully.
                    try:
                        # Call 'SetupDiEnumDeviceInfo' to retrieve information about the device at the current index in the device information set.
                        dev_info = win32api.SetupDiEnumDeviceInfo(h_dev, index)
                        # Call 'SetupDiGetDeviceInterfaceDetail' to retrieve the device interface path for the current device, providing access to the battery device.
                        device_path = win32api.SetupDiGetDeviceInterfaceDetail(h_dev, dev_info)
                        # Check if a valid device path was retrieved and contains "battery" in its name (case-insensitive), ensuring it is a relevant battery device.
                        if device_path and "battery" in device_path.lower():
                            # Append the device path to the devices list for further processing.
                            devices.append(device_path)
                        # Increment the index to process the next device in the set.
                        index += 1
                    # Handle any exceptions that occur during device enumeration, breaking the loop to prevent infinite iteration.
                    except:
                        break
                # Call 'SetupDiDestroyDeviceInfoList' to free the device information set handle, releasing system resources and preventing memory leaks.
                win32api.SetupDiDestroyDeviceInfoList(h_dev)
        # Note: The exception handling block is missing here, which may be intentional but could lead to unhandled errors if the try block fails; consider adding an except clause for robustness.
        # Handle any exceptions that occur during battery device enumeration, logging the error with a detailed message to ensure traceability and facilitate debugging.
        except Exception as e:
            logging.error(f"Failed to enumerate battery devices: {e}")
        # Return the list of discovered battery devices if any were found; otherwise, return a default list containing "\\.\Battery0" as a fallback to ensure the application can proceed.
        return devices if devices else [r"\\.\Battery0"]

    # Define the 'get_battery_devices' method to retrieve a list of battery devices, utilizing a caching mechanism to reduce redundant system queries and improve performance.
    def get_battery_devices(self):
        # Check if the device cache is still valid by comparing the current time with the last cache update time, ensuring the cache is less than 1 hour old (3600 seconds).
        if time.time() - self.device_cache_time < 3600:  # Cache for 1 hour
            # Return the cached list of battery devices if the cache is valid, avoiding unnecessary system calls to improve efficiency.
            return self.device_cache
        # Call the 'find_battery_devices' method to enumerate battery devices afresh if the cache is invalid or expired.
        devices = self.find_battery_devices()
        # Update the device cache with the newly retrieved list of battery devices to store the latest device information.
        self.device_cache = devices
        # Update the cache timestamp with the current time using 'time.time' to mark when the cache was last refreshed.
        self.device_cache_time = time.time()
        # Return the list of battery devices, either newly retrieved or from the cache, ensuring consistent access to device information.
        return devices

    # Define the 'get_cycle_count' method to retrieve the battery cycle count from multiple sources, prioritizing reliability and accuracy, with fallbacks to ensure robust operation.
    def get_cycle_count(self, battery_data, report_path):
        # Get the current timestamp using 'time.time' to check the validity of cached cycle count data and for logging purposes.
        current_time = time.time()
        # Check if a valid cycle count exists in the cache and if the cache is less than 1 hour old (3600 seconds) to avoid redundant queries.
        if (self.cache.get("cycle_count") is not None and
                current_time - self.cache.get("last_updated", 0) < 3600):
            # Return the cached cycle count, along with its estimation status and source, to utilize previously retrieved data and improve performance.
            return self.cache["cycle_count"], self.cache.get("is_estimated", False), self.cache.get("cycle_source", "Cached")
        # Initialize the cycle count as None, indicating no valid cycle count has been retrieved yet.
        cycle_count = None
        # Initialize the estimation flag as False, indicating that the cycle count is not estimated unless determined otherwise.
        is_estimated = False
        # Initialize the source as "Unknown", to be updated based on the method that successfully retrieves the cycle count.
        source = "Unknown"

        # Attempt Method 1: Retrieve cycle count using WMIC command-line tool, which queries the Win32_Battery class for battery information.
        try:
            # Execute the WMIC command to get the CycleCount property from Win32_Battery, capturing output as text with a 10-second timeout.
            result = subprocess.run(["wmic", "path", "Win32_Battery", "get", "CycleCount"],
                                   capture_output=True, text=True, check=True, timeout=10)
            # Extract the cycle count from the command output, taking the second line (first line is header) and stripping whitespace.
            output = result.stdout.strip().split('\n')[1].strip()
            # Verify that the output is non-empty, is a valid digit, and represents a positive cycle count to ensure data reliability.
            if output and output.isdigit() and int(output) > 0:
                # Convert the output to an integer to represent the cycle count.
                cycle_count = int(output)
                # Set the source to "WMIC" to indicate the cycle count was retrieved from the WMIC command.
                source = "WMIC"
                # Log the successful retrieval of the cycle count with its value for traceability and debugging.
                logging.info(f"Cycle count from WMIC: {cycle_count}")
                # Return the cycle count, estimation status (False), and source, indicating a successful retrieval.
                return cycle_count, is_estimated, source
        # Handle any exceptions that occur during WMIC execution, logging a warning to ensure traceability and proceeding to the next method.
        except Exception as e:
            logging.warning(f"WMIC cycle count failed: {e}")

        # Attempt Method 2: Retrieve cycle count using WMI in the root\wmi namespace, querying the BatteryCycleCount class for battery information.
        try:
            # Create a WMI connection object in the root\wmi namespace to access advanced battery information.
            c = wmi.WMI(namespace="root\\wmi")
            # Query the BatteryCycleCount class to retrieve all battery instances with cycle count data.
            batteries = c.BatteryCycleCount()
            # Iterate through each battery instance to find a valid cycle count.
            for battery in batteries:
                # Check if the CycleCount property is non-None and greater than 0 to ensure data validity.
                if battery.CycleCount is not None and battery.CycleCount > 0:
                    # Convert the cycle count to an integer for consistency.
                    cycle_count = int(battery.CycleCount)
                    # Set the source to "WMI BatteryCycleCount" to indicate the data source.
                    source = "WMI BatteryCycleCount"
                    # Log the successful retrieval of the cycle count with its value for traceability.
                    logging.info(f"Cycle count from WMI BatteryCycleCount: {cycle_count}")
                    # Return the cycle count, estimation status (False), and source, indicating a successful retrieval.
                    return cycle_count, is_estimated, source
        # Handle any exceptions that occur during WMI BatteryCycleCount query, logging a warning and proceeding to the next method.
        except Exception as e:
            logging.warning(f"WMI BatteryCycleCount failed: {e}")

        # Attempt Method 3: Retrieve cycle count from a powercfg battery report, generating a new report if necessary to ensure up-to-date data.
        try:
            # Check if the powercfg report file exists and is less than 5 minutes old (300 seconds); if not, generate a new report.
            if not os.path.exists(report_path) or time.time() - os.path.getmtime(report_path) > 300:
                # Execute the powercfg command to generate a battery report in XML format, saving it to the specified report path with a 30-second timeout.
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", report_path],
                              check=True, shell=True, timeout=30)
                # Introduce a 1-second delay to ensure the report file is fully written before attempting to read it.
                time.sleep(1)
            # Open the powercfg report file in read mode with UTF-8 encoding to handle special characters, using a context manager for proper file handling.
            with open(report_path, 'r', encoding='utf-8') as f:
                # Read the entire content of the report file into a string for parsing.
                content = f.read()
                # Use a regular expression to search for the CycleCount value within XML tags in the powercfg report.
                match = re.search(r'<CycleCount>(\d+)</CycleCount>', content)
                # Check if a valid match was found, indicating a cycle count value in the report.
                if match:
                    # Extract and convert the cycle count to an integer from the matched group.
                    cycle_count = int(match.group(1))
                    # Verify that the cycle count is positive to ensure data validity.
                    if cycle_count > 0:
                        # Set the source to "Powercfg Report" to indicate the data source.
                        source = "Powercfg Report"
                        # Log the successful retrieval of the cycle count with its value for traceability.
                        logging.info(f"Cycle count from Powercfg: {cycle_count}")
                        # Return the cycle count, estimation status (False), and source, indicating a successful retrieval.
                        return cycle_count, is_estimated, source
        # Handle any exceptions that occur during powercfg report processing, logging a warning and proceeding to the next method.
        except Exception as e:
            logging.warning(f"Powercfg cycle count failed: {e}")

        # Attempt Method 4: Retrieve cycle count using ACPI through Windows API calls via ctypes, querying battery information directly from hardware.
        try:
            # Iterate through each battery device path retrieved by 'find_battery_devices' to query cycle count information.
            for device_path in self.find_battery_devices():
                # Open a handle to the battery device using 'CreateFileW' with read/write access, ensuring proper interaction with the device.
                h_battery = ctypes.windll.kernel32.CreateFileW(
                    device_path, 0x80000000 | 0x40000000, 0, None, 3, 0, None)
                # Check if a valid handle was obtained, ensuring the device is accessible.
                if h_battery != -1:
                    # Ensure the handle is closed properly using a try-finally block to prevent resource leaks.
                    try:
                        # Define a ULONG variable to store the battery tag required for ACPI queries.
                        battery_tag = wintypes.ULONG()
                        # Call 'DeviceIoControl' with IOCTL_BATTERY_QUERY_TAG to retrieve the battery tag, which identifies the battery for subsequent queries.
                        success = ctypes.windll.kernel32.DeviceIoControl(
                            h_battery, IOCTL_BATTERY_QUERY_TAG, None, 0,
                            ctypes.byref(battery_tag), ctypes.sizeof(battery_tag), None, None)
                        # Check if the battery tag retrieval was successful.
                        if success:
                            # Create a BATTERY_QUERY_INFORMATION structure to specify the query for battery information.
                            bqi = BATTERY_QUERY_INFORMATION()
                            # Set the BatteryTag field to the retrieved tag value to identify the battery.
                            bqi.BatteryTag = battery_tag.value
                            # Set the InformationLevel to BatteryInformation to request general battery information, including cycle count.
                            bqi.InformationLevel = BATTERY_QUERY_INFORMATION_LEVEL.BatteryInformation
                            # Create a BATTERY_INFORMATION structure to store the retrieved battery information.
                            bi = BATTERY_INFORMATION()
                            # Call 'DeviceIoControl' with IOCTL_BATTERY_QUERY_INFORMATION to retrieve battery information, including cycle count.
                            success = ctypes.windll.kernel32.DeviceIoControl(
                                h_battery, IOCTL_BATTERY_QUERY_INFORMATION, ctypes.byref(bqi),
                                ctypes.sizeof(bqi), ctypes.byref(bi), ctypes.sizeof(bi), None, None)
                            # Check if the query was successful and if the cycle count is positive.
                            if success and bi.CycleCount > 0:
                                # Assign the retrieved cycle count to the cycle_count variable.
                                cycle_count = bi.CycleCount
                                # Set the source to indicate the cycle count was retrieved via ACPI for the specific device path.
                                source = f"ACPI ({device_path})"
                                # Log the successful retrieval of the cycle count with its value for traceability.
                                logging.info(f"Cycle count from ACPI: {cycle_count}")
                                # Return the cycle count, estimation status (False), and source, indicating a successful retrieval.
                                return cycle_count, is_estimated, source
                    # Ensure the battery handle is closed in the finally block to release system resources, even if an error occurs.
                    finally:
                        ctypes.windll.kernel32.CloseHandle(h_battery)
        # Handle any exceptions that occur during ACPI queries, logging a warning and proceeding to the next method.
        except Exception as e:
            logging.warning(f"ACPI cycle count failed: {e}")

        # Attempt Method 5: Estimate cycle count based on charge log data by summing cycle contributions from logged charge events.
        try:
            # Calculate the total cycle count by summing the cycle_contribution values from all charge log entries, defaulting to 0 if not present.
            total_cycles = sum(entry.get("cycle_contribution", 0) for entry in self.charge_log)
            # Check if a positive total cycle count was calculated, indicating valid charge log data.
            if total_cycles > 0:
                # Convert the total cycles to an integer for consistency.
                cycle_count = int(total_cycles)
                # Set the estimation flag to True, as this method estimates the cycle count based on logged data.
                is_estimated = True
                # Set the source to "Charge Log Analysis" to indicate the estimation method.
                source = "Charge Log Analysis"
                # Log the estimated cycle count with its value for traceability.
                logging.info(f"Estimated cycle count from charge log: {cycle_count}")
                # Return the estimated cycle count, estimation status (True), and source.
                return cycle_count, is_estimated, source
        # Handle any exceptions that occur during charge log analysis, logging a warning and proceeding to the next method.
        except Exception as e:
            logging.warning(f"Charge log cycle count failed: {e}")

        # Attempt Method 6: Estimate cycle count based on battery capacity degradation, using design and full charge capacities.
        try:
            # Retrieve the design capacity from the cache, falling back to the default value if not available.
            design_capacity = self.cache.get("design_capacity", self.default_values["design_capacity"])
            # Retrieve the full charge capacity from the cache, falling back to the default value if not available.
            full_charge_capacity = self.cache.get("full_charge_capacity", self.default_values["full_charge_capacity"])
            # Check if both capacities are positive to ensure valid calculations.
            if design_capacity > 0 and full_charge_capacity > 0:
                # Calculate the capacity retention ratio as the fraction of full charge capacity to design capacity.
                capacity_retention = full_charge_capacity / design_capacity
                # Calculate the wear level as the complement of capacity retention, representing capacity loss.
                wear_level = 1 - capacity_retention
                # Estimate the total expected cycle count for the battery based on the manufacturer and battery data.
                total_cycles = self.estimate_total_cycles(self.cache.get("manufacturer", "Unknown"), battery_data)
                # Check if there is positive wear to avoid division by zero or invalid calculations.
                if wear_level > 0:
                    # Estimate the cycle count based on wear level, assuming 20% wear per total cycle count.
                    cycle_count = int((wear_level / 0.2) * total_cycles)  # 20% wear per total cycles
                # If no wear is detected, set the cycle count to 0 as a baseline.
                else:
                    cycle_count = 0
                # Ensure the cycle count is at least 1 to avoid returning a zero or negative value.
                cycle_count = max(cycle_count, 1)
                # Set the estimation flag to True, as this method estimates the cycle count based on capacity data.
                is_estimated = True
                # Set the source to "Capacity-Based Estimation" to indicate the estimation method.
                source = "Capacity-Based Estimation"
                # Log the estimated cycle count with its value for traceability.
                logging.info(f"Estimated cycle count from capacity: {cycle_count}")
                # Return the estimated cycle count, estimation status (True), and source.
                return cycle_count, is_estimated, source
        # Handle any exceptions that occur during capacity-based estimation, logging a warning and proceeding to the fallback.
        except Exception as e:
            logging.warning(f"Capacity-based cycle count failed: {e}")

        # Fallback: Return default values if all methods fail to retrieve or estimate a cycle count, ensuring the application can continue.
        cycle_count = 0
        # Set the estimation flag to True, as the default value is an estimate rather than a measured value.
        is_estimated = True
        # Set the source to "Default" to indicate that no reliable data source was found.
        source = "Default"
        # Log a warning to indicate that all cycle count retrieval methods failed, resorting to the default value.
        logging.warning("All cycle count methods failed, using default")
        # Return the default cycle count (0), estimation status (True), and source.
        return cycle_count, is_estimated, source

    # Define the 'infer_manufacturer' method to determine the battery manufacturer using WMI queries to the Win32_ComputerSystem class, with a fallback to a default value.
    def infer_manufacturer(self, battery_data):
        # Attempt to retrieve the manufacturer information using WMI, handling potential errors to ensure robustness.
        try:
            # Create a WMI connection object to query system hardware information.
            c = wmi.WMI()
            # Iterate through Win32_ComputerSystem instances to retrieve the system manufacturer.
            for system in c.Win32_ComputerSystem():
                # Extract the first word from the Manufacturer field, stripping whitespace, to get a concise brand name.
                brand = system.Manufacturer.strip().split()[0]
                # Return the brand if it is valid and not "Unknown" or empty; otherwise, fall back to the default manufacturer.
                return brand if brand not in ["Unknown", ""] else self.default_values["manufacturer"]
        # Handle any exceptions that occur during WMI queries, returning the default manufacturer to ensure continued operation.
        except Exception:
            return self.default_values["manufacturer"]

    # Define the 'generate_serial_number' method to create a unique serial number for the battery when actual data is unavailable, using a random number for uniqueness.
    def generate_serial_number(self):
        # Generate a serial number in the format "SER" followed by a random 4-digit number (1000–9999) to ensure uniqueness.
        return f"SER{random.randint(1000, 9999)}"

    # Define the 'infer_chemistry' method to determine the battery chemistry, currently returning a default value as a placeholder for actual inference logic.
    def infer_chemistry(self, battery_data):
        # Return the default chemistry value ("Lithium-Ion") from 'default_values', as no actual inference logic is implemented in this method.
        return self.default_values["chemistry"]

    # Define the 'estimate_capacity' method to estimate the battery's design capacity based on current battery data from psutil, with a fallback to a default value.
    def estimate_capacity(self, battery_data):
        # Retrieve the current battery status using 'psutil.sensors_battery' to access real-time battery information.
        battery = psutil.sensors_battery()
        # Check if valid battery data is available and includes a percentage value to perform the estimation.
        if battery and battery.percent:
            # Estimate the design capacity by scaling a base value of 50,000 mWh by the current battery percentage, assuming a linear relationship.
            return int(50000 * (battery.percent / 100))
        # Return the default design capacity from 'default_values' if no valid battery data is available.
        return self.default_values["design_capacity"]

    # Define the 'estimate_full_capacity' method to estimate the battery's full charge capacity based on its design capacity, applying a fixed degradation factor.
    def estimate_full_capacity(self, design_capacity):
        # Estimate the full charge capacity as 90% of the design capacity, assuming typical initial degradation for a battery.
        return int(design_capacity * 0.9)

    # Define the 'estimate_total_cycles' method to estimate the total expected cycle count for a battery based on its manufacturer and system model, using the BATTERY_DATABASE.
    def estimate_total_cycles(self, manufacturer, battery_data):
        # Attempt to determine the system model and map it to a cycle count from the BATTERY_DATABASE, handling potential errors.
        try:
            # Create a WMI connection object to query system hardware information, specifically the system model.
            c = wmi.WMI()
            # Initialize the model as "Default" to use as a fallback if no specific model is matched.
            model = "Default"
            # Iterate through Win32_ComputerSystem instances to retrieve the system model.
            for system in c.Win32_ComputerSystem():
                # Extract the system model, converting it to lowercase and stripping whitespace for consistent matching.
                system_model = system.Model.strip().lower()
                # Check the BATTERY_DATABASE for the manufacturer to find a matching model series (e.g., ThinkPad, XPS).
                for series in self.BATTERY_DATABASE.get(manufacturer, {}).keys():
                    # If a non-default series name is found in the system model, use it as the model key.
                    if series.lower() in system_model and series != "Default":
                        model = series
                        break
            # Return the expected cycle count for the matched model from the BATTERY_DATABASE, falling back to the "Unknown" manufacturer if necessary.
            return self.BATTERY_DATABASE.get(manufacturer, self.BATTERY_DATABASE["Unknown"])[model]
        # Handle any exceptions that occur during WMI queries or database lookup, returning the default total cycle count.
        except Exception:
            return self.default_values["total_cycles"]

    # Define the 'get_usage_days' method to calculate the number of days the system has been in use based on the operating system installation date.
    def get_usage_days(self):
        # Attempt to retrieve the OS installation date using WMI, handling potential errors to ensure robustness.
        try:
            # Create a WMI connection object to query operating system information.
            c = wmi.WMI()
            # Iterate through Win32_OperatingSystem instances to retrieve the installation date.
            for os in c.Win32_OperatingSystem():
                # Get the InstallDate property, which represents the date and time the OS was installed.
                install_date = os.InstallDate
                # Check if a valid installation date is available to perform the calculation.
                if install_date:
                    # Parse the installation date string (format: YYYYMMDDHHMMSS) into a datetime object, ignoring microseconds.
                    install_date = datetime.datetime.strptime(install_date.split('.')[0], "%Y%m%d%H%M%S")
                    # Get the current date and time using 'datetime.datetime.now' for comparison.
                    now = datetime.datetime.now()
                    # Calculate the difference between the current time and installation date to determine usage days.
                    delta = now - install_date
                    # Return the number of days, ensuring at least 1 day to avoid zero or negative values.
                    return max(delta.days, 1)
            # Return the default usage days from 'default_values' if no valid installation date is found.
            return self.default_values["usage_days"]
        # Handle any exceptions that occur during WMI queries or date parsing, logging the error and returning the default value.
        except Exception as e:
            logging.error(f"Error in get_usage_days: {e}")
            return self.default_values["usage_days"]

    # Define the 'calculate_health' method to compute the battery health percentage based on cycle count and capacity metrics, ensuring a reliable health assessment.
    def calculate_health(self, cycle_count, total_cycles, design_capacity, full_charge_capacity):
        # Attempt to calculate the battery health, handling potential errors to ensure robustness.
        try:
            # Validate that all input parameters are numeric (integers or floats) to prevent calculation errors.
            if not all(isinstance(x, (int, float)) for x in [design_capacity, full_charge_capacity, cycle_count, total_cycles]):
                # Return 100% health if inputs are invalid to avoid errors and provide a safe default.
                return 100
            # Check if either capacity is non-positive, returning 100% health to avoid division by zero.
            if design_capacity <= 0 or full_charge_capacity <= 0:
                return 100
            # Calculate capacity-based health as the percentage of full charge capacity relative to design capacity.
            capacity_health = (full_charge_capacity / design_capacity) * 100
            # If total cycles or cycle count is invalid, return the capacity-based health capped at 100%.
            if total_cycles <= 0 or cycle_count is None:
                return min(capacity_health, 100)
            # Calculate cycle-based health, assuming a 20% health reduction per full cycle count relative to total cycles.
            cycle_health = 100 * (1 - (cycle_count / total_cycles) * 0.2)
            # Return the minimum of capacity-based health, cycle-based health, and 100% to provide a conservative health estimate.
            return min(capacity_health, cycle_health, 100)
        # Handle any exceptions that occur during health calculation, logging the error and returning 100% as a safe default.
        except Exception as e:
            logging.error(f"Error in calculate_health: {e}")
            return 100

    # Define the 'estimate_remaining_life' method to estimate the remaining battery life in years, months, and days based on capacity degradation and cycle count.
    def estimate_remaining_life(self, cycle_count, total_cycles, usage_days, design_capacity, full_charge_capacity):
        # Attempt to estimate the remaining battery life, handling potential errors to ensure robustness.
        try:
            # Validate that all input parameters are positive to prevent invalid calculations.
            if design_capacity <= 0 or full_charge_capacity <= 0 or usage_days <= 0 or total_cycles <= 0:
                # Return "N/A" if inputs are invalid to indicate that no reliable estimate can be made.
                return "N/A"
            # Calculate the capacity retention ratio as the fraction of full charge capacity to design capacity.
            capacity_retention = full_charge_capacity / design_capacity
            # If capacity retention is 100% or greater, return "Like New" to indicate the battery is in excellent condition.
            if capacity_retention >= 1.0:
                return "Like New"
            # If capacity retention is 50% or less, return "Replace Now" to indicate the battery is at or beyond end-of-life.
            if capacity_retention <= 0.5:
                return "Replace Now"
            # Calculate the degradation rate as the capacity loss per day based on historical usage.
            degradation_rate = (1 - capacity_retention) / usage_days
            # Calculate the remaining capacity to lose to reach 50% (end-of-life threshold).
            remaining_capacity_to_lose = capacity_retention - 0.5
            # If degradation rate is non-positive, assume infinite remaining life to avoid division by zero.
            if degradation_rate <= 0:
                days_remaining = float('inf')
            # Otherwise, calculate the remaining days until 50% capacity based on the degradation rate.
            else:
                days_remaining = remaining_capacity_to_lose / degradation_rate
            # Calculate the remaining cycles based on the difference between total cycles and current cycle count.
            cycles_remaining = max(total_cycles - cycle_count, 0)
            # Estimate cycles per day based on historical cycle count and usage days, defaulting to 0.1 if invalid.
            cycles_per_day = cycle_count / usage_days if cycle_count and usage_days > 0 else 0.1
            # Calculate the remaining days based on remaining cycles and cycles per day, assuming infinite if cycles per day is zero.
            cycle_days_remaining = cycles_remaining / cycles_per_day if cycles_per_day > 0 else float('inf')
            # Take the minimum of capacity-based and cycle-based remaining days for a conservative estimate.
            days_remaining = min(days_remaining, cycle_days_remaining)
            # If the remaining days is infinite, return "N/A" to indicate no reliable estimate.
            if days_remaining == float('inf'):
                return "N/A"
            # Convert the remaining days into years, months, and days for a user-friendly format.
            years = int(days_remaining // DAYS_PER_YEAR)
            months = int((days_remaining % DAYS_PER_YEAR) // DAYS_PER_MONTH)
            days = int(days_remaining % DAYS_PER_MONTH)
            # Return the formatted remaining life string, e.g., "1 years, 2 months, 3 days".
            return f"{years} years, {months} months, {days} days"
        # Handle any exceptions that occur during remaining life estimation, logging the error and returning "N/A".
        except Exception as e:
            logging.error(f"Error in estimate_remaining_life: {e}")
            return "N/A"

# Define a list of battery optimization tips to provide users with actionable advice for extending battery life and improving performance.
TIPS = [
    # Tip 1: Advise against fully discharging the battery to prevent excessive wear and stress on battery cells.
    "Avoid fully discharging your battery—it’s like overworking it!",
    # Tip 2: Recommend maintaining a cool operating environment to minimize heat-related battery degradation.
    "Keep your laptop cool; heat is a battery’s worst enemy.",
    # Tip 3: Suggest using battery saver mode during light usage to reduce power consumption and cycle count.
    "Use battery saver mode when you’re just chilling—it saves cycles.",
    # Tip 4: Advise disconnecting unused USB devices to prevent unnecessary power drain from peripherals.
    "Unplug USB stuff you’re not using; they sneakily drain power.",
    # Tip 5: Recommend keeping drivers updated to optimize power management and battery performance.
    "Update your drivers—think of it as a tune-up for your battery.",
    # Tip 6: Suggest charging within the 20–80% range to minimize stress on lithium-ion batteries and extend lifespan.
    "Charge between 20-80% for longer battery life.",
    # Tip 7: Warn against exposing the battery to extreme temperatures to prevent accelerated degradation.
    "Avoid extreme temperatures to keep your battery happy.",
    # Tip 8: Recommend dimming the screen to reduce power consumption when running on battery power.
    "Dim your screen to save some juice when on battery.",
    # Tip 9: Advise disabling Wi-Fi or Bluetooth when not in use to reduce background power consumption.
    "Disable Wi-Fi or Bluetooth when not use.",
    # Tip 10: Suggest closing unnecessary applications to minimize CPU and power usage, preserving battery life.
    "Close unnecessary apps to reduce power drain."
]

# Define a list of greeting phrases to personalize the user interface, incorporating the username for a friendly experience.
GREETING_PHRASES = [
    # Greeting 1: A casual greeting with the username, prompting the user to check battery status.
    "Hey {username}, how’s it going? Let’s check your battery!",
    # Greeting 2: A friendly greeting with the username, offering to examine battery health.
    "Hi {username}, good to see you! Time for a battery peek?",
    # Greeting 3: A polite greeting with the username, suggesting a review of battery health.
    "Hello {username}, hope you’re good! Let’s see your battery health.",
    # Greeting 4: A formal greeting with the username, encouraging a battery status check.
    "Greetings {username}, let’s check how your battery’s doing today!",
    # Greeting 5: A welcoming greeting for returning users, prompting attention to battery status.
    "Hey {username}, back again? Let’s give your battery some attention!",
]

# Define a dictionary of UI themes (Dark and Light) with color settings for various UI elements to ensure a consistent and visually appealing interface.
THEMES = {
    # Define the "Dark" theme with color settings optimized for dark backgrounds and high contrast.
    "Dark": {
        # Set the background color to a dark gray (#1e1e1e) for the main UI, providing a sleek appearance.
        "background": "#1e1e1e",
        # Set the text color to a light gray (#e0e0e0) for high readability against the dark background.
        "text": "#e0e0e0",
        # Set the header text color to a bright cyan (#00aaff) for visual distinction.
        "header": "#00aaff",
        # Set the header background color to a slightly lighter dark gray (#2a2a2a) for subtle contrast.
        "header_bg": "#2a2a2a",
        # Set the button color to cyan (#00aaff) for consistency with the header.
        "button": "#00aaff",
        # Set the button text color to white (#fff) for clear visibility on buttons.
        "button_text": "#fff",
        # Set the progress bar fill color to a bright cyan (#00ccff) for dynamic visual feedback.
        "progress_chunk": "#00ccff",
        # Set the glass effect background to a semi-transparent black (20% opacity) for a modern, frosted-glass look.
        "glass_bg": "rgba(0, 0, 0, 0.2)",
        # Set the critical alert color to red (#ff0000) to indicate urgent battery issues.
        "alert_critical": "#ff0000",
        # Set the warning alert color to yellow (#ffcc00) to indicate cautionary battery status.
        "alert_warning": "#ffcc00",
        # Set the good alert color to green (#00cc00) to indicate healthy battery status.
        "alert_good": "#00cc00",
        # Set the box background color to pure black (#000000) for UI elements like panels.
        "box_bg": "#000000",
        # Set the dialog background color to match the main background (#1e1e1e) for consistency.
        "dialog_bg": "#1e1e1e"
    },
    # Define the "Light" theme with color settings optimized for light backgrounds and high readability.
    "Light": {
        # Set the background color to a light gray (#f0f0f0) for a clean, bright appearance.
        "background": "#f0f0f0",
        # Set the text color to a dark gray (#333333) for high readability against the light background.
        "text": "#333333",
        # Set the header text color to a blue (#007acc) for visual distinction.
        "header": "#007acc",
        # Set the header background color to a medium gray (#d9d9d9) for subtle contrast.
        "header_bg": "#d9d9d9",
        # Set the button color to blue (#007acc) for consistency with the header.
        "button": "#007acc",
        # Set the button text color to white (#fff) for clear visibility on buttons.
        "button_text": "#fff",
        # Set the progress bar fill color to blue (#007acc) for dynamic visual feedback.
        "progress_chunk": "#007acc",
        # Set the glass effect background to a semi-transparent white (10% opacity) for a frosted-glass look.
        "glass_bg": "rgba(255, 255, 255, 0.1)",
        # Set the critical alert color to a darker red (#cc0000) to indicate urgent battery issues.
        "alert_critical": "#cc0000",
        # Set the warning alert color to a darker yellow (#cc9900) to indicate cautionary battery status.
        "alert_warning": "#cc9900",
        # Set the good alert color to a darker green (#009900) to indicate healthy battery status.
        "alert_good": "#009900",
        # Set the box background color to pure white (#ffffff) for UI elements like panels.
        "box_bg": "#ffffff",
        # Set the dialog background color to match the main background (#f0f0f0) for consistency.
        "dialog_bg": "#f0f0f0"
    }
}

# Define constants for battery health and cycle count thresholds to standardize calculations and comparisons across the application.
# Set the degradation threshold to 20% (0.2), representing the expected capacity loss per full cycle count for health calculations.
DEGRADATION_THRESHOLD = 0.2
# Set the charge cycle threshold to 80% (0.8), indicating the minimum charge change required to count as a full cycle.
CHARGE_CYCLE_THRESHOLD = 0.8
# Set the critical cycle percentage to 90% (0.9), indicating when the cycle count is critically high relative to total cycles.
CRITICAL_CYCLE_PERCENTAGE = 0.9
# Set the critical health percentage to 30% (0.3), indicating when the battery health is critically low and requires replacement.
CRITICAL_HEALTH_PERCENTAGE = 0.3

# Define the 'GlassFrame' class, a custom QFrame widget with a translucent, frosted-glass appearance and a drop shadow effect for a modern UI look.
class GlassFrame(QFrame):
    # Initialize the GlassFrame widget, optionally specifying a parent widget, to create a styled frame with visual effects.
    def __init__(self, parent=None):
        # Call the parent QFrame constructor to initialize the base widget, passing the parent widget for proper widget hierarchy.
        super().__init__(parent)
        # Set the WA_TranslucentBackground attribute to make the widget's background transparent, enabling the frosted-glass effect.
        self.setAttribute(Qt.WA_TranslucentBackground)
        # Create a QGraphicsDropShadowEffect to add a shadow around the frame, enhancing visual depth.
        shadow = QGraphicsDropShadowEffect()
        # Set the shadow's blur radius to 20 pixels for a soft, diffused shadow effect.
        shadow.setBlurRadius(20)
        # Set the shadow color to a semi-transparent black (100/255 alpha) for a subtle shadow appearance.
        shadow.setColor(QColor(0, 0, 0, 100))
        # Set the shadow offset to (2, 2) pixels to position the shadow slightly below and to the right of the frame.
        shadow.setOffset(2, 2)
        # Apply the shadow effect to the frame to enhance its visual appearance in the UI.
        self.setGraphicsEffect(shadow)

    # Define the 'paintEvent' method to customize the rendering of the GlassFrame widget, implementing the frosted-glass effect.
    def paintEvent(self, event):
        # Create a QPainter object to perform custom painting operations on the widget.
        painter = QPainter(self)
        # Enable antialiasing to smooth edges and curves in the painted shapes for a polished appearance.
        painter.setRenderHint(QPainter.Antialiasing)
        # Create a QRectF object representing the widget's rectangle, adjusted by 1 pixel on all sides to avoid edge artifacts.
        rect = QRectF(self.rect()).adjusted(1, 1, -1, -1)
        # Create a linear gradient from the top-left to the bottom-right of the rectangle for a subtle color transition.
        gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
        # Set the gradient color at position 0 (top-left) to a semi-transparent white (50/255 alpha) for the frosted effect.
        gradient.setColorAt(0, QColor(255, 255, 255, 50))
        # Set the gradient color at position 1 (bottom-right) to a more transparent white (20/255 alpha) for a smooth transition.
        gradient.setColorAt(1, QColor(255, 255, 255, 20))
        # Set the painter's brush to use the gradient, creating the frosted-glass background effect.
        painter.setBrush(QBrush(gradient))
        # Disable the pen to avoid drawing an outline, ensuring a clean fill.
        painter.setPen(Qt.NoPen)
        # Draw a rounded rectangle with 15-pixel corner radius to create the frosted-glass frame appearance.
        painter.drawRoundedRect(rect, 15, 15)

# Define the 'BatteryWidget' class, a custom QWidget for displaying battery health with a visual progress bar and animated glow effect.
class BatteryWidget(QWidget):
    # Initialize the BatteryWidget with a theme dictionary and optional parent widget, setting up the visual and animated components.
    def __init__(self, theme, parent=None):
        # Call the parent QWidget constructor to initialize the base widget, passing the parent for proper widget hierarchy.
        super().__init__(parent)
        # Store the provided theme dictionary to access color settings for the widget's appearance.
        self.theme = theme
        # Initialize the health percentage to 0, representing the current battery health to be displayed.
        self.health_percent = 0
        # Set the initial fill color for the battery progress bar based on the theme's progress_chunk color.
        self.fill_color = QColor(self.theme["progress_chunk"])
        # Set a fixed size of 300x120 pixels for the widget to ensure consistent layout in the UI.
        self.setFixedSize(300, 120)
        # Create a QPropertyAnimation for the glow radius to animate the widget's shadow effect.
        self.glow_animation = QPropertyAnimation(self, b"glowRadius")
        # Set the animation duration to 3000 milliseconds (3 seconds) for a smooth pulsing effect.
        self.glow_animation.setDuration(3000)
        # Set the starting glow radius to 5 pixels for the animation's initial state.
        self.glow_animation.setStartValue(5)
        # Set the ending glow radius to 15 pixels for the animation's final state.
        self.glow_animation.setEndValue(15)
        # Set the easing curve to InOutSine for a smooth, natural animation transition.
        self.glow_animation.setEasingCurve(QEasingCurve.InOutSine)
        # Set the animation to loop indefinitely (-1) for a continuous pulsing effect.
        self.glow_animation.setLoopCount(-1)
        # Start the glow animation to begin the visual effect immediately.
        self.glow_animation.start()

    # Define the 'set_health_percent' method to update the battery health percentage and adjust the fill color based on health thresholds.
    def set_health_percent(self, percent):
        # Store the provided health percentage to update the widget's display.
        self.health_percent = percent
        # Set the fill color to green (alert_good) if health is above 50%, indicating a healthy battery.
        if percent > 50:
            self.fill_color = QColor(self.theme["alert_good"])
        # Set the fill color to yellow (alert_warning) if health is between 30% and 50%, indicating a cautionary state.
        elif percent > 30:
            self.fill_color = QColor(self.theme["alert_warning"])
        # Set the fill color to red (alert_critical) if health is 30% or below, indicating a critical state.
        else:
            self.fill_color = QColor(self.theme["alert_critical"])
        # Trigger a repaint of the widget to reflect the updated health percentage and fill color.
        self.update()

    # Define the 'set_glow_radius' method to update the widget's shadow effect based on the animated glow radius.
    def set_glow_radius(self, radius):
        # Create a QGraphicsDropShadowEffect to apply a dynamic shadow to the widget.
        shadow = QGraphicsDropShadowEffect()
        # Set the shadow's blur radius to the provided value for a variable glow effect.
        shadow.setBlurRadius(radius)
        # Set the shadow color to match the current fill color, ensuring visual consistency.
        shadow.setColor(self.fill_color)
        # Set the shadow offset to (0, 0) to center the glow effect around the widget.
        shadow.setOffset(0, 0)
        # Apply the shadow effect to the widget to update its appearance during animation.
        self.setGraphicsEffect(shadow)

    # Define a Qt property 'glowRadius' to allow animation of the glow radius using the 'set_glow_radius' method.
    glowRadius = QtCore.pyqtProperty(int, fset=set_glow_radius)
    # Define a Qt property 'healthPercent' to allow updating the health percentage using the 'set_health_percent' method.
    healthPercent = QtCore.pyqtProperty(float, fset=set_health_percent)

    # Define the 'paintEvent' method to customize the rendering of the BatteryWidget, drawing the battery icon and health percentage.
    def paintEvent(self, event):
        # Create a QPainter object to perform custom painting operations on the widget.
        painter = QPainter(self)
        # Enable antialiasing to smooth edges and curves in the painted shapes for a polished appearance.
        painter.setRenderHint(QPainter.Antialiasing)

        # Draw the battery body as a rounded rectangle to represent the main battery shape.
        # Define the battery body rectangle at position (10, 10) with size 280x100 pixels.
        body_rect = QRectF(10, 10, 280, 100)
        # Set a gray pen with 4-pixel width for the battery outline to ensure visibility.
        painter.setPen(QPen(QColor(100, 100, 100), 4))
        # Set a dark gray brush for the battery body background to create a filled appearance.
        painter.setBrush(QBrush(QColor(50, 50, 50)))
        # Draw a rounded rectangle with 10-pixel corner radius for the battery body.
        painter.drawRoundedRect(body_rect, 10, 10)

        # Draw the battery cap as a small rectangle to represent the battery's positive terminal.
        # Define the cap rectangle at position (290, 45) with size 10x30 pixels.
        cap_rect = QRectF(290, 45, 10, 30)
        # Disable the pen to avoid drawing an outline for the cap.
        painter.setPen(Qt.NoPen)
        # Set a gray brush for the cap to match the battery body's aesthetic.
        painter.setBrush(QBrush(QColor(100, 100, 100)))
        # Draw the cap rectangle to complete the battery icon.
        painter.drawRect(cap_rect)

        # Draw the battery fill to represent the current health percentage as a progress bar.
        # Calculate the fill width based on the health percentage, scaling it to 270 pixels (max width within the body).
        fill_width = (self.health_percent / 100) * 270
        # Define the fill rectangle at position (15, 15) with the calculated width and 90-pixel height.
        fill_rect = QRectF(15, 15, fill_width, 90)
        # Disable the pen to avoid drawing an outline for the fill.
        painter.setPen(Qt.NoPen)
        # Set the brush to the current fill color (based on health status) for the progress bar.
        painter.setBrush(QBrush(self.fill_color))
        # Draw a rounded rectangle with 5-pixel corner radius for the fill to indicate battery health.
        painter.drawRoundedRect(fill_rect, 5, 5)

        # Draw the health percentage text in the center of the battery body for user visibility.
        # Set the pen to black for the text to ensure readability against the fill.
        painter.setPen(QColor(0, 0, 0))
        # Set the font to Arial, 24-point, bold for clear and prominent text display.
        font = QFont("Arial", 24, QFont.Bold)
        # Apply the font to the painter for rendering the percentage text.
        painter.setFont(font)
        # Format the health percentage as an integer followed by a percent sign (e.g., "75%").
        percentage_text = f"{int(self.health_percent)}%"
        # Use the body rectangle as the text bounding area for centering the text.
        text_rect = body_rect
        # Draw the percentage text centered within the body rectangle using Qt.AlignCenter.
        painter.drawText(text_rect, Qt.AlignCenter, percentage_text)

# Define the 'CycleCalibrationDialog' class, a custom QDialog for calibrating battery cycle counts with a draggable, frameless window design.
class CycleCalibrationDialog(QDialog):
    # Initialize the CycleCalibrationDialog with an optional parent widget, setting up the UI and behavior.
    def __init__(self, parent=None):
        # Call the parent QDialog constructor to initialize the base dialog, passing the parent for proper widget hierarchy.
        super().__init__(parent)
        # Set the dialog's window title to "Cycle Count Calibration" to indicate its purpose to the user.
        self.setWindowTitle("Cycle Count Calibration")
        # Set a fixed size of 400x300 pixels for the dialog to ensure consistent appearance and layout.
        self.setFixedSize(400, 300)
        # Set window flags to make the dialog frameless and always on top, creating a modern, floating window effect.
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        # Set the window icon to a custom logo located at the resource path "logo.ico" for branding.
        self.setWindowIcon(QIcon(resource_path("logo.ico")))
        # Initialize a flag to track whether the dialog is being dragged by the user, enabling custom window movement.
        self.dragging = False
        # Initialize a QPoint object to store the drag position, used for calculating mouse movement during dragging.
        self.drag_position = QPoint()
        # Create a container widget within the dialog to hold UI elements, styled with the parent's theme or defaulting to "Dark".
        self.container = QWidget(self)
        # Retrieve the current theme from the parent widget if available, defaulting to the "Dark" theme from THEMES.
        theme = THEMES[parent.current_theme if parent else "Dark"]
        self.container.setStyleSheet(f"background-color: {theme['dialog_bg']}; border-radius: 10px;")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 0)
        self.container.setGraphicsEffect(shadow)

        # Layout setup
        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        # Title
        title_label = QLabel("Cycle Count Calibration")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet(f"color: {theme['header']}; background: transparent;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Instructions
        instructions_label = QLabel(
            "We couldn’t accurately detect your battery’s cycle count. "
            "Please enter an approximate number of cycles based on your usage "
            "(e.g., one cycle = one full charge from 0% to 100%)."
        )
        instructions_label.setWordWrap(True)
        instructions_label.setStyleSheet(f"color: {theme['text']}; background: transparent; font-size: 12px;")
        layout.addWidget(instructions_label)

        # Input field
        self.cycle_input = QLineEdit()
        self.cycle_input.setPlaceholderText("Enter cycle count (e.g., 50)")
        self.cycle_input.setStyleSheet(f"padding: 5px; color: {theme['text']}; background-color: {theme['header_bg']}; border-radius: 3px;")
        layout.addWidget(self.cycle_input)

        # Buttons
        button_layout = QHBoxLayout()
        confirm_button = QPushButton("Confirm")
        confirm_button.setStyleSheet(f"background-color: {theme['button']}; color: {theme['button_text']}; padding: 8px; border-radius: 5px;")
        confirm_button.clicked.connect(self.validate_and_accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.setStyleSheet(f"background-color: #555; color: #fff; padding: 8px; border-radius: 5px;")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(confirm_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)

        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.container)

        # Fade-in animation
        self.setWindowOpacity(0)
        self.fade_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_animation.setDuration(300)
        self.fade_animation.setStartValue(0)
        self.fade_animation.setEndValue(1)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.fade_animation.start()

    def validate_and_accept(self):
        text = self.cycle_input.text().strip()
        try:
            cycle_count = int(text)
            if cycle_count > 0:
                self.accept()
            else:
                QMessageBox.warning(self, "Invalid Input", "Cycle count must be a positive integer.")
        except ValueError:
            QMessageBox.warning(self, "Invalid Input", "Please enter a valid integer.")

    def get_cycle_count(self):
        try:
            return int(self.cycle_input.text())
        except ValueError:
            return None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.drag_position = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        event.accept()

class AboutDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("About Battery-Z")
        self.setFixedSize(400, 300)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowIcon(QIcon(resource_path("logo.ico")))
        self.dragging = False
        self.drag_position = QPoint()

        # Main container with shadow
        self.container = QWidget(self)
        theme = THEMES[parent.current_theme if parent else "Dark"]
        self.container.setStyleSheet(f"background-color: {theme['dialog_bg']}; border-radius: 10px;")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 0)
        self.container.setGraphicsEffect(shadow)

        # Layout setup
        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Title
        title_label = QLabel("About Battery-Z")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet(f"color: {theme['header']}; background: transparent;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        # Content
        content_label = QLabel()
        content_label.setOpenExternalLinks(True)
        content_label.setStyleSheet(f"background: transparent; color: {theme['text']}; border: none; font-family: Arial; font-size: 12px;")
        content_label.setText(
            "<p style='text-align: center;'>"
            "This Tool is made with ❤️ by <b>Muhammad Kashan Tariq</b><br>"
            "A Software Engineering student from Air University Islamabad.<br><br>"
            f"<a href='https://www.linkedin.com/in/muhammad-kashan-tariq/' style='color: {theme['header']}; text-decoration: none;'>Contact Me</a><br>"
            f"<a href='https://github.com/Mr-Muhammad-Kashan' style='color: {theme['header']}; text-decoration: none;'>Github</a>"
            "</p>"
        )
        content_label.setWordWrap(True)
        content_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(content_label)

        # Close button
        close_button = QPushButton("Close")
        close_button.setStyleSheet(f"background-color: {theme['button']}; color: {theme['button_text']}; padding: 8px; border-radius: 5px;")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignCenter)

        # Main layout
        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.container)

        # Fade-in animation
        self.setWindowOpacity(0)
        self.fade_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_animation.setDuration(300)
        self.fade_animation.setStartValue(0)
        self.fade_animation.setEndValue(1)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.fade_animation.start()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.drag_position = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        event.accept()
        


class HistoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Battery History")
        self.setFixedSize(600, 400)
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowIcon(QIcon(resource_path("logo.ico")))
        self.dragging = False
        self.drag_position = QPoint()

        theme = THEMES[parent.current_theme if parent else "Dark"]
        self.container = QWidget(self)
        self.container.setStyleSheet(f"background-color: {theme['dialog_bg']}; border-radius: 10px;")
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 0)
        self.container.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)

        title_label = QLabel("Battery Percentage History")
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet(f"color: {theme['header']}; background: transparent;")
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)

        self.figure, self.ax = plt.subplots()
        self.canvas = FigureCanvas(self.figure)
        layout.addWidget(self.canvas)

        close_button = QPushButton("Close")
        close_button.setStyleSheet(f"background-color: {theme['button']}; color: {theme['button_text']}; padding: 8px; border-radius: 5px;")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button, alignment=Qt.AlignCenter)

        main_layout = QVBoxLayout(self)
        main_layout.addWidget(self.container)

        self.load_and_plot_data()

    def load_and_plot_data(self):
        history_file = os.path.join(appdata_path, "battery_history.csv")
        if os.path.exists(history_file):
            try:
                data = pd.read_csv(history_file, parse_dates=['timestamp'])
                if not data.empty:
                    self.ax.clear()
                    self.ax.plot(data['timestamp'], data['percentage'], marker='o')
                    self.ax.set_xlabel('Time')
                    self.ax.set_ylabel('Battery Percentage (%)')
                    self.ax.set_title('Battery Percentage Over Time')
                    self.canvas.draw()
                else:
                    self.ax.clear()
                    self.ax.text(0.5, 0.5, "No data available", horizontalalignment='center', verticalalignment='center')
                    self.canvas.draw()
            except Exception as e:
                logging.error(f"Failed to load history data: {e}")
                self.ax.clear()
                self.ax.text(0.5, 0.5, "Error loading data", horizontalalignment='center', verticalalignment='center')
                self.canvas.draw()
        else:
            self.ax.clear()
            self.ax.text(0.5, 0.5, "No data available", horizontalalignment='center', verticalalignment='center')
            self.canvas.draw()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.drag_position = event.globalPos() - self.pos()
            event.accept()

    def mouseMoveEvent(self, event):
        if self.dragging and event.buttons() == Qt.LeftButton:
            self.move(event.globalPos() - self.drag_position)
            event.accept()

    def mouseReleaseEvent(self, event):
        self.dragging = False
        event.accept()

class SplashScreen(QSplashScreen):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setStyleSheet("background-color: #000000; color: #e0e0e0;")  # Changed background to black
        
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)
        
        # Add logo above the title
        logo_label = QLabel()
        logo_pixmap = QPixmap(resource_path("logo.svg"))
        logo_label.setPixmap(logo_pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation))  # Adjust size as needed
        logo_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(logo_label)
        
        title_label = QLabel("Battery-Z")
        title_label.setFont(QFont("Arial", 24, QFont.Bold))
        title_label.setStyleSheet("color: #00FF00;")  # Changed text color to green
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        tagline_label = QLabel("Your Battery’s Best Friend")
        tagline_label.setFont(QFont("Arial", 12))
        tagline_label.setStyleSheet("color: #e0e0e0;")
        tagline_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(tagline_label)
        
        progress_bar = QProgressBar()
        progress_bar.setMaximum(100)
        progress_bar.setValue(0)
        progress_bar.setFixedWidth(300)
        progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #555;
                border-radius: 5px;
                text-align: center;
                background-color: #333;
            }
            QPropertyAnimation::chunk {
                background-color: #00ccff;
                width: 10px;
            }
        """)
        layout.addWidget(progress_bar)
        
        self.loading_label = QLabel("Initializing with Love...")
        self.loading_label.setFont(QFont("Arial", 10))
        self.loading_label.setStyleSheet("color: #e0e0e0;")
        self.loading_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.loading_label)
        
        container = QWidget()
        container.setLayout(layout)
        
        self.setPixmap(QPixmap(400, 250))
        container.setGeometry(0, 0, 400, 250)
        container.setParent(self)
        
        self.progress_value = 0
        self.progress_timer = QTimer(self)
        self.progress_timer.timeout.connect(self.update_progress)
        self.progress_timer.start(30)
        
        self.fade_animation = QPropertyAnimation(self, b"windowOpacity")
        self.fade_animation.setDuration(500)
        self.fade_animation.setStartValue(1.0)
        self.fade_animation.setEndValue(0.0)
        self.fade_animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.fade_animation.finished.connect(self.close)

    def update_progress(self):
        self.progress_value += 1
        progress_bar = self.findChild(QProgressBar)
        if progress_bar:
            progress_bar.setValue(self.progress_value)
        if self.progress_value >= 100:
            self.progress_timer.stop()
            self.loading_label.setText("Ready with Love!")
            QTimer.singleShot(500, self.start_fade_out)

    def start_fade_out(self):
        self.fade_animation.start()

class SYSTEM_POWER_STATUS(ctypes.Structure):
    _fields_ = [
        ("ACLineStatus", wintypes.BYTE),
        ("BatteryFlag", wintypes.BYTE),
        ("BatteryLifePercent", wintypes.BYTE),
        ("SystemStatusFlag", wintypes.BYTE),
        ("BatteryLifeTime", wintypes.DWORD),
        ("BatteryFullLifeTime", wintypes.DWORD),
    ]

class BatteryZ(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Battery-Z: Your Battery’s Best Friend [v1.0]")
        self.resize(1400, 900)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint)
        self.setWindowIcon(QIcon(resource_path("logo.ico")))
        self.current_theme = "Dark"
        self.appdata_path = appdata_path
        self.config_file = os.path.join(self.appdata_path, "battery_config.json")
        self.report_path = os.path.join(self.appdata_path, "battery_report.xml")
        self.charge_log_file = os.path.join(self.appdata_path, "charge_log.json")
        self.cycle_cache_file = os.path.join(self.appdata_path, "cycle_cache.json")
        self.battery_intel = BatteryIntelligence(os.path.join(self.appdata_path, "battery_cache.json"), os.path.join(self.appdata_path, "charge_log.json"))
        self.cycle_count = None
        self.is_estimated = False
        self.last_cycle_fetch_time = 0
        self.manual_cycle_count = None
        self.manual_is_estimated = False
        self.manual_cycle_source = "Manual Calibration"
        self.charge_log = []
        self.has_shown_health_alert = False
        self.current_tip_index = 0
        self.low_battery_notified = False
        self.low_health_notified = False
        self.high_temp_notified = False
        self.update_counter = 0
        self.setup_system_tray()
        self.initUI()
        self.apply_theme(self.current_theme)
        self.setup_timer()
        self.load_initial_data()

    def show_notification(self, title, message):
        self.tray_icon.showMessage(title, message, QSystemTrayIcon.Information, 5000)

    def apply_theme(self, theme_name):
        theme = THEMES[theme_name]
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {theme['background']};
                color: {theme['text']};
            }}
            QMenuBar {{
                background-color: {theme['header_bg']};
                color: {theme['text']};
            }}
            QMenuBar::item {{
                background-color: {theme['header_bg']};
                color: {theme['text']};
            }}
            QMenuBar::item:selected {{
                background-color: {theme['header']};
            }}
            QMenu {{
                background-color: {theme['header_bg']};
                color: {theme['text']};
            }}
            QMenu::item:selected {{
                background-color: {theme['header']};
            }}
        """)
        self.current_theme = theme_name
        self.update_widget_styles(theme)

    def update_widget_styles(self, theme):
        for widget in self.findChildren(QWidget):
            if isinstance(widget, QLabel):
                if widget == self.greeting_label:
                    widget.setStyleSheet(f"color: {theme['progress_chunk']}; padding: 10px; background: transparent;")
                elif hasattr(widget, 'is_header') and widget.is_header:
                    pass
                elif widget in [
                    self.laptop_label, self.battery_model_label, self.design_capacity_label,
                    self.full_charge_label, self.charging_label, self.percentage_label,
                    self.time_left_label, self.source_label, self.cycle_count_label,
                    self.data_source_label, self.usage_duration_label, self.health_status_label, self.cycle_select_label, self.temperature_label
                ]:
                    text = widget.text()
                    if ":" in text:
                        label_part, value_part = text.split(":", 1)
                        widget.setText(f"<b>{label_part}:</b>{value_part}")
                    widget.setStyleSheet(f"color: {theme['text']}; background: transparent;")
                else:
                    widget.setStyleSheet(f"color: {theme['text']}; background: transparent;")
            elif isinstance(widget, QPushButton):
                widget.setStyleSheet(f"""
                    QPushButton {{
                        background-color: {theme['button']};
                        color: {theme['button_text']};
                        padding: 8px;
                        border-radius: 5px;
                    }}
                    QPushButton:hover {{
                        background-color: {theme['progress_chunk']};
                    }}
                """)
            elif isinstance(widget, QComboBox):
                widget.setStyleSheet(f"color: {theme['text']}; background-color: {theme['header_bg']}; padding: 5px; border-radius: 3px;")
            elif widget in [self.summary_label, self.tip_label]:
                widget.setStyleSheet(f"padding: 10px; background: transparent; color: {theme['text']};")
            elif hasattr(widget, 'is_header_container') and widget.is_header_container:
                widget.setStyleSheet(f"""
                    background-color: {theme['glass_bg']};
                    border-radius: 10px;
                    padding: 5px;
                """)
            elif hasattr(widget, 'is_glass_frame') and widget.is_glass_frame:
                widget.setStyleSheet(f"background-color: rgba(30, 30, 30, 200);")
            # Ensure both battery widgets are styled consistently
            elif widget in [self.battery_widget, self.charging_widget]:
                # Re-apply theme to ensure consistency
                widget.fill_color = QColor(theme["progress_chunk"])
                widget.set_health_percent(widget.health_percent)  # Refresh to apply color based on percentage

    def create_header(self, text, theme):
        header = QLabel(text)
        header.is_header = True
        header.setFont(QFont("Arial", 12, QFont.Bold))  # Reduced font size for better fit
        header.setStyleSheet(f"color: {theme['header']}; background: transparent; padding: 8px;")
        header.setAlignment(Qt.AlignCenter)
        header.setWordWrap(True)  # Enable word wrap for headers

        container = QWidget()
        container.is_header_container = True
        container.setStyleSheet(f"""
            background-color: {theme['glass_bg']};
            border-radius: 10px;
            padding: 10px;
        """)
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(15, 5, 15, 5)  # Increased margins for better spacing
        container_layout.addWidget(header)

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(2, 2)
        container.setGraphicsEffect(shadow)

        self.header_animation = QPropertyAnimation(container, b"styleSheet")
        self.header_animation.setDuration(800)
        self.header_animation.setStartValue(f"""
            background-color: transparent;
            border-radius: 10px;
            padding: 10px;
        """)
        self.header_animation.setEndValue(f"""
            background-color: {theme['glass_bg']};
            border-radius: 10px;
            padding: 10px;
        """)
        self.header_animation.setEasingCurve(QEasingCurve.InOutQuad)
        self.header_animation.start()

        return container

    def initUI(self):
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)
        menu_bar.setStyleSheet("padding: 5px;")

        view_menu = menu_bar.addMenu("View")
        theme_menu = view_menu.addMenu("Theme")
        dark_action = QAction("Dark", self)
        light_action = QAction("Light", self)
        dark_action.triggered.connect(lambda: self.apply_theme("Dark"))
        light_action.triggered.connect(lambda: self.apply_theme("Light"))
        theme_menu.addAction(dark_action)
        theme_menu.addAction(light_action)
        refresh_action = QAction("Refresh", self)
        refresh_action.triggered.connect(self.refresh_data)
        view_menu.addAction(refresh_action)
        calibrate_action = QAction("Calibrate Cycle Count", self)
        calibrate_action.triggered.connect(self.show_cycle_calibration)
        view_menu.addAction(calibrate_action)

        coffee_action = QAction("Buy Me a Coffee 🍵", self)
        coffee_action.triggered.connect(self.open_coffee_link)
        menu_bar.addAction(coffee_action)

        about_action = QAction("About Me", self)
        about_action.triggered.connect(self.show_about_dialog)
        menu_bar.addAction(about_action)

        # Create the main widget to hold the layout
        main_widget = QWidget()

        # Create the horizontal layout for the main widget
        main_layout = QHBoxLayout()
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # Left Frame (Battery Basics and Health Check)
        left_frame = GlassFrame()
        left_frame.is_glass_frame = True
        left_layout = QVBoxLayout(left_frame)
        left_layout.setSpacing(12)  # Slightly reduced spacing for better fit
        left_layout.setContentsMargins(25, 25, 25, 25)  # Increased margins

        username = os.getlogin()
        greeting_text = random.choice(GREETING_PHRASES).format(username=username)
        self.greeting_label = QLabel(greeting_text)
        self.greeting_label.setFont(QFont("Arial", 11, QFont.Bold))  # Reduced font size
        self.greeting_label.setAlignment(Qt.AlignCenter)
        self.greeting_label.setWordWrap(True)  # Enable word wrap
        left_layout.addWidget(self.greeting_label)

        left_layout.addWidget(self.create_header("Battery Basics", THEMES[self.current_theme]))
        self.laptop_label = QLabel("Laptop Model: Detecting...")
        self.laptop_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.laptop_label)
        self.battery_model_label = QLabel("Battery Model: Detecting...")
        self.battery_model_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.battery_model_label)
        self.manufacturer_label = QLabel("Battery Manufacturer: N/A")
        self.manufacturer_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.manufacturer_label)
        self.serial_number_label = QLabel("Battery Serial Number: N/A")
        self.serial_number_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.serial_number_label)
        self.chemistry_label = QLabel("Battery Chemistry: N/A")
        self.chemistry_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.chemistry_label)
        self.design_capacity_label = QLabel("Design Capacity: N/A")
        self.design_capacity_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.design_capacity_label)
        self.full_charge_label = QLabel("Full Charge Capacity: N/A")
        self.full_charge_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.full_charge_label)
        self.charging_label = QLabel("Charging State: N/A")
        self.charging_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.charging_label)
        self.percentage_label = QLabel("Battery Percentage: N/A")
        self.percentage_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.percentage_label)
        self.time_left_label = QLabel("Estimated Time Remaining: N/A")
        self.time_left_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.time_left_label)
        self.source_label = QLabel("Cycle Source: N/A")
        self.source_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        # In the initUI method, add a new label after source_label
        self.data_source_label = QLabel("Data Source: N/A")
        self.data_source_label.setStyleSheet("font-family: Arial; font-size: 10pt; padding: 2px 0;")
        left_layout.addWidget(self.data_source_label)

        left_layout.addWidget(self.source_label)

        left_layout.addWidget(self.create_header("Battery Health Check", THEMES[self.current_theme]))
        self.cycle_count_label = QLabel("Used Cycle Count: N/A")
        self.cycle_count_label.setFont(QFont("Arial", 10))
        self.cycle_count_label.setWordWrap(True)
        left_layout.addWidget(self.cycle_count_label)
        self.usage_duration_label = QLabel("Usage Duration: N/A")
        self.usage_duration_label.setFont(QFont("Arial", 10))
        self.usage_duration_label.setWordWrap(True)
        left_layout.addWidget(self.usage_duration_label)
        self.health_status_label = QLabel("Health Status: N/A")
        self.health_status_label.setFont(QFont("Arial", 10))
        self.health_status_label.setWordWrap(True)
        left_layout.addWidget(self.health_status_label)

        self.cycle_select_label = QLabel("Select Estimated Total Cycles (if unknown):")
        self.cycle_select_label.setFont(QFont("Arial", 10))
        self.cycle_select_label.setWordWrap(True)
        self.cycle_select_label.setVisible(False)
        left_layout.addWidget(self.cycle_select_label)
        self.cycle_select_combo = QComboBox()
        self.cycle_select_combo.addItems(["300", "500", "600", "700", "900", "1000"])
        self.cycle_select_combo.setCurrentIndex(1)
        self.cycle_select_combo.setVisible(False)
        self.cycle_select_combo.currentIndexChanged.connect(self.update_data)
        left_layout.addWidget(self.cycle_select_combo)

        # Center Frame (Battery Widgets for Charging State and Health)
        center_frame = GlassFrame()
        center_frame.is_glass_frame = True
        center_frame.setMinimumSize(300, 400)
        center_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        center_layout = QVBoxLayout(center_frame)
        center_layout.setContentsMargins(10, 20, 10, 20)
        center_layout.setSpacing(30)
        center_layout.setAlignment(Qt.AlignCenter)

        # Heading for Charging State
        center_layout.addWidget(self.create_header("Charging State", THEMES[self.current_theme]))

        # New Charging State Widget
        self.charging_widget = BatteryWidget(THEMES[self.current_theme], self)
        center_layout.addWidget(self.charging_widget)

        # Heading for Battery Health
        center_layout.addWidget(self.create_header("Battery Health", THEMES[self.current_theme]))

        # Existing Battery Health Widget (moved downwards)
        self.battery_widget = BatteryWidget(THEMES[self.current_theme], self)
        center_layout.addWidget(self.battery_widget)

        # Right Frame (Summary, Tips)
        right_frame = GlassFrame()
        right_frame.is_glass_frame = True
        right_layout = QVBoxLayout(right_frame)
        right_layout.setSpacing(15)  # Increased spacing for better readability
        right_layout.setContentsMargins(20, 20, 20, 20)  # Increased margins

        right_layout.addWidget(self.create_header("Sensors", THEMES[self.current_theme]))
        self.temperature_label = QLabel("Battery Temperature: N/A °C")
        self.temperature_label.setWordWrap(True)
        self.temperature_label.setFont(QFont("Arial", 10))
        right_layout.addWidget(self.temperature_label)

        self.voltage_label = QLabel("Battery Voltage: N/A V")
        self.voltage_label.setWordWrap(True)
        self.voltage_label.setFont(QFont("Arial", 10))
        right_layout.addWidget(self.voltage_label)

        self.power_label = QLabel("Power Consumption: N/A W")
        self.power_label.setWordWrap(True)
        self.power_label.setFont(QFont("Arial", 10))
        right_layout.addWidget(self.power_label)

        right_layout.addWidget(self.create_header("Your Battery Summary", THEMES[self.current_theme]))
        self.summary_label = QLabel("Hang tight, I’m checking your battery with love...")
        self.summary_label.setWordWrap(True)
        self.summary_label.setFont(QFont("Arial", 12))  # Slightly larger font for readability
        right_layout.addWidget(self.summary_label)

        right_layout.addWidget(self.create_header("Quick Battery Tips", THEMES[self.current_theme]))
        self.tip_label = QLabel("")
        self.tip_label.setWordWrap(True)
        self.tip_label.setFont(QFont("Arial", 12))  # Slightly larger font for readability
        right_layout.addWidget(self.tip_label, stretch=1)  # Add stretch to allow tip label to expand

        # Button layout for Previous and Next buttons
        button_layout = QHBoxLayout()
        button_layout.setAlignment(Qt.AlignCenter)
        self.previous_tip_button = QPushButton("Previous Tip")
        self.previous_tip_button.setFixedWidth(150)
        self.previous_tip_button.clicked.connect(self.show_previous_tip)
        button_layout.addWidget(self.previous_tip_button)

        self.next_tip_button = QPushButton("Next Tip")
        self.next_tip_button.setFixedWidth(150)
        self.next_tip_button.clicked.connect(self.show_next_tip)
        button_layout.addWidget(self.next_tip_button)

        right_layout.addLayout(button_layout)
        history_button = QPushButton("View Battery History")
        history_button.clicked.connect(self.show_history)
        right_layout.addWidget(history_button)
        self.show_next_tip()

        # Apply animations to frames
        for frame in [left_frame, right_frame, center_frame]:
            frame.setWindowOpacity(0)
            fade_anim = QPropertyAnimation(frame, b"windowOpacity")
            fade_anim.setDuration(1000)
            fade_anim.setStartValue(0)
            fade_anim.setEndValue(1)
            fade_anim.setEasingCurve(QEasingCurve.InOutQuad)
            fade_anim.start()

            scale_anim = QPropertyAnimation(frame, b"geometry")
            scale_anim.setDuration(1000)
            rect = frame.geometry()
            scale_anim.setStartValue(QRectF(rect.x(), rect.y(), rect.width() * 0.9, rect.height() * 0.9))
            scale_anim.setEndValue(QRectF(rect))
            scale_anim.setEasingCurve(QEasingCurve.InOutQuad)
            scale_anim.start()

        self.battery_animation = QPropertyAnimation(self.battery_widget, b"healthPercent")
        self.battery_animation.setDuration(1000)
        self.battery_animation.setEasingCurve(QEasingCurve.InOutQuad)

        # Add the frames to the main layout with stretch factors for proper alignment
        main_layout.addWidget(left_frame, 1)
        main_layout.addWidget(center_frame, 0, Qt.AlignCenter)
        main_layout.addWidget(right_frame, 1)

        # Set the layout on the main widget
        main_widget.setLayout(main_layout)

        # Set the main widget as the central widget
        self.setCentralWidget(main_widget)

    def open_coffee_link(self):
        webbrowser.open("https://linktr.ee/its_Muhammad_Kashan")

    def show_about_dialog(self):
        dialog = AboutDialog(self)
        dialog.exec_()

    def show_cycle_calibration(self):
        dialog = CycleCalibrationDialog(self)
        if dialog.exec_():
            cycle_count = dialog.get_cycle_count()
            if cycle_count and cycle_count > 0:
                self.manual_cycle_count = cycle_count
                self.manual_is_estimated = True
                self.update_data()
                QMessageBox.information(self, "Calibration Saved", "Cycle count updated successfully!")

    def setup_timer(self):
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_data)
        self.timer.start(15000)  # Increased for older hardware

        self.charge_timer = QTimer(self)
        self.charge_timer.timeout.connect(self.log_charge_event)
        self.charge_timer.start(60000) # Log every minute

    def load_initial_data(self):
        self.generate_battery_report()
        self.load_cycle_cache()
        self.load_charge_log()
        self.update_data()

    def generate_battery_report(self):
        try:
            if not os.path.exists(self.report_path) or time.time() - os.path.getmtime(self.report_path) > 3600:
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", self.report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            logging.error(f"Couldn’t generate battery report: {e}")
            QMessageBox.critical(self, "Error", f"Couldn’t generate battery report: {e}")

    def save_cycle_cache(self, cycle_count, is_estimated, timestamp):
        cache_data = {
            "cycle_count": cycle_count,
            "is_estimated": is_estimated,
            "timestamp": timestamp
        }
        try:
            with open(self.cycle_cache_file, 'w') as f:
                json.dump(cache_data, f)
        except Exception as e:
            logging.error(f"Failed to save cycle cache: {e}")

    def load_cycle_cache(self):
        if os.path.exists(self.cycle_cache_file):
            try:
                with open(self.cycle_cache_file, 'r') as f:
                    cache_data = json.load(f)
                    if (time.time() - cache_data["timestamp"] < 3600 and
                            "cycle_count" in cache_data and "is_estimated" in cache_data):
                        self.cycle_count = cache_data["cycle_count"]
                        self.is_estimated = cache_data["is_estimated"]
                        self.last_cycle_fetch_time = cache_data["timestamp"]
                        logging.info("Loaded cycle count from cache")
            except Exception as e:
                logging.error(f"Failed to load cycle cache: {e}")

    def save_charge_log(self):
        try:
            with open(self.charge_log_file, 'w') as f:
                json.dump(self.charge_log, f)
        except Exception as e:
            logging.error(f"Failed to save charge log: {e}")

    def load_charge_log(self):
        if os.path.exists(self.charge_log_file):
            try:
                with open(self.charge_log_file, 'r') as f:
                    self.charge_log = json.load(f)
                    cutoff = time.time() - (90 * 24 * 3600)
                    self.charge_log = [entry for entry in self.charge_log if entry["timestamp"] > cutoff]
                    self.save_charge_log()
            except Exception as e:
                logging.error(f"Failed to load charge log: {e}")

    def log_charge_event(self):
        battery = psutil.sensors_battery()
        if not battery:
            return

        current_percent = battery.percent
        is_charging = battery.power_plugged
        timestamp = time.time()

        if self.charge_log and self.charge_log[-1]["is_charging"] != is_charging:
            last_entry = self.charge_log[-1]
            percent_change = current_percent - last_entry["percent"]
            if percent_change > 0 and not is_charging:
                if percent_change >= CHARGE_CYCLE_THRESHOLD * 100:
                    self.charge_log.append({
                        "timestamp": timestamp,
                        "percent": current_percent,
                        "is_charging": is_charging,
                        "cycle_contribution": percent_change / 100
                    })
            elif percent_change < 0 and is_charging:
                self.charge_log.append({
                    "timestamp": timestamp,
                    "percent": current_percent,
                    "is_charging": is_charging,
                    "cycle_contribution": 0
                })
        else:
            self.charge_log.append({
                "timestamp": timestamp,
                "percent": current_percent,
                "is_charging": is_charging,
                "cycle_contribution": 0
            })

        if len(self.charge_log) > 10000:
            self.charge_log = self.charge_log[-10000:]
        self.save_charge_log()

    def get_cycle_count_from_charge_log(self):
        total_cycles = 0
        for entry in self.charge_log:
            total_cycles += entry.get("cycle_contribution", 0)
        return int(total_cycles)


    def get_battery_info(self):
        battery = psutil.sensors_battery()
        c = wmi.WMI()
        laptop_model = "Unknown"
        battery_name = "Unknown"
        powercfg_data = {
            "design_capacity": "N/A",
            "full_charge_capacity": "N/A",
            "manufacturer": "N/A",
            "serial_number": "N/A",
            "chemistry": "N/A"
        }

        # Method 1: Powercfg Report
        try:
            if not os.path.exists(self.report_path) or time.time() - os.path.getmtime(self.report_path) > 300:
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", self.report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
            with open(self.report_path, 'r', encoding='utf-8') as f:
                content = f.read()
                design_match = re.search(r'<DesignCapacity>(\d+)</DesignCapacity>', content)
                full_match = re.search(r'<FullChargeCapacity>(\d+)</FullChargeCapacity>', content)
                manuf_match = re.search(r'<Manufacturer>(.*?)</Manufacturer>', content)
                serial_match = re.search(r'<SerialNumber>(.*?)</SerialNumber>', content)
                chem_match = re.search(r'<Chemistry>(.*?)</Chemistry>', content)
                if design_match:
                    powercfg_data["design_capacity"] = int(design_match.group(1))
                if full_match:
                    powercfg_data["full_charge_capacity"] = int(full_match.group(1))
                if manuf_match:
                    powercfg_data["manufacturer"] = manuf_match.group(1).strip()
                if serial_match:
                    powercfg_data["serial_number"] = serial_match.group(1).strip()
                if chem_match:
                    powercfg_data["chemistry"] = chem_match.group(1).strip()
        except Exception as e:
            logging.warning(f"Powercfg report parsing failed: {e}")

        # Method 2: WMI Win32_Battery
        try:
            for bat in c.Win32_Battery():
                battery_name = bat.Name or "Unknown"
                if powercfg_data["design_capacity"] == "N/A" and bat.DesignCapacity:
                    powercfg_data["design_capacity"] = bat.DesignCapacity
                if powercfg_data["full_charge_capacity"] == "N/A" and bat.FullChargeCapacity:
                    powercfg_data["full_charge_capacity"] = bat.FullChargeCapacity
        except Exception as e:
            logging.warning(f"WMI Win32_Battery failed: {e}")

        # Method 3: ACPI via ctypes
        try:
            h_battery = ctypes.windll.kernel32.CreateFileW(
                r"\\.\Battery0",
                0x80000000 | 0x40000000,
                0,
                None,
                3,
                0,
                None
            )
            if h_battery != -1:
                battery_tag = wintypes.ULONG()
                success = ctypes.windll.kernel32.DeviceIoControl(
                    h_battery,
                    IOCTL_BATTERY_QUERY_TAG,
                    None,
                    0,
                    ctypes.byref(battery_tag),
                    ctypes.sizeof(battery_tag),
                    None,
                    None
                )
                if success:
                    bqi = BATTERY_QUERY_INFORMATION()
                    bqi.BatteryTag = battery_tag.value
                    bqi.InformationLevel = BATTERY_QUERY_INFORMATION_LEVEL
                    bi = BATTERY_INFORMATION()
                    success = ctypes.windll.kernel32.DeviceIoControl(
                        h_battery,
                        IOCTL_BATTERY_QUERY_INFORMATION,
                        ctypes.byref(bqi),
                        ctypes.sizeof(bqi),
                        ctypes.byref(bi),
                        ctypes.sizeof(bi),
                        None,
                        None
                    )
                    if success:
                        if powercfg_data["design_capacity"] == "N/A" and bi.DesignedPetBatteryCapacity > 0:
                            powercfg_data["design_capacity"] = bi.DesignedPetBatteryCapacity
                        if powercfg_data["full_charge_capacity"] == "N/A" and bi.FullChargedCapacity > 0:
                            powercfg_data["full_charge_capacity"] = bi.FullChargedCapacity
                        if powercfg_data["chemistry"] == "N/A" and bi.Chemistry:
                            chem = bytes(bi.Chemistry).decode('ascii', errors='ignore').strip()
                            powercfg_data["chemistry"] = chem if chem in ["LION", "LIPO", "NIMH", "NICD"] else "N/A"
                ctypes.windll.kernel32.CloseHandle(h_battery)
        except Exception as e:
            logging.warning(f"ACPI battery info failed: {e}")

        # Method 4: System Info
        try:
            for system in c.Win32_ComputerSystem():
                manufacturer = system.Manufacturer.strip() or "Unknown"
                model = system.Model.strip() or "Unknown"
                laptop_model = f"{manufacturer} {model}"
            for product in c.Win32_ComputerSystemProduct():
                version = product.Version.strip() or ""
                if version and version.lower() not in ["unknown", "system product name", "to be filled by o.e.m."]:
                    laptop_model += f" {version}"
            for processor in c.Win32_Processor():
                processor_name = processor.Name.strip() or ""
                if processor_name:
                    laptop_model += f" ({processor_name})"
        except Exception as e:
            logging.warning(f"System info failed: {e}")

        # Use BatteryIntelligence to consolidate data
        try:
            intel_data = self.battery_intel.get_hardware_info(battery, powercfg_data)
        except Exception as e:
            logging.warning(f"BatteryIntelligence failed: {e}")
            intel_data = powercfg_data

        percent = battery.percent if battery else "N/A"
        charging = "Charging" if battery and battery.power_plugged else "Discharging"
        time_left = self.calculate_time_remaining(battery, percent, intel_data["full_charge_capacity"])

        return {
            "laptop_model": laptop_model,
            "battery_name": battery_name,
            "design_capacity": intel_data["design_capacity"],
            "full_charge_capacity": intel_data["full_charge_capacity"],
            "percent": percent,
            "charging": charging,
            "time_left": time_left,
            "manufacturer": intel_data["manufacturer"],
            "serial_number": intel_data["serial_number"],
            "chemistry": intel_data["chemistry"]
        }

    def calculate_time_remaining(self, battery, percent, full_charge_capacity):
        if not battery or percent == "N/A" or full_charge_capacity == "N/A":
            return "N/A"
        try:
            percent = float(percent)
            full_charge_capacity = int(full_charge_capacity)
            if full_charge_capacity <= 0:
                return "N/A"
            current_capacity = (percent / 100) * full_charge_capacity
            max_seconds = 86400

            if battery.power_plugged:
                if percent < 100:
                    charge_rate = full_charge_capacity * 0.5
                    if charge_rate <= 0:
                        return "N/A"
                    remaining_capacity = full_charge_capacity - current_capacity
                    seconds = min((remaining_capacity / charge_rate) * 3600, max_seconds)
                    hours, remainder = divmod(int(seconds), 3600)
                    minutes = remainder // 60
                    return f"{hours}h {minutes}m to full charge"
                return "Fully Charged"
            else:
                if (battery.secsleft != psutil.POWER_TIME_UNLIMITED and
                        0 < battery.secsleft <= max_seconds):
                    hours, remainder = divmod(int(battery.secsleft), 3600)
                    minutes = remainder // 60
                    return f"{hours}h {minutes}m remaining"
                else:
                    discharge_rate = full_charge_capacity * 0.2
                    if discharge_rate <= 0:
                        return "N/A"
                    seconds = min((current_capacity / discharge_rate) * 3600, max_seconds)
                    hours, remainder = divmod(int(seconds), 3600)
                    minutes = remainder // 60
                    return f"{hours}h {minutes}m remaining"
        except (ValueError, TypeError, ZeroDivisionError) as e:
            logging.error(f"Time remaining calculation error: {e}")
            return "Calculating..."

    def get_usage_days(self):
        try:
            import wmi
            c = wmi.WMI()
            for os in c.Win32_OperatingSystem():
                install_date = os.InstallDate
                if install_date:
                    try:
                        install_date = datetime.datetime.strptime(install_date.split('.')[0], "%Y%m%d%H%M%S")
                        now = datetime.datetime.now()
                        delta = now - install_date
                        return max(delta.days, 1)
                    except ValueError as ve:
                        logging.error(f"Invalid date format in get_usage_days: {ve}")
                        return self.battery_intel.default_values["usage_days"]
            logging.warning("No valid OS install date found")
            return self.battery_intel.default_values["usage_days"]
        except Exception as e:
            logging.error(f"Error in get_usage_days: {e}")
            return self.battery_intel.default_values["usage_days"]

    def get_charge_frequency(self):
        if not self.charge_log:
            return 0.5
        charge_events = sum(1 for entry in self.charge_log if entry["is_charging"] and entry["cycle_contribution"] > 0)
        days = (self.charge_log[-1]["timestamp"] - self.charge_log[0]["timestamp"]) / (24 * 3600) if len(self.charge_log) > 1 else 1
        return charge_events / max(days, 1)
    
    def get_battery_age(self):
        """
        Estimates the battery age in years based on the system installation date.
        Returns a float representing the age in years, or a default value if data is unavailable.
        """
        try:
            usage_days = self.get_usage_days()
            battery_age_years = usage_days / DAYS_PER_YEAR
            # Cap the age at a reasonable maximum (e.g., 10 years) to prevent outliers
            battery_age_years = min(max(battery_age_years, 0.1), 10.0)
            return battery_age_years
        except Exception as e:
            logging.error(f"Failed to estimate battery age: {e}")
            # Return a default age (2 years) if calculation fails
            return 2.0
        
    def get_battery_temperature(self):
        temperature = "N/A"
        
        # Method 1: WMI MSAcpi_ThermalZoneTemperature (Primary Method)
        try:
            c = wmi.WMI(namespace="root\\wmi")
            batteries = c.MSAcpi_ThermalZoneTemperature()
            for battery in batteries:
                temp_kelvin = battery.CurrentTemperature / 10.0
                temperature = round(temp_kelvin - 273.15, 1)
                if temperature > -50 and temperature < 150:  # Sanity check
                    return temperature
        except Exception as e:
            logging.error(f"WMI MSAcpi_ThermalZoneTemperature error: {e}")

        # Method 2: WMI Battery Status with Temperature (Fallback)
        try:
            c = wmi.WMI()
            for bat in c.Win32_Battery():
                if hasattr(bat, "Temperature"):
                    temp = bat.Temperature
                    if temp and isinstance(temp, (int, float)) and temp > -50 and temp < 150:
                        temperature = round(temp, 1)
                        return temperature
        except Exception as e:
            logging.error(f"WMI Win32_Battery temperature error: {e}")

        # Method 3: Powercfg Battery Report (Alternative Fallback)
        try:
            if not os.path.exists(self.report_path) or time.time() - os.path.getmtime(self.report_path) > 300:
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", self.report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
            with open(self.report_path, 'r', encoding='utf-8') as f:
                content = f.read()
                temp_match = re.search(r'<Temperature>(\d+\.?\d*)</Temperature>', content)
                if temp_match:
                    temp = float(temp_match.group(1))
                    if temp > -50 and temp < 150:
                        temperature = round(temp, 1)
                        return temperature
        except Exception as e:
            logging.error(f"Powercfg temperature error: {e}")

        # Method 4: ACPI via ctypes (Advanced Fallback)
        try:
            device_path = r"\\.\ACPI#Battery"
            h_device = ctypes.windll.kernel32.CreateFileW(
                device_path,
                0x80000000 | 0x40000000,
                0,
                None,
                3,
                0,
                None
            )
            if h_device != -1:
                buffer_size = 1024
                buffer = ctypes.create_string_buffer(buffer_size)
                bytes_returned = wintypes.DWORD()
                success = ctypes.windll.kernel32.DeviceIoControl(
                    h_device,
                    0x00294040,  # IOCTL_ACPI_GET_TEMPERATURE (custom approximation)
                    None,
                    0,
                    buffer,
                    buffer_size,
                    ctypes.byref(bytes_returned),
                    None
                )
                if success:
                    temp_data = buffer.raw[:bytes_returned.value]
                    temp = struct.unpack_from('<f', temp_data)[0] if bytes_returned.value >= 4 else None
                    if temp and temp > -50 and temp < 150:
                        temperature = round(temp - 273.15, 1)
                        ctypes.windll.kernel32.CloseHandle(h_device)
                        return temperature
                ctypes.windll.kernel32.CloseHandle(h_device)
        except Exception as e:
            logging.error(f"ACPI ctypes temperature error: {e}")

        return temperature
    
    
    def get_battery_voltage_and_power(self):
        voltage = "N/A"
        power = "N/A"
        try:
            c = wmi.WMI(namespace="root\\wmi")
            for battery in c.BatteryStatus():
                if hasattr(battery, "Voltage") and battery.Voltage:
                    voltage = round(battery.Voltage / 1000.0, 2)
                    if hasattr(battery, "Current") and battery.Current:
                        power = round((battery.Voltage / 1000.0) * (battery.Current / 1000.0), 2)
                        return voltage, power
        except Exception as e:
            logging.error(f"WMI voltage/power error: {e}")
        # Fallback: Estimate power from discharge rate
        try:
            battery = psutil.sensors_battery()
            if battery and battery.percent and not battery.power_plugged:
                info = self.get_battery_info()
                full_capacity = info["full_charge_capacity"]
                if isinstance(full_capacity, int) and full_capacity > 0:
                    time_left = self.calculate_time_remaining(battery, battery.percent, full_capacity)
                    if isinstance(time_left, str) and "remaining" in time_left:
                        hours = 0
                        minutes = 0
                        if "h" in time_left:
                            hours = int(time_left.split("h")[0])
                            minutes = int(time_left.split("h")[1].split("m")[0])
                        else:
                            minutes = int(time_left.split("m")[0])
                        total_hours = hours + (minutes / 60.0)
                        if total_hours > 0:
                            current_capacity = (battery.percent / 100) * full_capacity
                            power = round(current_capacity / total_hours, 2)
                            return voltage, power
        except Exception as e:
            logging.error(f"Power estimation error: {e}")
        return voltage, power
    
    
    def get_battery_details(self, report_path):
        manufacturer = "N/A"
        serial_number = "N/A"
        chemistry = "N/A"

        CHEMISTRY_MAP = {
            "LION": "Lithium-Ion (Li-Ion)",
            "LI": "Lithium-Ion (Li-Ion)",
            "LI-ION": "Lithium-Ion (Li-Ion)",
            "LIION": "Lithium-Ion (Li-Ion)",
            "LIPO": "Lithium Polymer (Li-Poly)",
            "LIP": "Lithium Polymer (Li-Poly)",
            "LI-P": "Lithium Polymer (Li-Poly)",
            "LI-PO": "Lithium Polymer (Li-Poly)",
            "LIONP": "Lithium Polymer (Li-Poly)",
            "NICD": "Nickel-Cadmium (NiCd)",
            "NIMH": "Nickel-Metal Hydride (NiMH)",
            "PB": "Lead-Acid",
            "SLA": "Sealed Lead-Acid",
            "LFP": "Lithium Iron Phosphate (LiFePO4)",
            "LCO": "Lithium Cobalt Oxide (LiCoO2)",
            "NMC": "Lithium Nickel Manganese Cobalt (NMC)",
            "NCA": "Lithium Nickel Cobalt Aluminum (NCA)",
            "UNKNOWN": "N/A",
            "": "N/A"
        }

        def is_valid_data(value, is_serial=False):
            if not value or value.strip().lower() in ["unknown", "n/a", "", "to be filled by o.e.m.", "none"]:
                return False
            if is_serial:
                return len(value.strip()) >= 4 and any(c.isalnum() for c in value)
            return True

        # Method 1: Powercfg XML Report
        try:
            if not os.path.exists(report_path) or time.time() - os.path.getmtime(report_path) > 300:
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
            with open(report_path, 'r', encoding='utf-8') as f:
                content = f.read()
                manuf_match = re.search(r'<Manufacturer>(.*?)</Manufacturer>', content)
                serial_match = re.search(r'<SerialNumber>(.*?)</SerialNumber>', content)
                chem_match = re.search(r'<Chemistry>(.*?)</Chemistry>', content)
                if manuf_match and is_valid_data(manuf_match.group(1)):
                    manufacturer = manuf_match.group(1).strip()
                if serial_match and is_valid_data(serial_match.group(1), True):
                    serial_number = serial_match.group(1).strip()
                if chem_match and is_valid_data(chem_match.group(1)):
                    chem_value = chem_match.group(1).strip().upper()
                    chemistry = CHEMISTRY_MAP.get(chem_value, "N/A")
                if chemistry != "N/A":
                    self.battery_intel.cache['source'] = "Powercfg XML"
                    self.battery_intel.cache['manufacturer'] = manufacturer
                    self.battery_intel.cache['serial_number'] = serial_number
                    self.battery_intel.cache['chemistry'] = chemistry
                    logging.info(f"Powercfg XML: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                    return manufacturer, serial_number, chemistry
        except Exception as e:
            logging.warning(f"Powercfg XML report failed: {e}")

        # Method 2: Powercfg HTML Report
        try:
            temp_dir = tempfile.gettempdir()
            html_report_path = os.path.join(temp_dir, "battery_report.html")
            if not os.path.exists(html_report_path) or time.time() - os.path.getmtime(html_report_path) > 300:
                subprocess.run(["powercfg", "/batteryreport", "/output", html_report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
            with open(html_report_path, 'r', encoding='utf-8') as f:
                soup = BeautifulSoup(f, 'html.parser')
                battery_section = soup.find(string=re.compile("Installed batteries", re.I))
                if battery_section:
                    battery_table = battery_section.find_next('table')
                    if battery_table:
                        rows = battery_table.find_all('tr')
                        for row in rows:
                            cells = row.find_all('td')
                            if len(cells) >= 2:
                                label = cells[0].text.strip().upper()
                                value = cells[1].text.strip()
                                if "MANUFACTURER" in label and is_valid_data(value):
                                    manufacturer = value
                                if "SERIAL NUMBER" in label and is_valid_data(value, True):
                                    serial_number = value
                                if "CHEMISTRY" in label and is_valid_data(value):
                                    chemistry = CHEMISTRY_MAP.get(value.upper(), "N/A")
                        if chemistry != "N/A":
                            self.battery_intel.cache['source'] = "Powercfg HTML"
                            self.battery_intel.cache['manufacturer'] = manufacturer
                            self.battery_intel.cache['serial_number'] = serial_number
                            self.battery_intel.cache['chemistry'] = chemistry
                            logging.info(f"Powercfg HTML: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                            return manufacturer, serial_number, chemistry
        except Exception as e:
            logging.warning(f"Powercfg HTML report failed: {e}")

        # Method 3: WMI Win32_Battery
        try:
            c = wmi.WMI()
            for bat in c.Win32_Battery():
                if hasattr(bat, "Manufacturer") and is_valid_data(bat.Manufacturer):
                    manufacturer = bat.Manufacturer.strip()
                if hasattr(bat, "SerialNumber") and is_valid_data(bat.SerialNumber, True):
                    serial_number = bat.SerialNumber.strip()
                if hasattr(bat, "Chemistry") and bat.Chemistry and is_valid_data(bat.Chemistry):
                    chemistry = CHEMISTRY_MAP.get(bat.Chemistry.upper(), "N/A")
                if chemistry != "N/A":
                    self.battery_intel.cache['source'] = "Win32_Battery"
                    self.battery_intel.cache['manufacturer'] = manufacturer
                    self.battery_intel.cache['serial_number'] = serial_number
                    self.battery_intel.cache['chemistry'] = chemistry
                    logging.info(f"WMI Win32_Battery: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                    return manufacturer, serial_number, chemistry
        except Exception as e:
            logging.warning(f"WMI Win32_Battery failed: {e}")

        # Method 4: WMI BatteryStaticData
        try:
            c = wmi.WMI(namespace="root\\wmi")
            for bat in c.BatteryStaticData():
                if hasattr(bat, "Manufacturer") and is_valid_data(bat.Manufacturer):
                    manufacturer = bat.Manufacturer.strip()
                if hasattr(bat, "SerialNumber") and is_valid_data(bat.SerialNumber, True):
                    serial_number = bat.SerialNumber.strip()
                if hasattr(bat, "Chemistry") and bat.Chemistry and is_valid_data(bat.Chemistry):
                    chemistry = CHEMISTRY_MAP.get(bat.Chemistry.upper(), "N/A")
                if chemistry != "N/A":
                    self.battery_intel.cache['source'] = "BatteryStaticData"
                    self.battery_intel.cache['manufacturer'] = manufacturer
                    self.battery_intel.cache['serial_number'] = serial_number
                    self.battery_intel.cache['chemistry'] = chemistry
                    logging.info(f"WMI BatteryStaticData: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                    return manufacturer, serial_number, chemistry
        except Exception as e:
            logging.warning(f"WMI BatteryStaticData failed: {e}")

        # Method 5: WMI BatteryStatus
        try:
            c = wmi.WMI(namespace="root\\wmi")
            for bat in c.BatteryStatus():
                if hasattr(bat, "Chemistry") and bat.Chemistry and is_valid_data(bat.Chemistry):
                    chemistry = CHEMISTRY_MAP.get(bat.Chemistry.upper(), "N/A")
                if chemistry != "N/A":
                    self.battery_intel.cache['source'] = "BatteryStatus"
                    self.battery_intel.cache['manufacturer'] = manufacturer
                    self.battery_intel.cache['serial_number'] = serial_number
                    self.battery_intel.cache['chemistry'] = chemistry
                    logging.info(f"WMI BatteryStatus: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                    return manufacturer, serial_number, chemistry
        except Exception as e:
            logging.warning(f"WMI BatteryStatus failed: {e}")

        # Method 6: ACPI via ctypes
        try:
            for device_path in self.battery_intel.find_battery_devices():
                h_battery = ctypes.windll.kernel32.CreateFileW(
                    device_path,
                    0x80000000 | 0x40000000,
                    0,
                    None,
                    3,
                    0,
                    None
                )
                if h_battery != -1:
                    try:
                        battery_tag = wintypes.ULONG()
                        success = ctypes.windll.kernel32.DeviceIoControl(
                            h_battery,
                            IOCTL_BATTERY_QUERY_TAG,
                            None,
                            0,
                            ctypes.byref(battery_tag),
                            ctypes.sizeof(battery_tag),
                            None,
                            None
                        )
                        if success:
                            bqi = BATTERY_QUERY_INFORMATION()
                            bqi.BatteryTag = battery_tag.value
                            # Query Chemistry
                            bqi.InformationLevel = BATTERY_QUERY_INFORMATION_LEVEL.BatteryInformation
                            bi = BATTERY_INFORMATION()
                            success = ctypes.windll.kernel32.DeviceIoControl(
                                h_battery,
                                IOCTL_BATTERY_QUERY_INFORMATION,
                                ctypes.byref(bqi),
                                ctypes.sizeof(bqi),
                                ctypes.byref(bi),
                                ctypes.sizeof(bi),
                                None,
                                None
                            )
                            if success and bi.Chemistry:
                                chem = bytes(bi.Chemistry).decode('ascii', errors='ignore').strip().upper()
                                chemistry = CHEMISTRY_MAP.get(chem, "N/A")
                            # Query Manufacturer and Serial Number
                            for info_level, buffer_size in [
                                (BATTERY_QUERY_INFORMATION_LEVEL.BatteryManufactureName, 128),
                                (BATTERY_QUERY_INFORMATION_LEVEL.BatterySerialNumber, 128)
                            ]:
                                bqi.InformationLevel = info_level
                                buffer = ctypes.create_string_buffer(buffer_size)
                                success = ctypes.windll.kernel32.DeviceIoControl(
                                    h_battery,
                                    IOCTL_BATTERY_QUERY_INFORMATION,
                                    ctypes.byref(bqi),
                                    ctypes.sizeof(bqi),
                                    buffer,
                                    buffer_size,
                                    None,
                                    None
                                )
                                if success:
                                    value = buffer.value.decode('ascii', errors='ignore').strip()
                                    if info_level == BATTERY_QUERY_INFORMATION_LEVEL.BatteryManufactureName and is_valid_data(value):
                                        manufacturer = value
                                    elif info_level == BATTERY_QUERY_INFORMATION_LEVEL.BatterySerialNumber and is_valid_data(value, True):
                                        serial_number = value
                            if chemistry != "N/A":
                                self.battery_intel.cache['source'] = f"ACPI ({device_path})"
                                self.battery_intel.cache['manufacturer'] = manufacturer
                                self.battery_intel.cache['serial_number'] = serial_number
                                self.battery_intel.cache['chemistry'] = chemistry
                                logging.info(f"ACPI: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                                return manufacturer, serial_number, chemistry
                    finally:
                        ctypes.windll.kernel32.CloseHandle(h_battery)
        except Exception as e:
            logging.warning(f"ACPI ctypes failed: {e}")

        # Method 7: Registry Query
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Class\{72631E54-78A4-11D0-BCF7-00AA00B7B32A}")
            for i in range(winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    device_desc, _ = winreg.QueryValueEx(subkey, "DeviceDescription")
                    if "battery" in device_desc.lower():
                        try:
                            manuf, _ = winreg.QueryValueEx(subkey, "ProviderName")
                            if is_valid_data(manuf):
                                manufacturer = manuf.strip()
                        except:
                            pass
                        try:
                            serial, _ = winreg.QueryValueEx(subkey, "SerialNumber")
                            if is_valid_data(serial, True):
                                serial_number = serial.strip()
                        except:
                            pass
                        try:
                            chem, _ = winreg.QueryValueEx(subkey, "Chemistry")
                            if is_valid_data(chem):
                                chemistry = CHEMISTRY_MAP.get(chem.upper(), "N/A")
                        except:
                            pass
                        if chemistry != "N/A":
                            self.battery_intel.cache['source'] = "Registry"
                            self.battery_intel.cache['manufacturer'] = manufacturer
                            self.battery_intel.cache['serial_number'] = serial_number
                            self.battery_intel.cache['chemistry'] = chemistry
                            logging.info(f"Registry: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
                            return manufacturer, serial_number, chemistry
                except:
                    pass
                finally:
                    winreg.CloseKey(subkey)
            winreg.CloseKey(key)
        except Exception as e:
            logging.warning(f"Registry query failed: {e}")

        # Fallback: Default to Lithium-Ion if no valid chemistry is found
        chemistry = "Lithium-Ion (Li-Ion)"
        self.battery_intel.cache['source'] = "Default"
        self.battery_intel.cache['manufacturer'] = manufacturer
        self.battery_intel.cache['serial_number'] = serial_number
        self.battery_intel.cache['chemistry'] = chemistry
        logging.info(f"Default fallback: manufacturer={manufacturer}, serial_number={serial_number}, chemistry={chemistry}")
        return manufacturer, serial_number, chemistry
    
    
    def get_cycle_count_wmic(self):
        try:
            result = subprocess.run(["wmic", "path", "Win32_Battery", "get", "CycleCount"], capture_output=True, text=True, check=True)
            output = result.stdout.strip().split('\n')[1].strip()  # Skip header
            if output and output.isdigit():
                cycle_count = int(output)
                if cycle_count > 0:
                    logging.info(f"Cycle count from WMIC: {cycle_count}")
                    return cycle_count, False, "WMIC"
            return None, False, "WMIC Failed"
        except Exception as e:
            logging.error(f"WMIC cycle count failed: {e}")
            return None, False, "WMIC Failed"
    
    
    def get_battery_info_wmic(self):
        try:
            result = subprocess.run(["wmic", "path", "Win32_Battery", "get", "DesignCapacity,FullChargeCapacity,Name"], capture_output=True, text=True, check=True)
            output = result.stdout.strip().split('\n')[1].strip()  # Skip header
            if output:
                parts = output.split()
                if len(parts) >= 2:
                    design_capacity = int(parts[0]) if parts[0].isdigit() else "N/A"
                    full_charge_capacity = int(parts[1]) if parts[1].isdigit() else "N/A"
                    battery_name = " ".join(parts[2:]) if len(parts) > 2 else "Unknown"
                    return {
                        "design_capacity": design_capacity,
                        "full_charge_capacity": full_charge_capacity,
                        "battery_name": battery_name
                    }
            return {"design_capacity": "N/A", "full_charge_capacity": "N/A", "battery_name": "Unknown"}
        except Exception as e:
            logging.error(f"WMIC battery info failed: {e}")
        return {"design_capacity": "N/A", "full_charge_capacity": "N/A", "battery_name": "Unknown"}



    def generate_summary(self, health_percent, username, remaining_life, cycle_count, total_cycles, usage_days, design_capacity, full_charge_capacity):
        cycles_per_day = cycle_count / usage_days if usage_days > 0 and cycle_count is not None else 0
        remaining_cycles = max(total_cycles - cycle_count, 0) if cycle_count is not None else 0
        charge_frequency = self.get_charge_frequency()
        
        base_message = (
            f"Hi {username}, let’s talk about your laptop’s battery! A battery cycle is like a full charge from 0% to 100%. "
            f"Your battery has gone through {cycle_count} cycles out of {total_cycles} it’s designed for, leaving {remaining_cycles} cycles to go. "
            f"Over {usage_days} days of use, you’re averaging {cycles_per_day:.2f} cycles per day. "
            f"You charge about {charge_frequency:.2f} times per day, which affects how long your battery will last. "
            f"We estimate your battery has {remaining_life} left before it needs replacing."
        )

        if health_percent > 80:
            message = (
                f"{base_message}\n\nYour battery is in top shape, like a champion athlete! You’re doing great keeping it healthy. "
                f"Keep up the good work, and it’ll stay strong for a long time. Check back anytime for more insights!"
            )
            color = THEMES[self.current_theme]["alert_good"]
        elif health_percent > 60:
            message = (
                f"{base_message}\n\nYour battery is doing well, like a reliable friend. It’s in good health, but a little care can go a long way. "
                f"Try charging between 20-80% and avoid extreme heat. We’re here to help whenever you need us!"
            )
            color = THEMES[self.current_theme]["alert_good"]
        elif health_percent > 40:
            message = (
                f"{base_message}\n\nYour battery is starting to show its age, like a seasoned traveler. It’s still got some life, but keep an eye on it. "
                f"Avoid draining it to 0% often, and consider a replacement in the near future. Pop back for updates!"
            )
            color = THEMES[self.current_theme]["alert_warning"]
            if not self.has_shown_health_alert:
                QMessageBox.warning(self, "Battery Alert", f"Hey {username}, your battery’s getting tired. Start planning for a replacement soon!")
                self.has_shown_health_alert = True
        else:
            message = (
                f"{base_message}\n\nOh, {username}, your battery’s worked hard and needs a rest. It’s nearing the end of its life. "
                f"Plan to replace it soon to keep your laptop running smoothly. We’re here to guide you every step of the way!"
            )
            color = THEMES[self.current_theme]["alert_critical"]
            if not self.has_shown_health_alert:
                QMessageBox.critical(self, "Battery Alert", f"Hey {username}, your battery’s almost done! Time to get a new one ASAP!")
                self.has_shown_health_alert = True

        return message, color
    
        # Default values for battery information
    def update_data(self):
        try:
            info = self.get_battery_info()
            if self.manual_cycle_count is not None:
                cycle_count = self.manual_cycle_count
                is_estimated = self.manual_is_estimated
                cycle_source = self.manual_cycle_source
            else:
                battery = psutil.sensors_battery()
                cycle_count, is_estimated, cycle_source = self.battery_intel.get_cycle_count(battery, self.report_path)
                current_time = time.time()
                self.save_cycle_cache(cycle_count, is_estimated, current_time)
            total_cycles = info["total_cycles"]
        except Exception as e:
            logging.error(f"Error updating data: {e}")
            info = {
                "laptop_model": "Unknown",
                "battery_name": "Unknown",
                "design_capacity": "N/A",
                "full_charge_capacity": "N/A",
                "percent": "N/A",
                "charging": "Unknown",
                "time_left": "N/A",
                "manufacturer": "N/A",
                "serial_number": "N/A",
                "chemistry": "N/A",
                "cycle_count": 0,
                "cycle_source": "Default",
                "is_estimated": True,
                "total_cycles": 600
            }
            if self.manual_cycle_count is not None:
                cycle_count = self.manual_cycle_count
                is_estimated = self.manual_is_estimated
                cycle_source = self.manual_cycle_source
            else:
                cycle_count = 0
                is_estimated = True
                cycle_source = "Default"
            total_cycles = 600

        # Log battery history
        history_file = os.path.join(self.appdata_path, "battery_history.csv")
        if not os.path.exists(history_file):
            with open(history_file, "w") as f:
                f.write("timestamp,percentage\n")
        if info["percent"] != "N/A":
            try:
                percentage = float(info["percent"])
                timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                with open(history_file, "a") as f:
                    f.write(f"{timestamp},{percentage}\n")
            except Exception as e:
                logging.error(f"Failed to log battery history: {e}")

        # Trim history file periodically
        self.update_counter += 1
        if self.update_counter % 10 == 0:
            self.trim_history_file()

        laptop_model = info["laptop_model"]
        self.laptop_label.setText(f"<b>Laptop Model:</b> {laptop_model}")
        self.battery_model_label.setText(f"<b>Battery Model:</b> {info['battery_name']}")
        manufacturer, serial_number, chemistry = self.get_battery_details(self.report_path)
        self.manufacturer_label.setText(f"<b>Battery Manufacturer:</b> {manufacturer}")
        self.serial_number_label.setText(f"<b>Battery Serial Number:</b> {serial_number}")
        self.chemistry_label.setText(f"<b>Battery Chemistry:</b> {chemistry}")
        self.design_capacity_label.setText(f"<b>Design Capacity:</b> {info['design_capacity']:,} mWh" if isinstance(info['design_capacity'], int) else "<b>Design Capacity:</b> N/A")
        self.full_charge_label.setText(f"<b>Full Charge Capacity:</b> {info['full_charge_capacity']:,} mWh" if isinstance(info['full_charge_capacity'], int) else "<b>Full Charge Capacity:</b> N/A")
        self.charging_label.setText(f"<b>Charging State:</b> {info['charging']}")
        self.percentage_label.setText(f"<b>Battery Percentage:</b> {info['percent']}%")
        self.time_left_label.setText(f"<b>Estimated Time Remaining:</b> {info['time_left']}")
        self.source_label.setText(f"<b>Cycle Source:</b> {cycle_source}")
        self.data_source_label.setText(f"<b>Data Source:</b> {self.battery_intel.cache.get('source', 'N/A')}")

        cycle_display = f"{cycle_count} (Estimated)" if is_estimated and cycle_count > 0 else str(cycle_count) if cycle_count >= 0 else "N/A"
        self.cycle_count_label.setText(f"<b>Used Cycle Count:</b> {cycle_display}")
        usage_days = self.battery_intel.get_usage_days()
        self.usage_duration_label.setText(f"<b>Usage Duration:</b> {usage_days} days")
        design_capacity = info["design_capacity"]
        full_charge_capacity = info["full_charge_capacity"]

        if info["percent"] != "N/A":
            try:
                charge_percent = float(info["percent"])
                self.charging_animation = QPropertyAnimation(self.charging_widget, b"healthPercent")
                self.charging_animation.setDuration(1000)
                self.charging_animation.setStartValue(self.charging_widget.health_percent)
                self.charging_animation.setEndValue(charge_percent)
                self.charging_animation.setEasingCurve(QEasingCurve.InOutQuad)
                self.charging_animation.start()
            except (ValueError, TypeError) as e:
                logging.error(f"Error updating charging widget percentage: {e}")
                self.charging_widget.set_health_percent(0)
        else:
            self.charging_widget.set_health_percent(0)

        if isinstance(design_capacity, int) and isinstance(full_charge_capacity, int) and design_capacity > 0 and full_charge_capacity > 0:
            health_percent = self.battery_intel.calculate_health(cycle_count, total_cycles, design_capacity, full_charge_capacity)
            self.battery_animation.setStartValue(self.battery_widget.health_percent)
            self.battery_animation.setEndValue(health_percent)
            self.battery_animation.start()
            if health_percent > 80:
                status = "Good"
                color = THEMES[self.current_theme]["alert_good"]
            elif health_percent > 60:
                status = "Moderate"
                color = THEMES[self.current_theme]["alert_warning"]
            else:
                status = "Poor"
                color = THEMES[self.current_theme]["alert_critical"]
            self.health_status_label.setText(f'<b>Health Status:</b> <span style="color: {color}">{status}</span>')
        else:
            self.battery_widget.set_health_percent(0)
            self.health_status_label.setText(f"<b>Health Status:</b> N/A")

        temperature = self.get_battery_temperature()
        self.temperature_label.setText(f"<b>Battery Temperature:</b> {temperature} °C")
        if temperature != "N/A" and isinstance(temperature, (int, float)):
            if temperature > 45 and not self.high_temp_notified:
                self.show_notification("High Battery Temperature", "Your battery temperature is high. Please cool down your device.")
                self.high_temp_notified = True
            elif temperature <= 45:
                self.high_temp_notified = False

        voltage, power = self.get_battery_voltage_and_power()
        self.voltage_label.setText(f"<b>Battery Voltage:</b> {voltage} V")
        self.power_label.setText(f"<b>Power Consumption:</b> {power} W")

        if info["percent"] != "N/A" and isinstance(info["percent"], (int, float)):
            battery_percent = float(info["percent"])
            if battery_percent < 20 and info["charging"] == "Discharging" and not self.low_battery_notified:
                self.show_notification("Low Battery", "Your battery is low. Please plug in your charger.")
                self.low_battery_notified = True
            elif battery_percent >= 20:
                self.low_battery_notified = False

        if health_percent is not None and isinstance(health_percent, (int, float)):
            if health_percent < 30 and not self.low_health_notified:
                self.show_notification("Battery Health Alert", "Your battery health is critically low. Consider replacing it.")
                self.low_health_notified = True
            elif health_percent >= 30:
                self.low_health_notified = False

        username = os.getlogin()
        health_percent = self.battery_intel.calculate_health(cycle_count, total_cycles, design_capacity, full_charge_capacity)
        # Calculate remaining_life
        remaining_life = self.battery_intel.estimate_remaining_life(cycle_count, total_cycles, usage_days, design_capacity, full_charge_capacity)
        summary_text, summary_color = self.generate_summary(health_percent, username, remaining_life, cycle_count, total_cycles, usage_days, design_capacity, full_charge_capacity)
        self.summary_label.setText(summary_text)
        self.summary_label.setStyleSheet(f"padding: 10px; background: transparent; color: {summary_color};")
        
        
    def trim_history_file(self):
        history_file = os.path.join(self.appdata_path, "battery_history.csv")
        if os.path.exists(history_file):
            try:
                with open(history_file, "r") as f:
                    lines = f.readlines()
                if len(lines) > 1001:  # Header + 1000 entries
                    with open(history_file, "w") as f:
                        f.writelines(lines[-1000:])
            except Exception as e:
                logging.error(f"Failed to trim history file: {e}")
                
    def show_history(self):
        history_dialog = HistoryDialog(self)
        history_dialog.exec_()
        
        
    def refresh_data(self):
        self.manual_cycle_count = None
        self.update_data()

    # Update get_battery_info to use BatteryIntelligence
    def get_battery_info(self):
        battery = psutil.sensors_battery()
        c = wmi.WMI()
        laptop_model = "Unknown"
        powercfg_data = {}

        try:
            if not os.path.exists(self.report_path) or time.time() - os.path.getmtime(self.report_path) > 3600:
                subprocess.run(["powercfg", "/batteryreport", "/xml", "/output", self.report_path], check=True, shell=True, timeout=30)
                time.sleep(1)
            with open(self.report_path, 'r', encoding='utf-8') as f:
                content = f.read()
                design_match = re.search(r'<DesignCapacity>(\d+)</DesignCapacity>', content)
                full_match = re.search(r'<FullChargeCapacity>(\d+)</FullChargeCapacity>', content)
                cycle_match = re.search(r'<CycleCount>(\d+)</CycleCount>', content)
                manuf_match = re.search(r'<Manufacturer>(.*?)</Manufacturer>', content)
                serial_match = re.search(r'<SerialNumber>(.*?)</SerialNumber>', content)
                chem_match = re.search(r'<Chemistry>(.*?)</Chemistry>', content)
                if design_match:
                    powercfg_data["design_capacity"] = int(design_match.group(1))
                if full_match:
                    powercfg_data["full_charge_capacity"] = int(full_match.group(1))
                if cycle_match:
                    powercfg_data["cycle_count"] = int(cycle_match.group(1))
                if manuf_match:
                    powercfg_data["manufacturer"] = manuf_match.group(1).strip()
                if serial_match:
                    powercfg_data["serial_number"] = serial_match.group(1).strip()
                if chem_match:
                    powercfg_data["chemistry"] = chem_match.group(1).strip()
        except Exception as e:
            logging.error(f"Powercfg report parsing error: {e}")

        try:
            for system in c.Win32_ComputerSystem():
                manufacturer = system.Manufacturer.strip() or "Unknown"
                model = system.Model.strip() or "Unknown"
                laptop_model = f"{manufacturer} {model}"
            for product in c.Win32_ComputerSystemProduct():
                version = product.Version.strip() or ""
                if version and version.lower() not in ["unknown", "system product name", "to be filled by o.e.m."]:
                    laptop_model += f" {version}"
            for processor in c.Win32_Processor():
                processor_name = processor.Name.strip() or ""
                if processor_name:
                    laptop_model += f" ({processor_name})"
        except Exception as e:
            logging.error(f"Laptop model error: {e}")

        intel_data = self.battery_intel.get_hardware_info(battery, powercfg_data, self.report_path)
        self.battery_intel.log_charge_event(battery)

        percent = battery.percent if battery else "N/A"
        charging = "Charging" if battery and battery.power_plugged else "Discharging"
        time_left = self.calculate_time_remaining(battery, percent, intel_data["full_charge_capacity"])

        return {
            "laptop_model": laptop_model,
            "battery_name": intel_data["battery_name"],
            "design_capacity": intel_data["design_capacity"],
            "full_charge_capacity": intel_data["full_charge_capacity"],
            "percent": percent,
            "charging": charging,
            "time_left": time_left,
            "manufacturer": intel_data["manufacturer"],
            "serial_number": intel_data["serial_number"],
            "chemistry": intel_data["chemistry"],
            "cycle_count": intel_data["cycle_count"],
            "cycle_source": intel_data.get("cycle_source", "N/A"),
            "is_estimated": intel_data.get("is_estimated", False),
            "total_cycles": intel_data["total_cycles"]
        }

    def show_next_tip(self):
        self.current_tip_index = (self.current_tip_index + 1) % len(TIPS)
        self.tip_label.setText(TIPS[self.current_tip_index])
        
    def show_previous_tip(self):
        self.current_tip_index = (self.current_tip_index - 1) % len(TIPS)
        self.tip_label.setText(TIPS[self.current_tip_index])
        
        
    def setup_system_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(resource_path("logo.ico")))
        tray_menu = QMenu()
        
        open_action = QAction("Open", self)
        open_action.triggered.connect(self.show)
        tray_menu.addAction(open_action)
        
        refresh_action = QAction("Refresh", self)
        refresh_action.triggered.connect(self.update_data)
        tray_menu.addAction(refresh_action)
        
        exit_action = QAction("Exit", self)
        exit_action.triggered.connect(self.close)
        tray_menu.addAction(exit_action)
        
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()


# Constants
DAYS_PER_YEAR = 365.25
DAYS_PER_MONTH = 30.44
IOCTL_BATTERY_QUERY_TAG = 0x294040
IOCTL_BATTERY_QUERY_INFORMATION = 0x294044

class BATTERY_QUERY_INFORMATION_LEVEL:
    BatteryInformation = 0
    BatteryGranularityInformation = 1
    BatteryTemperature = 2
    BatteryEstimatedTime = 3
    BatteryDeviceName = 4
    BatteryManufactureDate = 5
    BatteryManufactureName = 6
    BatteryUniqueID = 7
    BatterySerialNumber = 8

class BATTERY_QUERY_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("BatteryTag", wintypes.ULONG),
        ("InformationLevel", ctypes.c_int),
        ("AtRate", wintypes.LONG),
    ]

class BATTERY_INFORMATION(ctypes.Structure):
    _fields_ = [
        ("Capabilities", wintypes.ULONG),
        ("Technology", wintypes.BYTE),
        ("Reserved", wintypes.BYTE * 3),
        ("Chemistry", wintypes.BYTE * 4),
        ("DesignedPetBatteryCapacity", wintypes.ULONG),
        ("FullChargedCapacity", wintypes.ULONG),
        ("DefaultAlert1", wintypes.ULONG),
        ("DefaultAlert2", wintypes.ULONG),
        ("CriticalBias", wintypes.ULONG),
        ("CycleCount", wintypes.ULONG),
    ]

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if __name__ == '__main__':
    def is_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False

    if platform.system() != "Windows":
        print("This application is designed for Windows only.")
        sys.exit(1)

    if not is_admin():
        # Relaunch as admin
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit(0)  # Exit the current non-admin instance
        except Exception as e:
            print(f"Failed to elevate privileges: {e}")
            sys.exit(1)

    app = QApplication(sys.argv)
    splash = SplashScreen()
    splash.show()
    QTimer.singleShot(3000, splash.close)
    window = BatteryZ()
    QTimer.singleShot(3000, window.show)
    sys.exit(app.exec_())

"""
Dear User,
Battery-Z is crafted with boundless love and precision to give your laptop’s battery the care it deserves. 
With military-grade algorithms, stunning visuals, and heartfelt messages, every line of code is designed to empower you with knowledge and peace of mind. 
Enjoy your battery monitoring journey, and may your laptop stay powered with love for years to come!
- Muhammad Kashan Tariq
""" 