# ======================================================================================================================
#                                                                                                                      # 
#                       ██████╗  █████╗ ████████╗████████╗███████╗██████╗ ██╗   ██╗      ███████╗                      # 
#                       ██╔══██╗██╔══██╗╚══██╔══╝╚══██╔══╝██╔════╝██╔══██╗╚██╗ ██╔╝      ╚══███╔╝                      # 
#                       ██████╔╝███████║   ██║      ██║   █████╗  ██████╔╝ ╚████╔╝ █████╗  ███╔╝                       #     
#                       ██╔══██╗██╔══██║   ██║      ██║   ██╔══╝  ██╔══██╗  ╚██╔╝  ╚════╝ ███╔╝                        # 
#                       ██████╔╝██║  ██║   ██║      ██║   ███████╗██║  ██║   ██║         ███████╗                      # 
#                       ╚═════╝ ╚═╝  ╚═╝   ╚═╝      ╚═╝   ╚══════╝╚═╝  ╚═╝   ╚═╝         ╚══════╝ (V2.1.0)             # 
# ======================================================================================================================
#                           
# 88888888ba                
# 88      "8b               
# 88      ,8P               
# 88aaaaaa8P'  8b       d8  
# 88""""""8b,  `8b     d8'  
# 88      `8b   `8b   d8'   
# 88      a8P    `8b,d8'    
# 88888888P"       Y88'     
#                  d8'      
#                 d8'   
# =========================   
#                                        __                __                    ___       __     __  
#  |\/| |  | |__|  /\   |\/|  |\/|  /\  |  \    |__/  /\  /__` |__|  /\  |\ |     |   /\  |__) | /  \ 
#  |  | \__/ |  | /~~\  |  |  |  | /~~\ |__/    |  \ /~~\ .__/ |  | /~~\ | \|     |  /~~\ |  \ | \__X 
# ======================================================================================================================
#                                                                                                     
#   -- PROJECT:         Battery-Z: A Universal Battery Diagnostic and Predictive Lifespan Utility
#   -- FILE:            main.py
#   -- DESCRIPTION:     A monolithic, military-grade Python application for providing comprehensive battery
#                       diagnostics, health analysis, and predictive Remaining Useful Life (RUL) estimation
#                       for all Windows-based computing devices.
#   
#   -- AUTHOR:          Muhammad Kashan Tariq
#   -- TITLE:           Software Engineer
#   -- VERSION:         2.1.0
#   -- CREATION DATE:   2025-10-28
#
# ======================================================================================================================

# ======================================================================================================================
# ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
# ======================================================================================================================
#
#                                           --- ARCHITECTURAL BLUEPRINT ---
#
#   This application is architected as a "Modular Monolith". Despite being contained within a single file, the
#   code is logically partitioned into distinct, high-cohesion classes that manage specific domains of
#   responsibility. This approach ensures maintainability, readability, and scalability while adhering to the
#   single-file deployment constraint.
#
#   The primary components are:
#
#   1.  StateManager: A centralized, observable data store that holds the complete state of the application.
#       It acts as the single source of truth, decoupling data acquisition from the UI.
#
#   2.  DataFetcher (Worker): The engine for all system interrogation. It operates on a separate thread and
#       implements the "Strategic Data Cascade" for each data point, using WMI, ctypes, subprocess calls,
#       and other methods to ensure maximum data fidelity across all hardware.
#
#   3.  RULPredictor: The core intelligence of the application. This class contains the sophisticated,
#       non-linear algorithm for calculating the battery's Remaining Useful Life, leveraging the embedded
#       knowledge base and real-time data.
#
#   4.  SplashScreen (QWidget): A dedicated splash screen that displays during the initial intensive data
#       fetching and analysis phase, providing user feedback and ensuring the main UI is only shown when
#       fully responsive.
#
#   5.  MainWindow (QMainWindow): The primary application window. It is responsible for constructing and managing
#       all UI widgets, connecting signals to slots, and orchestrating the user experience. It reads data
#       exclusively from the StateManager to update its view.
#
#   6.  ThemeManager: A utility class that manages the application's visual appearance by loading and applying
#       Qt Style Sheets (QSS) for various themes.
#
#   7.  AnimatedBatteryWidget (QWidget): A custom, self-contained widget for the animated "liquid battery"
#       visualizations, ensuring smooth, hardware-friendly animations.
#
#   Data Flow:
#   [Startup] -> SplashScreen -> DataFetcher (Worker Thread) -> RULPredictor -> StateManager -> MainWindow -> [UI Display]
#
# ======================================================================================================================
# ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
# ======================================================================================================================

# ============================================================================
# main.py - RECONSTRUCTED
# Version: 2.0.0 (Integrated)
# Objective: Combine the superior UI of main.py with the robust, accurate
#            backend logic and algorithms from reference.py. This file is
#            production-ready, extensively documented, and optimized for
#            performance and accuracy on the Windows platform.
# ============================================================================


# ============================================================================
# SECTION 1: IMPORTS AND DEPENDENCIES
# Description: This section imports all necessary libraries for the application.
#              It combines imports from both the original main.py and reference.py
#              to ensure all functionalities (UI, data fetching, analysis) are
#              supported. Dependencies are grouped by standard library,
#              third-party, and PyQt5 for clarity.
# ============================================================================

# --- Standard Library Imports ---
import sys                          # Required for system-specific parameters and functions, like exiting the app.
import os                           # Provides a way of using operating system dependent functionality.
import platform                     # Used to retrieve platform-identifying data like OS and architecture.
import subprocess                   # Allows running external commands (e.g., powercfg, powershell).
import tempfile                     # Used for creating temporary files, specifically for the battery report.
import json                         # For reading and writing cache files in JSON format.
import datetime                     # Provides classes for manipulating dates and times.
import traceback                    # For printing stack traces when an error occurs, crucial for debugging.
import getpass                      # To get the current user's username for a personalized greeting.
import threading                    # For running data-fetching tasks in the background without freezing the UI.
import time                         # Provides time-related functions, used for caching and delays.
import re                           # Regular expressions for parsing text output from command-line tools.
import math                         # For mathematical operations in battery health calculations.
import warnings                     # To control warning messages, used here to ignore specific warnings.
import random                       # For selecting random welcome quotes and tips.
import winreg                       # To access the Windows Registry for fallback data retrieval.
import ctypes                       # To call functions in DLLs/shared libraries (e.g., Windows kernel32.dll).
from ctypes import wintypes         # Provides Windows-specific data types for ctypes.
from pathlib import Path            # For object-oriented filesystem paths.
from typing import Dict, List, Optional, Tuple, Any, Union # For type hinting to improve code clarity and maintainability.
from dataclasses import dataclass, field # For creating simple classes primarily for storing data.
from collections import defaultdict, deque # Advanced container datatypes.
from enum import Enum               # For creating enumerations.
import logging                      # For logging errors, warnings, and info to a file for debugging.
from logging.handlers import RotatingFileHandler # For managing log files to prevent them from growing too large.

# --- NEWLY ADDED IMPORTS FOR ADVANCED SENSORS AND CALCULATIONS ---
from collections import deque       # A double-ended queue, perfect for creating a moving average filter for sensor data.
try:                                # NumPy is a powerful library for numerical operations.
    import numpy as np              # Used here for polynomial fitting in the RUL prediction algorithm.
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False
    print("[!] WARNING: NumPy not found (pip install numpy). RUL prediction accuracy will be reduced.")

# --- Third-Party Library Imports ---
# These libraries must be installed via pip (e.g., pip install wmi psutil beautifulsoup4 PyQt5)

# Suppress ignorable warnings to keep the console output clean.
warnings.filterwarnings('ignore')

# Attempt to import WMI, a critical library for Windows system information.
try:
    # wmi provides a high-level interface to Windows Management Instrumentation.
    import wmi
    # pythoncom is needed to initialize the COM library for multi-threaded WMI access.
    import pythoncom
    # Set a flag indicating WMI is available.
    WMI_AVAILABLE = True
    # Log success for debugging purposes during startup.
    print("[✓] WMI module loaded successfully.")
# Handle the case where the wmi library is not installed.
except ImportError:
    # Set the flag to False.
    WMI_AVAILABLE = False
    # Print a critical error message as WMI is essential for core functionality.
    print("[X] CRITICAL ERROR: WMI module not found. Please run 'pip install wmi'.")
    # Exit the application because it cannot function without WMI.
    sys.exit(1)

# Attempt to import psutil, a cross-platform process and system utilities library.
try:
    # psutil is used for a quick check of battery presence and status.
    import psutil
    # Set a flag indicating psutil is available.
    PSUTIL_AVAILABLE = True
    # Log success for debugging purposes.
    print("[✓] psutil module loaded successfully.")
# Handle the case where psutil is not installed.
except ImportError:
    # Set the flag to False.
    PSUTIL_AVAILABLE = False
    # Print an error message. The app can degrade gracefully but will be less effective.
    print("[X] ERROR: psutil module not found. Some features will be limited. Please run 'pip install psutil'.")

# Attempt to import BeautifulSoup, a library for parsing HTML and XML files.
try:
    # BeautifulSoup is used to parse the HTML battery report generated by powercfg.
    from bs4 import BeautifulSoup
    # Set a flag indicating BeautifulSoup is available.
    BS4_AVAILABLE = True
    # Log success for debugging purposes.
    print("[✓] BeautifulSoup module loaded successfully.")
# Handle the case where BeautifulSoup is not installed.
except ImportError:
    # Set the flag to False.
    BS4_AVAILABLE = False
    # Print a warning. Parsing powercfg reports will be limited, reducing data accuracy.
    print("[!] WARNING: BeautifulSoup not found. PowerCfg parsing will be limited. Please run 'pip install beautifulsoup4'.")

# Check if ctypes is available (it's a standard library but good to confirm).
try:
    # Re-import specific components from ctypes for clarity.
    from ctypes import wintypes, Structure, byref, windll, c_ulong, c_wchar_p, POINTER
    # Set a flag indicating ctypes is available.
    CTYPES_AVAILABLE = True
    # Log success.
    print("[✓] ctypes module loaded successfully.")
# This should rarely fail, but handle it just in case of a broken Python installation.
except ImportError:
    # Set the flag to False.
    CTYPES_AVAILABLE = False
    # Print a warning as some native Windows API calls will fail.
    print("[!] WARNING: ctypes module not available. Some native API calls will fail.")

# Attempt to import PyQt5, the GUI framework for the application.
try:
    # Import all necessary UI components from PyQt5.QtWidgets.
    from PyQt5.QtWidgets import (
        QApplication, QMainWindow, QWidget, QFrame, QLabel, QPushButton,
        QVBoxLayout, QHBoxLayout, QGridLayout, QScrollArea, QSystemTrayIcon,
        QGraphicsDropShadowEffect, QGraphicsBlurEffect, QSizePolicy, QSplashScreen, QMessageBox,
        QProgressBar, QMenuBar, QMenu, QAction, QDialog, QInputDialog, QSpacerItem
    )
    # Import all necessary graphics components from PyQt5.QtGui.
    from PyQt5.QtGui import (
        QPainter, QColor, QPen, QBrush, QRadialGradient, QFont,
        QLinearGradient, QPainterPath, QPixmap, QIcon, QConicalGradient
    )
    # Import all necessary core components from PyQt5.QtCore.
    from PyQt5.QtCore import (
        Qt, QTimer, QSize, QPoint, QPointF, QRect, QRectF,
        QPropertyAnimation, pyqtProperty, QEasingCurve, QThread,
        pyqtSignal, QObject, QByteArray, QSequentialAnimationGroup, QParallelAnimationGroup
    )
    # Import the SVG renderer for displaying the logo.
    from PyQt5.QtSvg import QSvgRenderer
    # Set a flag indicating PyQt5 is available.
    PYQT5_AVAILABLE = True
    # Log success.
    print("[✓] PyQt5 GUI framework loaded successfully.")
# Handle the case where PyQt5 is not installed.
except ImportError as e:
    # Print a critical, user-friendly error message with installation instructions.
    print(f"\n[X] CRITICAL ERROR: PyQt5 is not installed or failed to load: {e}")
    print("This is a graphical application and requires PyQt5 to run.")
    print("Please install it by running: pip install PyQt5==5.15.10")
    # Exit the application as the UI cannot be created.
    sys.exit(1)


# ============================================================================
# SECTION 2: GLOBAL CONSTANTS AND CONFIGURATIONS
# Description: This section defines all global variables, constants, and
#              configuration parameters for the application. This centralizes
#              control over the app's behavior, appearance, and internal logic.
# ============================================================================

# --- Application Metadata ---
APP_VERSION = "2.0.0"                                                                   # The current version of the application.
APP_NAME = "Battery-Z"                                                                  # The official name of the application.
APP_SUBTITLE = "Your Battery's Best Friend"                                             # The slogan displayed in the UI.
BUILD_DATE = "2025-10-10"                                                               # The build date for this version.
AUTHOR_NAME = "Muhammad Kashan Tariq"                                                   # The name of the developer.
AUTHOR_GITHUB = "https://github.com/Mr-Muhammad-Kashan"                                 # Developer's GitHub profile link.
AUTHOR_LINKEDIN = "https://www.linkedin.com/in/muhammad-kashan-tariq"                  # Developer's LinkedIn profile link.
BUY_ME_COFFEE = "https://mr-muhammad-kashan.github.io/Buy-Me-A-Coffee-Website/"         # Link for users to support the developer.

# --- Asset Paths ---
LOGO_ICO_PATH = "logo.ico"                                  # Path to the icon file for the window and system tray.
LOGO_SVG_PATH = "logo.svg"                                  # Path to the scalable vector graphics logo for the UI.

# --- Data Mappings and Dictionaries ---

# This dictionary maps numeric chemistry codes (from WMI's Win32_Battery class) to human-readable names.
# This is critical for fixing the issue where the UI showed numbers instead of names.
CHEMISTRY_CODES = {
    1: "Other",
    2: "Unknown",
    3: "Lead Acid",
    4: "Nickel Cadmium (NiCd)",
    5: "Nickel Metal Hydride (NiMH)",
    6: "Lithium-ion (Li-ion)",
    7: "Zinc Air",
    8: "Lithium Polymer (Li-Po)"
}

# This dictionary provides a comprehensive mapping of various string representations of battery chemistries
# to a standardized, human-readable format. This is used as a fallback to normalize data from
# different sources (e.g., powercfg reports, registry) which may use different naming conventions.
CHEMISTRY_NAMES = {
    "li-ion": "Lithium-ion",
    "lion": "Lithium-ion",
    "lithium ion": "Lithium-ion",
    "lithium-ion": "Lithium-ion",
    "li ion": "Lithium-ion",
    "liion": "Lithium-ion",
    "ncm": "Lithium-ion NCM",
    "nca": "Lithium-ion NCA",
    "li-poly": "Lithium Polymer",
    "lipo": "Lithium Polymer",
    "lithium polymer": "Lithium Polymer",
    "lithium-polymer": "Lithium Polymer",
    "li-po": "Lithium Polymer",
    "lipolymer": "Lithium Polymer",
    "li-ion polymer": "Lithium-ion Polymer",
    "lithium ion polymer": "Lithium-ion Polymer",
    "lfp": "Lithium Iron Phosphate",
    "lifepo4": "Lithium Iron Phosphate",
    "lithium iron phosphate": "Lithium Iron Phosphate",
    "nimh": "Nickel Metal Hydride",
    "ni-mh": "Nickel Metal Hydride",
    "nickel metal hydride": "Nickel Metal Hydride",
    "nicd": "Nickel Cadmium",
    "ni-cd": "Nickel Cadmium",
    "nickel cadmium": "Nickel Cadmium",
    "lead acid": "Lead Acid",
    "pb": "Lead Acid"
}

# --- Health and Notification Thresholds ---
EOL_HEALTH_THRESHOLD = 0.80             # Health percentage (80%) below which RUL prediction is calculated.
CRITICAL_HEALTH_THRESHOLD = 0.60        # Health percentage (60%) considered critical, may trigger stronger recommendations.
HIGH_TEMP_THRESHOLD = 55.0              # Temperature in Celsius (55°C) above which a high-temperature alert is triggered.
LOW_BATTERY_THRESHOLD = 20              # Battery percentage (20%) below which a low battery notification is sent.

# This list of tuples defines the mapping from a health percentage range to a descriptive status and a color code for the UI.
# Format: (min_percentage, max_percentage, "Status String", "#HexColorCode")
HEALTH_STATUS_MAP = [
    (95, 100, "Excellent", "#00ff00"),  # 95-100% is Excellent (Green)
    (85, 94, "Very Good", "#7cfc00"),   # 85-94% is Very Good (Lime Green)
    (75, 84, "Good", "#9a67ea"),        # 75-84% is Good (Purple - as per original UI design)
    (60, 74, "Fair", "#ffa500"),        # 60-74% is Fair (Orange)
    (40, 59, "Moderate", "#ff8c00"),    # 40-59% is Moderate (Dark Orange)
    (0, 39, "Poor", "#ff0000")          # 0-39% is Poor (Red)
]

# --- Manufacturer and Chemistry Specific Data ---

# ============================================================================
# MANUFACTURER DATABASE
# This comprehensive database contains typical battery characteristics for a wide
# range of laptop manufacturers and models. It is the foundation for providing
# accurate, context-aware predictions.
# Data points are based on public specifications, reviews, and typical values for the product tier.
# Structure: { manufacturer_key: { "models": { model_keyword: {details} }, "default": {details} } }
# ============================================================================
MANUFACTURER_DATABASE = {
    "dell": {
        "models": {
            "xps": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "latitude": {"cycles": 1200, "quality": "Premium", "chem": "Li-Po"},
            "precision": {"cycles": 1200, "quality": "Premium", "chem": "Li-Po"},
            "alienware": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"},
            "inspiron": {"cycles": 800, "quality": "Budget", "chem": "Li-ion"},
            "vostro": {"cycles": 900, "quality": "Mid-range", "chem": "Li-ion"},
            "g series": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "hp": {
        "models": {
            "spectre": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "envy": {"cycles": 900, "quality": "Premium", "chem": "Li-Po"},
            "elitebook": {"cycles": 1200, "quality": "Premium", "chem": "Li-Po"},
            "probook": {"cycles": 1000, "quality": "Mid-range", "chem": "Li-ion"},
            "zbook": {"cycles": 1200, "quality": "Premium", "chem": "Li-Po"},
            "omen": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"},
            "pavilion": {"cycles": 700, "quality": "Budget", "chem": "Li-ion"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "lenovo": {
        "models": {
            "thinkpad": {"cycles": 1200, "quality": "Premium", "chem": "Li-Po"},
            "yoga": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "legion": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"},
            "ideapad": {"cycles": 700, "quality": "Budget", "chem": "Li-ion"},
            "thinkbook": {"cycles": 1000, "quality": "Mid-range", "chem": "Li-Po"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "asus": {
        "models": {
            "zenbook": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "rog": {"cycles": 800, "quality": "Premium", "chem": "Li-ion"},
            "proart": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "tuf": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"},
            "vivobook": {"cycles": 700, "quality": "Budget", "chem": "Li-ion"},
            "expertbook": {"cycles": 1100, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "apple": {
        "models": {
            "macbook pro": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "macbook air": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"}
    },
    "microsoft": {
        "models": {
            "surface book": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "surface laptop": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
            "surface pro": {"cycles": 900, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 500, "quality": "Premium", "chem": "Li-Po"}
    },
    "acer": {
        "models": {
            "swift": {"cycles": 900, "quality": "Mid-range", "chem": "Li-ion"},
            "predator": {"cycles": 700, "quality": "Mid-range", "chem": "Li-ion"},
            "nitro": {"cycles": 700, "quality": "Budget", "chem": "Li-ion"},
            "aspire": {"cycles": 600, "quality": "Budget", "chem": "Li-ion"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "razer": {
        "models": {
            "blade": {"cycles": 800, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 500, "quality": "Premium", "chem": "Li-Po"}
    },
    "msi": {
        "models": {
            "ge": {"cycles": 800, "quality": "Premium", "chem": "Li-ion"},
            "gs": {"cycles": 800, "quality": "Premium", "chem": "Li-Po"},
            "gt": {"cycles": 800, "quality": "Premium", "chem": "Li-ion"},
        },
        "default": {"cycles": 500, "quality": "Mid-range", "chem": "Li-ion"}
    },
    "samsung": {
        "models": {
            "galaxy book": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 600, "quality": "Mid-range", "chem": "Li-Po"}
    },
    "lg": {
        "models": {
            "gram": {"cycles": 1000, "quality": "Premium", "chem": "Li-Po"},
        },
        "default": {"cycles": 500, "quality": "Premium", "chem": "Li-Po"}
    },
    "framework": {
        "models": {},
        "default": {"cycles": 1000, "quality": "Premium", "chem": "Li-ion"}
    },
    # Generic fallback for any other manufacturers.
    "generic": {
        "models": {},
        "default": {"cycles": 800, "quality": "Mid-range", "chem": "Li-ion"}
    }
}

# ============================================================================
# A database of default rated cycle counts based on battery chemistry. This serves as a fallback
# if the manufacturer and model are not found in the database above.
# VALUES UPDATED BASED ON RESEARCH DATA FOR HIGHER ACCURACY.
DEFAULT_RATED_CYCLES = {
    # Mainstream Chemistries
    "Lithium-ion": 1000,              # Generic baseline
    "Lithium Polymer": 800,           # Lower cycle life in some configs
    "Lithium-ion Polymer": 1000,      # Often a hybrid
    
    # Specific Cathode Chemistries
    "Lithium Cobalt Oxide": 750,      # (LCO) - Common in older/thinner laptops
    "Lithium-ion NCM": 1500,          # (NMC) - Very common, balanced performance
    "Lithium-ion NCA": 800,           # (NCA) - High energy, slightly lower cycles
    "Lithium Iron Phosphate": 3000,   # (LFP) - Very high cycle life, rare in laptops
    "Lithium Manganese Oxide": 500,   # (LMO) - Less common, lower cycle life

    # Older Chemistries
    "Nickel Metal Hydride": 500,      # (NiMH)
    "Nickel Cadmium": 1500,           # (NiCd)
    "Lead Acid": 300,

    # Fallbacks
    "Unknown": 800, # A sensible default for modern unknown chemistries.
    "Other": 1000
}

# --- UI and Application Behavior Configuration ---
BASE_WINDOW_WIDTH = 1600                # The base width of the main window at 96 DPI.
BASE_WINDOW_HEIGHT = 950                # The base height of the main window at 96 DPI.
BASE_DPI = 96.0                         # The baseline DPI for UI scaling calculations.
REALTIME_POLL_INTERVAL = 3000           # The interval in milliseconds (3 seconds) for polling real-time sensor data.
POWERCFG_TIMEOUT = 30                   # The timeout in seconds for the powercfg command to prevent hangs.

# --- Global State Variables ---
# This global variable allows the user to manually override the detected cycle count for testing or calibration.
# It is modified via the UI and checked during the data analysis phase.
CUSTOM_CYCLE_COUNT = None

# A list of welcome quotes. A random one is chosen on each application startup.
WELCOME_QUOTES = [
    "Have a great day! Let's check your battery health.",
    "Welcome back! Time to monitor your power.",
    "Your battery's best friend is here!",
    "Let's keep your laptop running longer.",
    "Smart monitoring for smart users.",
    "Powering your productivity, one electron at a time.",
    "Your battery guardian is active.",
    "Maximize your battery life with intelligence.",
    "Battery health secured. Performance maximized.",
    "Real-time intelligence. Predictive accuracy.",
    "Every cycle counts. Every charge matters.",
    "From kernel to dashboard: Total battery visibility.",
    "Military-grade precision for your mobile power.",
    "Battery analytics at your fingertips.",
    "Intelligent power management starts here."
]

# A list of battery care and optimization tips displayed in the UI.
BATTERY_TIPS = [
    "Keep your battery between 20-80% for maximum lifespan - avoid extremes.",
    "Avoid extreme temperatures; your battery is happiest between 15-25°C.",
    "Perform a full charge cycle (5% to 100%) once every 1-3 months to calibrate the sensor.",
    "High temperatures are the #1 killer of batteries. Ensure your laptop has good ventilation.",
    "Using the original charger is always best for your battery's health and safety.",
    "Lowering your screen brightness is one of the easiest ways to extend battery runtime.",
    "Disable Wi-Fi and Bluetooth when you don't need them to conserve power.",
    "Don't leave your laptop plugged in at 100% all the time. Use manufacturer charge-limiting tools if available.",
    "Close background applications you aren't using to reduce CPU load and power draw.",
    "Storing an unused laptop? Keep it in a cool place with about 40-50% charge.",
    "Update your system's BIOS/firmware; updates often include better power management.",
    "Fast charging generates more heat. Use standard charging when you're not in a hurry.",
    "Replace your battery when its health drops below 70-80% for the best performance and safety."
]


# ============================================================================
# SECTION 3: PLATFORM VALIDATION AND INITIALIZATION
# Description: This section contains code that runs at the very start to
#              validate the execution environment and set up initial state.
# ============================================================================

# Define a function to validate that the application is running on Windows.
def validate_platform() -> bool:
    """
    Checks if the current operating system is Windows. If not, it prints an
    error and exits. If it is Windows, it prints system details.

    Returns:
        bool: True if the platform is Windows, otherwise the program exits.
    """
    # Check if the platform system is 'Windows'.
    is_windows = platform.system() == "Windows"
    # If it is Windows, proceed to print details.
    if is_windows:
        # Get the Windows version (e.g., '10', '11').
        win_version = platform.release()
        # Print a success message with the detected OS version.
        print(f"[✓] Platform: Windows {win_version}")
        # Print the Python version being used.
        print(f"[✓] Python: {sys.version.split()[0]}")
        # Print the system architecture (e.g., 'AMD64').
        print(f"[✓] Architecture: {platform.machine()}")
    # If the OS is not Windows, print an error and terminate.
    else:
        # Print a critical error message.
        print("[X] CRITICAL ERROR: This application is designed for Windows OS only.")
        # Exit the application with a non-zero status code to indicate an error.
        sys.exit(1)
    # Return the boolean result.
    return is_windows

# Execute the platform validation immediately upon script load.
# The result is stored in a global constant for easy access elsewhere in the code.
IS_WINDOWS = validate_platform()


# ============================================================================
# SECTION 4: DATA STRUCTURE DEFINITIONS
# Description: This section defines the core data structures used throughout the
#              application, such as the BatteryData dataclass and ctypes
#              structures for Windows API interaction.
# ============================================================================

# This dataclass serves as the central data container for all battery-related information.
# It aggregates data from all sources into a single, structured object.
@dataclass
class BatteryData:
    """
    A dataclass to hold all static and dynamic battery information gathered
    from various system sources. It provides a unified and consistent data model
    for the rest of the application to use.
    """
    # --- System Information ---
    laptop_manufacturer: Optional[str] = None       # e.g., "Dell Inc."
    laptop_model: Optional[str] = None              # e.g., "XPS 15 9520"

    # --- Battery Identification ---
    battery_present: bool = False                   # Flag indicating if a battery is detected in the system.
    battery_name: Optional[str] = None              # The model name of the battery, e.g., "DELL 71R31C7".
    battery_manufacturer: Optional[str] = None      # The manufacturer of the battery cell, e.g., "SANYO".
    battery_serial: Optional[str] = None            # The unique serial number of the battery.
    battery_chemistry: Optional[str] = None         # The human-readable chemistry, e.g., "Lithium-ion".
    battery_chemistry_code: Optional[int] = None    # The numeric chemistry code from WMI.

    # --- Capacity Metrics ---
    design_capacity_mwh: Optional[int] = None       # The original rated capacity of the battery in milliwatt-hours.
    full_charge_capacity_mwh: Optional[int] = None  # The current maximum capacity the battery can hold in mWh.
    current_capacity_mwh: Optional[int] = None      # The current charge level in mWh.

    # --- Voltage and Power Metrics ---
    design_voltage_mv: Optional[int] = None         # The rated design voltage in millivolts.
    current_voltage_mv: Optional[int] = None        # The current voltage in millivolts.
    power_draw_watts: Optional[float] = None        # The current power draw or charge rate in Watts.
    charge_rate_mw: Optional[int] = None            # The charge rate in milliwatts (if charging).
    discharge_rate_mw: Optional[int] = None         # The discharge rate in milliwatts (if discharging).

    # --- Usage and Health Metrics ---
    cycle_count: Optional[int] = None               # The number of charge/discharge cycles the battery has undergone.
    rated_cycle_life: int = 1000                    # The manufacturer's rated cycle life, looked up or defaulted.
    temperature_celsius: Optional[float] = None     # The current battery temperature in degrees Celsius.

    # --- Real-time Status ---
    current_percentage: Optional[int] = None        # The current charge percentage (0-100).
    is_charging: Optional[bool] = None              # True if the battery is currently charging.
    ac_online: Optional[bool] = None                # True if the AC adapter is plugged in.
    time_to_empty_seconds: Optional[int] = None     # Estimated time until the battery is empty, in seconds.
    time_to_full_seconds: Optional[int] = None      # Estimated time until the battery is fully charged, in seconds.
    
    # --- Metadata and Flags ---
    fetch_timestamp: Optional[datetime.datetime] = None # Timestamp of when the data was last fetched.
    fetch_errors: List[str] = field(default_factory=list) # A list of any errors encountered during data fetching.
    high_temp_notified: bool = False                # Flag to prevent repeated high-temperature notifications.
    low_battery_notified: bool = False              # Flag to prevent repeated low-battery notifications.

    # This method provides a quick check to see if the core capacity data required for health calculation is present and valid.
    def has_valid_capacity_data(self) -> bool:
        """
        Checks if the dataclass contains valid and logical design and full charge
        capacity values, which are essential for any health calculation.

        Returns:
            bool: True if capacity data is valid, otherwise False.
        """
        # Return True only if both design and full charge capacities are not None, and the design capacity is a plausible value (e.g., > 1000 mWh).
        return (self.design_capacity_mwh is not None and
                self.full_charge_capacity_mwh is not None and
                self.design_capacity_mwh > 1000)

# This ctypes structure maps to the Windows SYSTEM_POWER_STATUS struct, allowing direct calls to the Kernel32 GetSystemPowerStatus API.
# This is a highly reliable and fast method for getting basic, real-time battery status.
if CTYPES_AVAILABLE and IS_WINDOWS:
    class SYSTEM_POWER_STATUS(Structure):
        """
        A ctypes structure that mirrors the native Windows SYSTEM_POWER_STATUS struct.
        This allows Python to directly interact with low-level Windows APIs for power information.
        """
        # Define the fields of the structure, matching the Windows API documentation.
        _fields_ = [
            ('ACLineStatus',       ctypes.c_byte),  # 0=Offline, 1=Online, 255=Unknown
            ('BatteryFlag',        ctypes.c_byte),  # Charging status flags
            ('BatteryLifePercent', ctypes.c_byte),  # Current percentage (0-100, 255=Unknown)
            ('SystemStatusFlag',   ctypes.c_byte),  # Battery saver status
            ('BatteryLifeTime',    wintypes.DWORD), # Remaining seconds of battery life
            ('BatteryFullLifeTime',wintypes.DWORD)  # Not typically used
        ]


# ============================================================================
# SECTION 5: UTILITY AND HELPER FUNCTIONS
# Description: This section contains general-purpose helper functions used
#              throughout the application for tasks like data conversion,
#              formatting, and safe value handling.
# ============================================================================

def get_user_display_name() -> str:
    """
    Retrieves the current user's display name using multiple fallback methods
    for maximum reliability across different Windows environments.

    Returns:
        str: The user's name, or "User" as a fallback.
    """
    # Method 1: Environment variable (fastest and most common).
    try:
        # Get the 'USERNAME' environment variable.
        username = os.environ.get('USERNAME')
        # If a valid username is found, print and return it.
        if username and len(username) > 0:
            print(f"[✓] PC Username (Environment): {username}")
            return username
    except Exception as e:
        # Log if this method fails.
        print(f"[!] Environment username method failed: {e}")

    # Method 2: Windows API call (very reliable).
    if CTYPES_AVAILABLE:
        try:
            # Get a handle to the GetUserNameW function from advapi32.dll.
            GetUserNameW = windll.advapi32.GetUserNameW
            # Define the function's argument types for ctypes.
            GetUserNameW.argtypes = [wintypes.LPWSTR, POINTER(wintypes.DWORD)]
            # Define the function's return type.
            GetUserNameW.restype = wintypes.BOOL
            
            # Create a buffer to store the username.
            size = wintypes.DWORD(257)
            name_buffer = ctypes.create_unicode_buffer(size.value)
            # Call the function.
            if GetUserNameW(name_buffer, byref(size)):
                # If successful, get the value from the buffer.
                username = name_buffer.value
                # If valid, print and return it.
                if username and len(username) > 0:
                    print(f"[✓] PC Username (Windows API): {username}")
                    return username
        except Exception as e:
            # Log if this method fails.
            print(f"[!] GetUserNameW API call failed: {e}")
    
    # Method 3: getpass module (standard library fallback).
    try:
        # Use the getpass library to get the user name.
        username = getpass.getuser()
        # If valid, print and return it.
        if username and len(username) > 0:
            print(f"[✓] PC Username (getpass): {username}")
            return username
    except Exception as e:
        # Log if this method fails.
        print(f"[!] getpass method failed: {e}")
    
    # Method 4: WMI query (robust but slower).
    if WMI_AVAILABLE:
        try:
            # Initialize COM for the current thread.
            pythoncom.CoInitialize()
            # Create a WMI connection.
            c = wmi.WMI()
            # Query the Win32_ComputerSystem class.
            for cs in c.Win32_ComputerSystem():
                # Check if the UserName attribute exists.
                if hasattr(cs, 'UserName') and cs.UserName:
                    # The username might be in "DOMAIN\user" format, so split and take the last part.
                    username = cs.UserName.split('\\')[-1]
                    print(f"[✓] PC Username (WMI): {username}")
                    # Uninitialize COM before returning.
                    pythoncom.CoUninitialize()
                    return username
            # Uninitialize COM if no user was found.
            pythoncom.CoUninitialize()
        except Exception as e:
            # Log if WMI fails.
            print(f"[!] WMI username method failed: {e}")
            # Ensure COM is uninitialized even on failure.
            try:
                pythoncom.CoUninitialize()
            except:
                pass
    
    # Final fallback if all other methods fail.
    return "User"

# This function selects and returns a random quote from the global list.
def get_random_quote() -> str:
    """Selects a random welcome quote from the global list."""
    # Use random.choice to pick one element from the WELCOME_QUOTES list.
    return random.choice(WELCOME_QUOTES)

# This function safely converts any value to an integer, returning a default value on failure.
def safe_int(value: Any, default: int = 0) -> int:
    """
    Safely converts a value to an integer. Handles None, strings, and floats.

    Args:
        value (Any): The input value to convert.
        default (int): The value to return if conversion fails. Defaults to 0.

    Returns:
        int: The converted integer or the default value.
    """
    # Use a try-except block to handle potential conversion errors.
    try:
        # If the value is None, return the default immediately.
        if value is None:
            return default
        # Convert the value first to a float (to handle "123.0") and then to an int.
        return int(float(value))
    # If a ValueError or TypeError occurs during conversion.
    except (ValueError, TypeError):
        # Return the specified default value.
        return default

# This function safely converts any value to a float, returning a default value on failure.
def safe_float(value: Any, default: float = 0.0) -> float:
    """
    Safely converts a value to a float. Handles None and strings.

    Args:
        value (Any): The input value to convert.
        default (float): The value to return if conversion fails. Defaults to 0.0.

    Returns:
        float: The converted float or the default value.
    """
    # Use a try-except block to handle potential conversion errors.
    try:
        # If the value is None, return the default immediately.
        if value is None:
            return default
        # Convert the value to a float.
        return float(value)
    # If a ValueError or TypeError occurs during conversion.
    except (ValueError, TypeError):
        # Return the specified default value.
        return default

# This function formats a large integer with commas for better readability in the UI.
def format_with_commas(value: int) -> str:
    """Formats an integer with thousand separators (commas)."""
    # Use an f-string with the ':' format specifier to add commas.
    return f"{value:,}"

# This function converts a duration in seconds into a human-readable "Xh Ym" format.
def format_time_duration(seconds: int) -> str:
    """
    Converts a duration in seconds to a human-readable string like 'Xh Ym'.
    Handles special cases for 'Calculating...' and AC power.
    """
    # Windows API sometimes returns -1 (0xFFFFFFFF) when calculating, handle this special case.
    if seconds < 0 or seconds == 0xFFFFFFFF:
        return "Calculating..."
    # If on AC power, seconds might be 0.
    if seconds == 0:
        return "On A/C Power"
    
    # Use divmod to get hours and the remaining seconds in one operation.
    hours, remainder = divmod(seconds, 3600)
    # Calculate minutes from the remainder.
    minutes = remainder // 60
    
    # Handle cases where the duration is longer than a day.
    if hours > 24:
        # Calculate days and remaining hours.
        days = hours // 24
        hours = hours % 24
        # Return a format including days.
        return f"{days}d {hours}h {minutes}m"
    
    # Return the standard "Xh Ym" format.
    return f"{hours}h {minutes}m"

# This function calculates battery health based on design and full charge capacity.
def calculate_health_percentage(design: float, full_charge: float) -> float:
    """
    Calculates the battery health percentage based on its current full charge
    capacity relative to its original design capacity.

    Args:
        design (float): The design capacity in mWh.
        full_charge (float): The current full charge capacity in mWh.

    Returns:
        float: The health percentage, clamped between 0.0 and 100.0.
    """
    # Check for invalid or zero values to prevent division by zero errors.
    if not design or not full_charge or design <= 0:
        return 0.0
    # Calculate the health ratio.
    health = (full_charge / design) * 100.0
    # Clamp the result between 0 and 100 to handle cases where full_charge might exceed design.
    return min(max(health, 0.0), 100.0)

# This function returns a status string and color based on the health percentage.
def get_health_status(health_pct: float) -> Tuple[str, str]:
    """
    Maps a health percentage to a descriptive status string and a UI color code.

    Args:
        health_pct (float): The battery health percentage.

    Returns:
        Tuple[str, str]: A tuple containing the status label (e.g., "Good")
                        and the hex color code (e.g., "#00ff00").
    """
    # Iterate through the global HEALTH_STATUS_MAP.
    for min_val, max_val, label, color in HEALTH_STATUS_MAP:
        # Check if the health percentage falls within the current range.
        if min_val <= health_pct <= max_val:
            # Return the corresponding label and color.
            return label, color
    # Return a default "Unknown" status if no range matches.
    return "Unknown", "#888888"

# This function normalizes various chemistry strings into a standard format.
def normalize_chemistry_name(chemistry_str: Union[str, int]) -> str:
    """
    Normalizes a battery chemistry value (which can be a string, a number string,
    or a code) into a standardized, human-readable format. This is a critical
    function for ensuring consistent data presentation.

    Args:
        chemistry_str (Union[str, int]): The raw chemistry value from a system API.

    Returns:
        str: The standardized chemistry name (e.g., "Lithium-ion").
    """
    # Handle null or empty input.
    if not chemistry_str:
        return "Unknown"
    
    # Convert input to a string and clean it up.
    chem_str_cleaned = str(chemistry_str).lower().strip()
    
    # --- Method 1: Check for numeric code ---
    # If the string is a digit, it's likely a WMI code.
    if chem_str_cleaned.isdigit():
        # Convert to an integer.
        code = int(chem_str_cleaned)
        # Look up the code in the CHEMISTRY_CODES dictionary.
        # .get() provides a default value of "Unknown" if the code is not found.
        return CHEMISTRY_CODES.get(code, "Unknown")
        
    # --- Method 2: Check for known string variations ---
    # Iterate through the CHEMISTRY_NAMES dictionary for common abbreviations and names.
    for key, value in CHEMISTRY_NAMES.items():
        # If a known key is found within the cleaned string.
        if key in chem_str_cleaned:
            # Return the standardized value.
            return value
    
    # --- Method 3: Heuristic-based detection for less common strings ---
    # These checks catch variations that aren't in the main dictionary.
    if "li" in chem_str_cleaned and "ion" in chem_str_cleaned:
        if "poly" in chem_str_cleaned:
            return "Lithium-ion Polymer"
        return "Lithium-ion"
    if "li" in chem_str_cleaned and "poly" in chem_str_cleaned:
        return "Lithium Polymer"
    if "nickel" in chem_str_cleaned:
        if "metal" in chem_str_cleaned or "mh" in chem_str_cleaned:
            return "Nickel Metal Hydride"
        if "cadmium" in chem_str_cleaned or "cd" in chem_str_cleaned:
            return "Nickel Cadmium"
            
    # --- Final Fallback ---
    # If no match is found, return the original string, capitalized for presentation.
    # Return "Unknown" if the original string was empty after cleaning.
    return str(chemistry_str).strip().title() or "Unknown"

# This function retrieves the manufacturer's rated cycle life from a database.
def get_manufacturer_rated_cycles(manufacturer: str, model: str, chemistry: str) -> int:
    """
    [ENHANCED] Looks up the rated cycle life for a battery using the new,
    comprehensive manufacturer database with a multi-tiered fallback system.

    Args:
        manufacturer (str): The laptop manufacturer (e.g., "Dell Inc.").
        model (str): The laptop model (e.g., "XPS 15 9520").
        chemistry (str): The normalized battery chemistry.

    Returns:
        int: The most accurate rated cycle life found in the database.
    """
    # Sanitize inputs for case-insensitive matching.
    mfr_lower = (manufacturer or "").lower()
    model_lower = (model or "").lower()
    
    # --- Step 1: Search the new Manufacturer Database ---
    matched_mfr_key = "generic" # Start with the most generic fallback.
    for mfr_key in MANUFACTURER_DATABASE:
        if mfr_key in mfr_lower:
            matched_mfr_key = mfr_key
            break # Found the manufacturer, stop searching.
            
    # Get the data for the matched manufacturer.
    mfr_data = MANUFACTURER_DATABASE[matched_mfr_key]
    
    # --- Step 2: Search for a specific model keyword within that manufacturer ---
    for model_key, model_details in mfr_data["models"].items():
        if model_key in model_lower:
            logging.info(f"Database match found for {manufacturer} {model_key}. Using {model_details['cycles']} cycles.")
            return model_details['cycles']
            
    # --- Step 3: If no model matches, use the manufacturer's default ---
    if matched_mfr_key != "generic":
        logging.info(f"No specific model matched for {model}. Using default for {manufacturer}: {mfr_data['default']['cycles']} cycles.")
        return mfr_data['default']['cycles']
        
    # --- Step 4: If manufacturer is unknown, fall back to chemistry-based defaults ---
    chem_cycles = DEFAULT_RATED_CYCLES.get(chemistry, 1000)
    logging.info(f"No manufacturer match. Falling back to chemistry-based default for {chemistry}: {chem_cycles} cycles.")
    return chem_cycles

# This function attempts to get the Windows installation date.
def get_windows_install_date() -> Optional[datetime.datetime]:
    """
    Retrieves the Windows installation date from the Registry, with a fallback
    to parsing the 'systeminfo' command output. This is used to estimate the
    age of the system and, by proxy, the battery.

    Returns:
        Optional[datetime.datetime]: A datetime object of the install date, or None if not found.
    """
    # Method 1: Windows Registry (fast and reliable).
    try:
        # Open the required registry key. HKEY_LOCAL_MACHINE stores system-wide settings.
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        # Query the "InstallDate" value, which is a UNIX timestamp.
        install_date_value, _ = winreg.QueryValueEx(key, "InstallDate")
        # Close the registry key to free resources.
        winreg.CloseKey(key)
        # Convert the timestamp to a standard datetime object.
        install_date = datetime.datetime.fromtimestamp(install_date_value)
        # Return the successful result.
        return install_date
    except Exception:
        # If this fails, silently pass and try the next method.
        pass
    
    # Method 2: systeminfo command (slower but good fallback).
    try:
        # Run the 'systeminfo' command.
        result = subprocess.run(
            ["systeminfo"],
            capture_output=True, # Capture the command's stdout.
            text=True,           # Decode the output as text.
            timeout=10,          # Set a timeout to prevent hangs.
            creationflags=subprocess.CREATE_NO_WINDOW # Hide the console window.
        )
        # Iterate through each line of the output.
        for line in result.stdout.split('\n'):
            # Look for the line containing the install date.
            if "Original Install Date" in line:
                # Extract the date string after the colon.
                date_str = line.split(':', 1)[1].strip()
                # Parse the specific date format used by systeminfo.
                install_date = datetime.datetime.strptime(date_str, "%m/%d/%Y, %I:%M:%S %p")
                # Return the successful result.
                return install_date
    except Exception:
        # If this also fails, silently pass.
        pass
        
    # Return None if all methods failed.
    return None

# ============================================================================
# PART 2
# ============================================================================

# ============================================================================
# SECTION 6: CORE LOGIC - BATTERY INTELLIGENCE CLASS
# Description: This is the central class for all backend logic. It encapsulates
#              data fetching, caching, and health/RUL calculations. This class
#              replaces the multiple logic classes from the original main.py
#              (WindowsBatteryDataFetcher, BatteryHealthCalculator, RULPredictor)
#              with a single, more robust, and maintainable implementation
#              inspired by reference.py.
# ============================================================================

class BatteryIntelligence:
    """
    This class handles all backend operations: fetching battery data from
    multiple Windows APIs, caching the results for performance, parsing the data,
    and performing health and remaining life calculations. It is designed to be
    the single source of truth for battery status.
    """
    
    # The __init__ method is the constructor for the class.
    def __init__(self):
        """
        Initializes the BatteryIntelligence class by setting up paths for data
        persistence, loading cached data, and preparing for data collection.
        """
        # --- Path Configuration for Data Persistence ---
        # Get the user's AppData/Roaming directory path. Storing data here is the correct
        # practice for Windows applications, as it's a user-specific, non-roaming location.
        self.appdata_path = os.path.join(os.getenv('APPDATA'), 'BatteryZ_Data')
        # Create the directory if it doesn't exist to prevent errors when writing files.
        os.makedirs(self.appdata_path, exist_ok=True)
        
        # Define the full paths for all persistent data files.
        self.cache_file = os.path.join(self.appdata_path, 'battery_cache.json')
        self.report_path = os.path.join(self.appdata_path, 'battery_report.xml')
        
        # --- Logging Setup ---
        # Configure logging to write to a file within our AppData folder.
        # This provides a persistent record of operations and errors for debugging.
        log_file = os.path.join(self.appdata_path, 'battery_z.log')
        # Use a RotatingFileHandler to automatically manage log file size, preventing it from
        # growing indefinitely. It will keep 3 backup logs, each up to 5MB.
        log_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=3)
        # Set the logging level to INFO, and define a clear format for log messages.
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[log_handler]
        )
        
        # --- Cache Initialization ---
        # Load the cache from disk. This is a dictionary holding semi-static battery
        # information to speed up subsequent app launches.
        self.cache = self.load_cache()
        # Log the result of the cache loading operation.
        logging.info("BatteryIntelligence initialized. Cache loaded with %d items.", len(self.cache))

    # --- Caching Methods ---
    
    def load_cache(self) -> Dict:
        """
        Loads the battery data cache from a JSON file. The cache is considered
        valid for 24 hours to balance performance with data freshness.

        Returns:
            Dict: The loaded cache dictionary, or an empty dictionary if the
                cache is invalid, missing, or corrupt.
        """
        # Check if the cache file actually exists on disk.
        if os.path.exists(self.cache_file):
            # Use a try-except block to handle potential file reading or JSON parsing errors.
            try:
                # Open the cache file for reading.
                with open(self.cache_file, 'r') as f:
                    # Parse the JSON content into a Python dictionary.
                    cache = json.load(f)
                    # Define cache validity period in seconds (24 hours).
                    cache_validity_seconds = 86400
                    # Check if the cache is still fresh by comparing its timestamp to the current time.
                    if time.time() - cache.get("last_updated", 0) < cache_validity_seconds:
                        # If the cache is fresh, log it and return the data.
                        logging.info("Valid cache found. Loading data from cache.")
                        return cache
                    else:
                        # If the cache is stale, log it and it will be overwritten.
                        logging.info("Cache is stale. A new cache will be created.")
            # If any error occurs (e.g., corrupted JSON), log the error.
            except Exception as e:
                logging.error("Failed to load or validate cache: %s", e)
        # If the file doesn't exist or loading failed, return an empty dictionary.
        return {}

    def save_cache(self):
        """
        Saves the current state of the self.cache dictionary to the JSON file.
        It also injects a 'last_updated' timestamp.
        """
        # Use a try-except block to handle potential file writing errors.
        try:
            # Add/update the timestamp to mark when this cache was saved.
            self.cache["last_updated"] = time.time()
            # Open the cache file for writing.
            with open(self.cache_file, 'w') as f:
                # Dump the cache dictionary to the file with an indent for readability.
                json.dump(self.cache, f, indent=4)
            # Log the successful save operation.
            logging.info("Cache saved successfully to %s", self.cache_file)
        # If an error occurs (e.g., disk full, permissions error), log it.
        except Exception as e:
            logging.error("Failed to save cache: %s", e)
            
    # --- Primary Data Orchestration Method ---
    
    def get_all_data(self) -> BatteryData:
        """
        This is the main orchestration method. It calls all individual data
        fetching functions, populates the BatteryData object, performs calculations,
        and returns the final, consolidated result.

        Returns:
            BatteryData: A fully populated dataclass with all available battery info.
        """
        # Create a new, empty BatteryData object for this fetch cycle.
        data = BatteryData()
        # Record the start time of the fetch operation.
        data.fetch_timestamp = datetime.datetime.now()
        
        # --- Step 1: Check for Battery Presence ---
        # Use psutil as a fast, primary check for whether a battery exists.
        if PSUTIL_AVAILABLE:
            try:
                # If psutil.sensors_battery() returns None, no battery is detected.
                if psutil.sensors_battery() is None:
                    data.battery_present = False
                    logging.warning("No battery detected via psutil. Assuming desktop PC.")
                else:
                    data.battery_present = True
            except Exception as e:
                logging.error("psutil check failed: %s. Assuming battery is present as a fallback.", e)
                # If psutil fails, assume a battery is present and let other methods confirm.
                data.battery_present = True
        else:
            # If psutil isn't available, we must assume a battery might be present.
            data.battery_present = True
            
        # If no battery is present, we can stop early.
        if not data.battery_present:
            # Populate basic system info even for desktops.
            data.laptop_manufacturer, data.laptop_model = self._get_system_info()
            return data
            
        # --- Step 2: Fetch and Consolidate Data from All Sources ---
        # This is where the multi-layered fallback strategy happens.
        
        # Fetch basic system info first.
        data.laptop_manufacturer, data.laptop_model = self._get_system_info()
        
        # Generate the powercfg report if needed. This is a slow operation, so we
        # do it once and then parse data from it in multiple functions.
        self._generate_battery_report()
        
        # Fetch static data (things that don't change often, like serial number, design capacity).
        static_info = self._get_static_battery_info()
        data.battery_name = static_info.get("name")
        data.battery_manufacturer = static_info.get("manufacturer")
        data.battery_serial = static_info.get("serial")
        data.design_capacity_mwh = static_info.get("design_capacity")
        data.full_charge_capacity_mwh = static_info.get("full_charge_capacity")
        
        # Fetch cycle count using its own multi-fallback logic.
        data.cycle_count = self._get_cycle_count()
        
        # Fetch chemistry using its own multi-fallback logic.
        raw_chem = self._get_chemistry()
        # Normalize the fetched chemistry name for consistency.
        data.battery_chemistry = normalize_chemistry_name(raw_chem)
        
        # Look up rated cycles based on the fetched data.
        data.rated_cycle_life = get_manufacturer_rated_cycles(
            data.laptop_manufacturer, 
            data.laptop_model, 
            data.battery_chemistry
        )
        
        # --- Step 3: Fetch Real-time Dynamic Data ---
        # This data changes frequently (charge level, voltage, etc.).
        dynamic_info = self._get_dynamic_battery_status()
        data.current_percentage = dynamic_info.get("percent")
        data.ac_online = dynamic_info.get("ac_online")
        data.is_charging = dynamic_info.get("is_charging")
        data.time_to_empty_seconds = dynamic_info.get("time_remaining")
        data.current_voltage_mv = dynamic_info.get("voltage_mv")
        data.power_draw_watts = dynamic_info.get("power_watts")
        data.temperature_celsius = self._get_temperature() # Temperature has its own fallback chain.
        
        # --- Step 4: Final Calculations and Data Cleanup ---
        
        # If a custom cycle count is set by the user, override the fetched value.
        if CUSTOM_CYCLE_COUNT is not None:
            data.cycle_count = CUSTOM_CYCLE_COUNT
            
        # Save the consolidated static data to cache for the next run.
        self.cache.update({
            "manufacturer": data.battery_manufacturer,
            "serial_number": data.battery_serial,
            "chemistry": data.battery_chemistry,
            "design_capacity": data.design_capacity_mwh,
            "full_charge_capacity": data.full_charge_capacity_mwh,
            "cycle_count": data.cycle_count,
            "total_cycles": data.rated_cycle_life,
            "battery_name": data.battery_name
        })
        self.save_cache()
        
        # Log a summary of the key fetched values for debugging.
        logging.info(
            "Data fetch complete. Cycles: %s, Design: %s mWh, FCC: %s mWh, Chem: %s",
            data.cycle_count, data.design_capacity_mwh, data.full_charge_capacity_mwh, data.battery_chemistry
        )
        
        # Return the final, populated data object.
        return data

    # --- Calculation Methods (from reference.py) ---

    def calculate_health(self, data: BatteryData) -> float:
        """
        [ALGORITHM FIXED] Calculates battery health using a research-based,
        multi-factor model. The bug that incorrectly capped the health value,
        ignoring the cycle count's impact, has been removed.

        Args:
            data (BatteryData): The populated battery data object.

        Returns:
            float: The calculated health percentage (0-100).
        """
        if not data.has_valid_capacity_data():
            logging.warning("Cannot calculate health due to invalid capacity data.")
            return 0.0

        # --- Step 1: Gather Inputs and Validate ---
        design_cap = float(data.design_capacity_mwh)
        full_cap = float(data.full_charge_capacity_mwh)
        cycle_count = float(data.cycle_count if data.cycle_count is not None and data.cycle_count >= 0 else 0)
        total_cycles = float(data.rated_cycle_life if data.rated_cycle_life > 0 else 1000)
        
        # --- Step 2: Calculate Primary Health Indicators (HIs) ---
        # HI 1: Capacity Health (SOH_c) - The direct physical measurement.
        soh_c = (full_cap / design_cap) * 100.0
        
        # HI 2: Cycle Health (SOH_cyc) - The wear based on usage.
        cycle_ratio = min(cycle_count / total_cycles, 1.5) # Allow ratio to go beyond 1.0 for old batteries
        soh_cyc = 100.0 * (1 - (0.2 * (cycle_ratio ** 1.5))) # Non-linear degradation model
        
        # --- Step 3: Fuse the Health Indicators using dynamic weighting ---
        if cycle_ratio < 0.1: # Battery is new
            capacity_weight = 0.4
            cycle_weight = 0.6
        elif cycle_ratio > 0.8: # Battery is old
            capacity_weight = 0.8
            cycle_weight = 0.2
        else: # For mid-life batteries
            capacity_weight = 0.7
            cycle_weight = 0.3
            
        final_health = (soh_c * capacity_weight) + (soh_cyc * cycle_weight)
        
        # --- Step 4: Final Clamping and Validation ---
        # BUG FIX: The line 'final_health = min(final_health, soh_c)' has been REMOVED.
        # This line was the root cause of the previous issue, as it was incorrectly
        # overriding the weighted calculation and preventing the cycle count from
        # having its full, intended effect on the final health score.
        final_health = max(0.0, min(final_health, 100.0))
        
        logging.info(
            "Health calculated. SOH_c: %.1f%%, SOH_cyc: %.1f%%. Final Weighted SOH: %.1f%%",
            soh_c, soh_cyc, final_health
        )
        
        return final_health

# CODE REPLACEMENT 3
# ============================================================================
    def estimate_remaining_life(self, data: BatteryData) -> Dict[str, Union[int, str]]:
        """
        [ALGORITHM REPLACEMENT] Estimates the remaining useful life (RUL) by
        projecting future health degradation based on historical usage patterns
        until it crosses the 80% end-of-life threshold. This is far more
        accurate than simple linear extrapolation.

        Args:
            data (BatteryData): The populated battery data object.

        Returns:
            Dict: A dictionary containing 'years', 'months', 'days', and a 'status' string.
        """
        # --- Default return value and validation ---
        default_rul = {"years": 0, "months": 0, "days": 0, "status": "N/A"}
        
        if (not data.has_valid_capacity_data() or
            data.cycle_count is None or data.cycle_count < 0 or
            data.rated_cycle_life <= 0):
            logging.warning("Cannot estimate RUL due to insufficient data for projection.")
            return default_rul

        # --- Step 1: Calculate historical usage rate ---
        install_date = get_windows_install_date()
        if install_date and (datetime.datetime.now() - install_date).days > 0:
            age_days = (datetime.datetime.now() - install_date).days
        else:
            # Fallback: estimate age from cycle count assuming 0.7 cycles/day for a typical user.
            age_days = int(data.cycle_count / 0.7) if data.cycle_count > 0 else 1
        
        age_days = max(age_days, 1) # Avoid division by zero
        cycles_per_day = data.cycle_count / age_days
        
        if cycles_per_day <= 0.01: # Handle very low or zero usage
            logging.warning("Usage rate is too low to make a reliable RUL projection.")
            return {"years": 10, "months": 0, "days": 0, "status": "Low Usage"}

        # --- Step 2: Project future degradation day-by-day ---
        # The industry standard for battery end-of-life is 80% of original capacity.
        REPLACEMENT_THRESHOLD_SOH = 80.0
        
        current_soh = self.calculate_health(data)
        if current_soh < REPLACEMENT_THRESHOLD_SOH:
            return {"years": 0, "months": 0, "days": 0, "status": "Replace Now"}
            
        days_to_eol = -1
        # Project forward for a maximum of 15 years (5475 days) to prevent infinite loops.
        for day in range(1, 5475):
            # Project the future cycle count based on historical usage rate.
            projected_cycles = data.cycle_count + (day * cycles_per_day)
            
            # Use the same non-linear cycle health formula from calculate_health to project SOH.
            cycle_ratio = projected_cycles / data.rated_cycle_life
            projected_soh = 100.0 * (1 - (0.2 * (cycle_ratio ** 1.5)))

            # Check if the projected health has crossed the replacement threshold.
            if projected_soh < REPLACEMENT_THRESHOLD_SOH:
                days_to_eol = day
                break
        
        if days_to_eol == -1:
            logging.info("RUL projection exceeds 15 years. Capping result.")
            return {"years": 15, "months": 0, "days": 0, "status": "Excellent"}

        # --- Step 3: Convert remaining days to Years, Months, Days ---
        years = days_to_eol // 365
        months = (days_to_eol % 365) // 30
        days = (days_to_eol % 365) % 30

        logging.info(
            "RUL Estimated. Usage: %.2f cycles/day. Days to 80%% SOH: %d -> %d years, %d months",
            cycles_per_day, days_to_eol, years, months
        )
        
        return {"years": years, "months": months, "days": days, "status": "Calculated"}
    
    def _generate_battery_report(self):
        """
        Generates a powercfg battery report in XML format. It only generates a new
        report if the existing one is missing or stale (older than 1 hour). This
        is a key performance optimization.
        """
        # Check if a fresh report already exists.
        if os.path.exists(self.report_path) and time.time() - os.path.getmtime(self.report_path) < 3600:
            logging.info("Recent powercfg report found. Skipping generation.")
            return

        # If no fresh report, generate a new one.
        logging.info("Generating new powercfg battery report...")
        try:
            # Use subprocess to run the powercfg command. The /xml flag is more machine-readable than /html.
            subprocess.run(
                ["powercfg", "/batteryreport", "/xml", "/output", self.report_path],
                check=True,         # Raise an exception if the command fails.
                capture_output=True,# Suppress output from appearing in the console.
                timeout=POWERCFG_TIMEOUT # Prevent the app from hanging.
            )
            # Add a small delay to ensure the file is fully written to disk before we try to read it.
            time.sleep(1)
            logging.info("Powercfg report generated successfully at %s", self.report_path)
        except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as e:
            # If the command fails, log a detailed error.
            logging.error("Failed to generate powercfg battery report: %s", e)
            # Try to delete a potentially corrupted or incomplete report file.
            if os.path.exists(self.report_path):
                try:
                    os.remove(self.report_path)
                except OSError as E:
                    logging.error("Could not remove corrupted report file: %s" ,E)
    
    def _get_system_info(self) -> Tuple[str, str]:
        """
        Retrieves the system manufacturer and model using WMI. This provides
        context for the battery data.
        
        Returns:
            Tuple[str, str]: A tuple containing (manufacturer, model).
        """
        # Default values if WMI fails.
        manufacturer, model = "Unknown", "System"
        
        # Check if WMI is available.
        if not WMI_AVAILABLE:
            return manufacturer, model
            
        # Use a try-except block for the WMI query.
        try:
            # Initialize COM for the current thread.
            pythoncom.CoInitialize()
            # Create a WMI connection.
            c = wmi.WMI()
            # Query the Win32_ComputerSystemProduct class for system info.
            system_info = c.Win32_ComputerSystemProduct()[0]
            # Get the vendor (manufacturer) and name (model).
            manufacturer = system_info.Vendor.strip()
            model = system_info.Name.strip()
        except Exception as e:
            # Log any errors that occur.
            logging.error("Failed to get system info via WMI: %s", e)
        finally:
            # CRITICAL: Always uninitialize COM in a multi-threaded app.
            pythoncom.CoUninitialize()
            
        # Return the retrieved or default values.
        return manufacturer, model

# ============================================================================
# PART 3
# ============================================================================

    def _get_static_battery_info(self) -> Dict:
        """
        Retrieves static battery information (data that doesn't change frequently)
        using a multi-layered fallback strategy to ensure maximum accuracy and
        availability. It prioritizes the most reliable sources first.

        Returns:
            Dict: A dictionary containing static battery info like 'design_capacity',
                  'full_charge_capacity', 'manufacturer', 'serial', and 'name'.
        """
        # Initialize a dictionary to hold the results.
        info = {}
        
        # --- Fallback Strategy ---
        # The goal is to fill the 'info' dictionary. We try methods in order of
        # reliability. Once a value is found, we don't overwrite it with a
        # value from a less reliable source.

        # --- Method 1: Parse the Powercfg XML Report ---
        # This is often the most comprehensive and reliable source.
        try:
            # Check if the report file exists.
            if os.path.exists(self.report_path):
                # Open and read the entire file content.
                with open(self.report_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # Use regex to find and extract key values from the XML structure.
                design_match = re.search(r'<DesignCapacity>(\d+)</DesignCapacity>', content)
                full_match = re.search(r'<FullChargeCapacity>(\d+)</FullChargeCapacity>', content)
                manuf_match = re.search(r'<Manufacturer>(.*?)</Manufacturer>', content)
                serial_match = re.search(r'<SerialNumber>(.*?)</SerialNumber>', content)
                name_match = re.search(r'<Name>(.*?)</Name>', content)
                
                # For each matched value, convert it to the correct type and add to our info dict.
                if design_match:
                    info['design_capacity'] = safe_int(design_match.group(1))
                if full_match:
                    info['full_charge_capacity'] = safe_int(full_match.group(1))
                if manuf_match and manuf_match.group(1).strip():
                    info['manufacturer'] = manuf_match.group(1).strip()
                if serial_match and serial_match.group(1).strip():
                    info['serial'] = serial_match.group(1).strip()
                if name_match and name_match.group(1).strip():
                    info['name'] = name_match.group(1).strip()
                
                logging.info("Successfully parsed data from powercfg XML report.")
        except Exception as e:
            logging.error("Failed to parse powercfg XML report: %s", e)

        # --- Method 2: WMI (ROOT\WMI and root\cimv2) ---
        # WMI is a powerful native Windows API that often provides direct hardware access.
        if WMI_AVAILABLE:
            try:
                # Initialize COM for this thread.
                pythoncom.CoInitialize()
                # Connect to the advanced 'root\wmi' namespace.
                c_wmi = wmi.WMI(namespace="root\\wmi")
                # Connect to the standard 'root\cimv2' namespace.
                c_cimv2 = wmi.WMI()

                # Get static data like serial, manufacturer from BatteryStaticData class.
                static_data_list = c_wmi.BatteryStaticData()
                if static_data_list:
                    static_data = static_data_list[0]
                    if 'manufacturer' not in info and hasattr(static_data, 'ManufactureName') and static_data.ManufactureName.strip():
                        info['manufacturer'] = static_data.ManufactureName.strip('\x00').strip()
                    if 'serial' not in info and hasattr(static_data, 'SerialNumber') and static_data.SerialNumber.strip():
                        info['serial'] = static_data.SerialNumber.strip('\x00').strip()
                    if 'design_capacity' not in info and hasattr(static_data, 'DesignedCapacity') and static_data.DesignedCapacity > 0:
                        info['design_capacity'] = static_data.DesignedCapacity
                
                # Get full charge capacity from BatteryFullChargedCapacity class.
                fcc_data_list = c_wmi.BatteryFullChargedCapacity()
                if fcc_data_list and 'full_charge_capacity' not in info:
                    if hasattr(fcc_data_list[0], 'FullChargedCapacity') and fcc_data_list[0].FullChargedCapacity > 0:
                        info['full_charge_capacity'] = fcc_data_list[0].FullChargedCapacity

                # Get battery name/model from Win32_Battery class as another fallback.
                battery_data_list = c_cimv2.Win32_Battery()
                if battery_data_list and 'name' not in info:
                    if hasattr(battery_data_list[0], 'DeviceID') and battery_data_list[0].DeviceID.strip():
                         info['name'] = battery_data_list[0].DeviceID.strip()

            except Exception as e:
                logging.error("Failed to get static info via WMI: %s", e)
            finally:
                # Always ensure COM is uninitialized.
                pythoncom.CoUninitialize()

        # --- Final Sanity Checks and Fallbacks ---
        # If after all methods, some data is still missing, use defaults or derive them.
        if 'design_capacity' not in info or info['design_capacity'] <= 0:
            info['design_capacity'] = 50000 # Default to 50Wh
            logging.warning("Design capacity not found. Using default value: %d mWh", info['design_capacity'])
        
        if 'full_charge_capacity' not in info or info['full_charge_capacity'] <= 0:
            # As a last resort, estimate FCC as 90% of design capacity.
            info['full_charge_capacity'] = int(info['design_capacity'] * 0.9)
            logging.warning("Full charge capacity not found. Estimating based on design capacity: %d mWh", info['full_charge_capacity'])
            
        return info

    def _get_cycle_count(self) -> Optional[int]:
        """
        Retrieves the battery cycle count using an extensive chain of fallback
        methods to ensure the highest possible accuracy. It tries direct hardware
        queries first, then report parsing, and finally estimation.

        Returns:
            Optional[int]: The detected cycle count, or None if all methods fail.
        """
        logging.info("Attempting to fetch battery cycle count...")
        
        # --- Method 1: WMI (ROOT\WMI - BatteryCycleCount) ---
        # This is often the most direct and accurate hardware query.
        if WMI_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                c = wmi.WMI(namespace="root\\wmi")
                cycle_data = c.BatteryCycleCount()
                if cycle_data and hasattr(cycle_data[0], 'CycleCount'):
                    count = cycle_data[0].CycleCount
                    logging.info("SUCCESS: Cycle count from WMI (root\\wmi) is %d.", count)
                    pythoncom.CoUninitialize()
                    return int(count)
            except Exception as e:
                logging.warning("WMI (root\\wmi) for cycle count failed: %s. Trying next method.", e)
            finally:
                pythoncom.CoUninitialize()
        
        # --- Method 2: Powercfg XML Report ---
        # If WMI fails, the next best source is the generated report.
        try:
            if os.path.exists(self.report_path):
                with open(self.report_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                match = re.search(r'<CycleCount>(\d+)</CycleCount>', content)
                if match:
                    count = int(match.group(1))
                    logging.info("SUCCESS: Cycle count from powercfg XML is %d.", count)
                    return count
        except Exception as e:
            logging.warning("Parsing powercfg XML for cycle count failed: %s. Trying next method.", e)

        # --- Method 3: PowerShell (as a subprocess) ---
        # This can sometimes succeed where the Python WMI library fails.
        try:
            # This command queries the same WMI class as Method 1, but via PowerShell.
            command = "Get-CimInstance -Namespace ROOT\\WMI -ClassName BatteryCycleCount | Select-Object -ExpandProperty CycleCount"
            result = subprocess.run(
                ["powershell", "-Command", command],
                capture_output=True, text=True, timeout=10, check=True
            )
            output = result.stdout.strip()
            if output.isdigit():
                count = int(output)
                logging.info("SUCCESS: Cycle count from PowerShell is %d.", count)
                return count
        except Exception as e:
            logging.warning("PowerShell for cycle count failed: %s. Trying next method.", e)

        # --- Method 4: Estimation from Capacity Degradation (Last Resort) ---
        # This is an estimation, not a direct reading, and is only used if all other methods fail.
        logging.warning("All direct methods for cycle count failed. Falling back to estimation.")
        try:
            design = self.cache.get('design_capacity')
            fcc = self.cache.get('full_charge_capacity')
            total_cycles = self.cache.get('total_cycles', 1000)
            
            if design and fcc and design > 0 and fcc > 0:
                # The logic from reference.py: assume 20% total wear over the battery's lifespan.
                # We can reverse this to estimate how many cycles correspond to the current wear level.
                wear_level = 1.0 - (fcc / design)
                if wear_level > 0:
                    # (wear_level / 0.20) gives the fraction of lifespan used. Multiply by total cycles.
                    estimated_count = int((wear_level / 0.20) * total_cycles)
                    logging.info("SUCCESS: Estimated cycle count from capacity degradation is %d.", estimated_count)
                    return estimated_count
        except Exception as e:
            logging.error("Cycle count estimation failed: %s", e)

        # If all methods fail, return None.
        logging.error("CRITICAL: Could not determine cycle count from any available method.")
        return None

    def _get_chemistry(self) -> Optional[Union[str, int]]:
        """
        Retrieves the battery chemistry using multiple fallback methods. It prioritizes
        string-based names but will also fetch numeric codes. The result should
        be passed to `normalize_chemistry_name`.

        Returns:
            Optional[Union[str, int]]: The raw chemistry data, or None.
        """
        logging.info("Attempting to fetch battery chemistry...")
        
        # --- Method 1: WMI (ROOT\WMI - BatteryStaticData) ---
        # Often provides a descriptive string like "LION".
        if WMI_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                c = wmi.WMI(namespace="root\\wmi")
                static_data = c.BatteryStaticData()
                if static_data and hasattr(static_data[0], 'Chemistry') and static_data[0].Chemistry:
                    chem = static_data[0].Chemistry.strip('\x00').strip()
                    if chem:
                        logging.info("SUCCESS: Chemistry from WMI (root\\wmi) is '%s'.", chem)
                        pythoncom.CoUninitialize()
                        return chem
            except Exception as e:
                logging.warning("WMI (root\\wmi) for chemistry failed: %s.", e)
            finally:
                pythoncom.CoUninitialize()

        # --- Method 2: Powercfg XML Report ---
        try:
            if os.path.exists(self.report_path):
                with open(self.report_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                match = re.search(r'<Chemistry>(.*?)</Chemistry>', content)
                if match and match.group(1).strip():
                    chem = match.group(1).strip()
                    logging.info("SUCCESS: Chemistry from powercfg XML is '%s'.", chem)
                    return chem
        except Exception as e:
            logging.warning("Parsing powercfg XML for chemistry failed: %s.", e)

        # --- Method 3: WMI (root\cimv2 - Win32_Battery) ---
        # This usually returns a numeric code.
        if WMI_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                c = wmi.WMI()
                battery_data = c.Win32_Battery()
                if battery_data and hasattr(battery_data[0], 'Chemistry'):
                    code = battery_data[0].Chemistry
                    logging.info("SUCCESS: Chemistry code from WMI (cimv2) is %d.", code)
                    pythoncom.CoUninitialize()
                    return code
            except Exception as e:
                logging.warning("WMI (cimv2) for chemistry failed: %s.", e)
            finally:
                pythoncom.CoUninitialize()

        # If all methods fail, return None.
        logging.error("CRITICAL: Could not determine chemistry from any available method.")
        return None

    def _get_dynamic_battery_status(self) -> Dict:
        """
        Retrieves real-time, dynamic battery status information (charge, status,
        time remaining, etc.) using the fastest and most reliable methods.

        Returns:
            Dict: A dictionary containing the dynamic status.
        """
        # Initialize a dictionary to hold the results.
        status = {}
        
        # --- Method 1: Ctypes call to GetSystemPowerStatus ---
        # This is the fastest and most direct way to get real-time status.
        if CTYPES_AVAILABLE:
            try:
                # Create an instance of our ctypes structure.
                power_status = SYSTEM_POWER_STATUS()
                # Call the kernel32 function, passing the structure by reference.
                if windll.kernel32.GetSystemPowerStatus(byref(power_status)):
                    # Extract the values from the structure.
                    if power_status.ACLineStatus != 255:
                        status['ac_online'] = (power_status.ACLineStatus == 1)
                    if power_status.BatteryLifePercent != 255:
                        status['percent'] = power_status.BatteryLifePercent
                    if power_status.BatteryLifeTime != 0xFFFFFFFF:
                        status['time_remaining'] = power_status.BatteryLifeTime
                    
                    # Infer charging status from AC line status and percentage.
                    if status.get('ac_online') and status.get('percent', 100) < 100:
                        status['is_charging'] = True
                    elif status.get('ac_online') is False:
                        status['is_charging'] = False
                    else:
                        # Could be fully charged or unknown.
                        status['is_charging'] = False
            except Exception as e:
                logging.warning("GetSystemPowerStatus API call failed: %s. Trying fallbacks.", e)
        
        # --- Method 2: psutil (Excellent Fallback) ---
        # If ctypes fails, or to fill in missing pieces, use psutil.
        if PSUTIL_AVAILABLE:
            try:
                # Get the battery sensor data from psutil.
                battery = psutil.sensors_battery()
                if battery:
                    # Fill in any values not already found by the ctypes method.
                    if 'percent' not in status:
                        status['percent'] = int(battery.percent)
                    if 'ac_online' not in status:
                        status['ac_online'] = battery.power_plugged
                    if 'is_charging' not in status:
                        # psutil's power_plugged is a good proxy for charging.
                        status['is_charging'] = battery.power_plugged and battery.percent < 100
                    if 'time_remaining' not in status and battery.secsleft != psutil.POWER_TIME_UNLIMITED:
                        status['time_remaining'] = battery.secsleft
            except Exception as e:
                logging.warning("psutil.sensors_battery failed: %s. Trying WMI.", e)
                
        # --- Method 3: WMI (Last Resort for dynamic data) ---
        # WMI can also provide this, but is generally slower than the other two methods.
        if WMI_AVAILABLE and ('percent' not in status or 'voltage_mv' not in status):
            try:
                pythoncom.CoInitialize()
                c_wmi = wmi.WMI(namespace="root\\wmi")
                c_cimv2 = wmi.WMI()
                
                # Get voltage and power draw from root\wmi.
                wmi_status_list = c_wmi.BatteryStatus()
                if wmi_status_list:
                    wmi_status = wmi_status_list[0]
                    if hasattr(wmi_status, 'Voltage') and wmi_status.Voltage > 0:
                        status['voltage_mv'] = wmi_status.Voltage
                    # DischargeRate is in milliwatts, negative for discharge, positive for charge.
                    if hasattr(wmi_status, 'DischargeRate') and wmi_status.DischargeRate != 0:
                        status['power_watts'] = abs(wmi_status.DischargeRate) / 1000.0

                # Get percentage from root\cimv2 as a final fallback.
                if 'percent' not in status:
                    wmi_battery = c_cimv2.Win32_Battery()
                    if wmi_battery and hasattr(wmi_battery[0], 'EstimatedChargeRemaining'):
                        status['percent'] = wmi_battery[0].EstimatedChargeRemaining

            except Exception as e:
                logging.error("Final WMI fallback for dynamic status failed: %s", e)
            finally:
                pythoncom.CoUninitialize()

        # Return the consolidated status dictionary.
        return status
        
    def _get_temperature(self) -> Optional[float]:
        """
        Retrieves the battery temperature using WMI, as it's the most common
        source for this data on Windows.

        Returns:
            Optional[float]: Temperature in Celsius, or None if not available.
        """
        # This data is almost exclusively available via this WMI class.
        if not WMI_AVAILABLE:
            return None
            
        try:
            pythoncom.CoInitialize()
            c = wmi.WMI(namespace="root\\wmi")
            # Query the MSAcpi_ThermalZoneTemperature class. Often, one of the zones corresponds to the battery.
            temp_zones = c.MSAcpi_ThermalZoneTemperature()
            if temp_zones:
                # The temperature is given in tenths of a Kelvin.
                temp_kelvin = temp_zones[0].CurrentTemperature / 10.0
                # Convert from Kelvin to Celsius.
                temp_celsius = temp_kelvin - 273.15
                
                # Perform a sanity check on the value.
                if -20 < temp_celsius < 120:
                    logging.info("SUCCESS: Temperature from WMI is %.1f °C.", temp_celsius)
                    pythoncom.CoUninitialize()
                    return round(temp_celsius, 1)
        except Exception as e:
            # It's common for this query to fail if the hardware doesn't expose temperature data.
            logging.warning("Could not retrieve temperature from WMI: %s", e)
        finally:
            pythoncom.CoUninitialize()
            
        # Return None if not found.
        return None
# ============================================================================
# PART 4
# ============================================================================

# ============================================================================
# SECTION 7: QT WORKER CLASSES FOR BACKGROUND PROCESSING
# Description: This section defines the QObject-based worker classes that run
#              on separate threads. This is a critical design pattern for any
#              responsive GUI application. It offloads long-running tasks
#              (like initial data fetching and continuous polling) from the main
#              UI thread, preventing the application from freezing.
# ============================================================================

class RealtimeUpdateWorker(QObject):
    """
    This worker is responsible for polling for dynamic, real-time battery data
    at a regular interval. It runs on a separate thread and emits signals to
    the main UI thread with fresh data, without blocking the user interface.
    """
    # --- Signals ---
    # Define a PyQt signal that will be emitted when new data is available.
    # It will carry a dictionary object containing the updated sensor readings.
    data_updated = pyqtSignal(dict)
    # Define a signal for high-temperature alerts, carrying the temperature value.
    high_temp_alert = pyqtSignal(float)
    # Define a signal for low-battery alerts, carrying the percentage value.
    low_battery_alert = pyqtSignal(int)
    
    # The __init__ method is the constructor for the class.
    def __init__(self, intelligence_instance: BatteryIntelligence):
        """
        Initializes the RealtimeUpdateWorker.

        Args:
            intelligence_instance (BatteryIntelligence): An instance of the main
                backend class to use for fetching data.
        """
        # Call the constructor of the parent QObject class.
        super().__init__()
        # Store the instance of the BatteryIntelligence class.
        self.intelligence = intelligence_instance
        # A flag to control the execution loop of the worker.
        self.running = True
        # Flags to prevent spamming notifications for the same event.
        self.high_temp_notified = False
        self.low_battery_notified = False

    # This method stops the worker's execution loop.
    def stop(self):
        """Stops the worker's polling loop safely."""
        # Set the running flag to False. The loop in poll_realtime_data will terminate.
        self.running = False
        # Log that the worker has been requested to stop.
        logging.info("RealtimeUpdateWorker stop requested.")

    # This is the main workhorse method that runs in a loop on the background thread.
    def poll_realtime_data(self):
        """
        Continuously polls for real-time battery data in a loop until stopped.
        It fetches dynamic status, checks for alert conditions, and emits signals.
        """
        # Log that the worker's polling loop has started.
        logging.info("Realtime polling worker started on a new thread.")
        # Loop indefinitely as long as the 'running' flag is True.
        while self.running:
            # Use a try-except block to catch any errors during a poll cycle.
            try:
                # Call the backend method to get the latest dynamic battery status.
                realtime_data = self.intelligence._get_dynamic_battery_status()
                # Also fetch the latest temperature reading.
                temperature = self.intelligence._get_temperature()
                # Add the temperature to the data dictionary.
                realtime_data['temperature_celsius'] = temperature
                
                # --- Alerting Logic ---
                # Check for high temperature.
                if temperature is not None and temperature > HIGH_TEMP_THRESHOLD:
                    # If temperature is high and we haven't notified yet...
                    if not self.high_temp_notified:
                        # Emit the high temperature signal.
                        self.high_temp_alert.emit(temperature)
                        # Set the flag to True to prevent re-notifying immediately.
                        self.high_temp_notified = True
                else:
                    # Reset the flag once the temperature is back to normal.
                    self.high_temp_notified = False
                
                # Check for low battery.
                current_percent = realtime_data.get('percent')
                is_charging = realtime_data.get('is_charging', True) # Default to True to prevent false alerts
                if current_percent is not None and current_percent < LOW_BATTERY_THRESHOLD and not is_charging:
                    # If battery is low, not charging, and we haven't notified...
                    if not self.low_battery_notified:
                        # Emit the low battery signal.
                        self.low_battery_alert.emit(current_percent)
                        # Set the flag to True.
                        self.low_battery_notified = True
                else:
                    # Reset the flag once battery is charged or plugged in.
                    self.low_battery_notified = False
                
                # --- Data Emission ---
                # Emit the signal with the newly fetched data dictionary.
                self.data_updated.emit(realtime_data)
                
            # Catch any exception during the poll cycle to prevent the thread from crashing.
            except Exception as e:
                # Log the error for debugging.
                logging.error("An error occurred in the real-time polling loop: %s", e)
            
            # --- Sleep ---
            # Wait for the specified interval before the next poll.
            # We check the 'running' flag frequently within the sleep period
            # to allow for a faster shutdown if stop() is called.
            for _ in range(int(REALTIME_POLL_INTERVAL / 100)):
                if not self.running:
                    break
                time.sleep(0.1)
        
        # Log that the worker's loop has terminated.
        logging.info("Realtime polling worker has stopped.")

class BatteryDataWorker(QObject):
    """
    This worker handles the heavy, one-time data acquisition process that runs
    at application startup. It runs on a separate thread and emits progress
    signals to the splash screen, ensuring the UI remains responsive during
    the potentially long initial load.
    """
    # --- Signals ---
    # A signal to update the splash screen's progress bar and status message.
    progress = pyqtSignal(str, int)
    # A signal emitted when all data has been successfully fetched, carrying the final BatteryData object.
    finished = pyqtSignal(object)
    # A signal emitted if an unrecoverable error occurs during data fetching.
    error = pyqtSignal(str)
    
    # The __init__ method is the constructor for the class.
    def __init__(self):
        """Initializes the BatteryDataWorker."""
        # Call the parent constructor.
        super().__init__()
        # Create an instance of our backend logic class. This ensures the worker
        # has its own instance to use on its own thread.
        self.intelligence = BatteryIntelligence()

    # This is the main method called to start the data fetching process.
    def fetch_data(self):
        """
        Orchestrates the entire initial data fetching process, emitting
        progress signals along the way.
        """
        # Use a try-except block to catch any fatal errors during the entire process.
        try:
            # Emit progress signals to the splash screen at various stages.
            # This provides visual feedback to the user that the application is working.
            self.progress.emit("Initializing...", 10)
            time.sleep(0.1) # A small delay to make the splash screen feel smoother.
            
            self.progress.emit("Detecting system hardware...", 20)
            time.sleep(0.1)
            
            self.progress.emit("Querying Windows Management Instrumentation...", 40)
            time.sleep(0.1)

            # The most time-consuming step is often generating the powercfg report.
            self.progress.emit("Generating system battery report...", 60)
            # The get_all_data method will handle the actual report generation.
            
            self.progress.emit("Analyzing battery data...", 80)
            
            # Call the main orchestration method from our backend class.
            # This single call performs all the complex fetching and consolidation logic.
            battery_data = self.intelligence.get_all_data()
            
            # Once complete, emit the final progress update.
            self.progress.emit("Finalizing...", 95)
            time.sleep(0.2)
            
            self.progress.emit("Complete!", 100)
            
            # Emit the 'finished' signal, passing the complete BatteryData object to the main thread.
            self.finished.emit(battery_data)
            
        # If any exception occurs during the process...
        except Exception as e:
            # Format a user-friendly error message.
            error_msg = f"A critical error occurred during data acquisition: {str(e)}"
            # Log the full error and traceback for debugging.
            logging.error(error_msg)
            logging.error(traceback.format_exc())
            # Emit the 'error' signal to notify the main thread of the failure.
            self.error.emit(error_msg)


# ============================================================================
# SECTION 8: UI HELPER AND MANAGER CLASSES
# Description: This section contains classes that manage UI-related aspects like
#              DPI scaling, styling, system tray integration, and notifications.
#              These classes are preserved from the original main.py to maintain
#              the UI's appearance and behavior.
# ============================================================================

class DisplayManager:
    """
    Manages display properties like DPI and screen size to ensure the UI
    scales correctly on high-resolution displays.
    """
    def __init__(self, screen):
        self.screen = screen
        self.geometry = screen.availableGeometry()
        self.dpi = screen.logicalDotsPerInch()
        # Calculate a scale factor based on the system DPI relative to the base DPI.
        self.scale_factor = max(1.0, self.dpi / BASE_DPI)
        
        print(f"[✓] Display Manager Initialized: {self.geometry.width()}x{self.geometry.height()} @ {self.dpi} DPI (Scale: {self.scale_factor:.2f}x)")

    def get_optimal_window_size(self) -> QSize:
        """Calculates the ideal window size based on the screen size and scale factor."""
        scaled_width = int(BASE_WINDOW_WIDTH * self.scale_factor * 0.9)
        scaled_height = int(BASE_WINDOW_HEIGHT * self.scale_factor * 0.9)
        
        # Ensure the window is not larger than the available screen geometry.
        width = min(scaled_width, self.geometry.width())
        height = min(scaled_height, self.geometry.height())
        
        return QSize(width, height)

    def scale(self, base_value: int) -> int:
        """Scales a base pixel value (e.g., width, margin) by the DPI scale factor."""
        return int(base_value * self.scale_factor)

    def scale_font_size(self, base_size: int) -> int:
        """Scales a base font size by the DPI scale factor."""
        return max(8, int(base_size * self.scale_factor * 0.95))

class GlobalStylesheet:
    """
    Generates the global CSS stylesheet for the application, using the
    DisplayManager to ensure all UI elements are properly scaled.
    THIS CLASS AND ITS OUTPUT ARE PRESERVED EXACTLY FROM THE ORIGINAL MAIN.PY.
    """
    def __init__(self, dm: DisplayManager):
        self.dm = dm

    def get(self) -> str:
        # This large f-string generates a dynamic stylesheet with scaled pixel values.
        # It defines the dark, "military-grade" theme of the application.
        return f"""
            #MainWindow, #ScrollArea, #ScrollArea > QWidget > QWidget {{
                background-color: #000000;
                border: none;
            }}
            
            QScrollBar:vertical {{
                border: none;
                background: #1a1a1a;
                width: {self.dm.scale(12)}px;
                margin: 0;
                border-radius: {self.dm.scale(6)}px;
            }}
            QScrollBar::handle:vertical {{
                background: #555555;
                min-height: {self.dm.scale(30)}px;
                border-radius: {self.dm.scale(6)}px;
            }}
            QScrollBar::handle:vertical:hover {{ background: #00ff00; }}
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0px; }}
            
            QScrollBar:horizontal {{
                border: none;
                background: #2a2a2a;
                height: {self.dm.scale(12)}px;
                margin: 0;
                border-radius: {self.dm.scale(6)}px;
            }}
            QScrollBar::handle:horizontal {{
                background: #555555;
                min-width: {self.dm.scale(30)}px;
                border-radius: {self.dm.scale(6)}px;
            }}
            QScrollBar::handle:horizontal:hover {{ background: #00ff00; }}
            
            QLabel {{ 
                color: #ffffff; 
                background-color: transparent; 
            }}
            
            #CardTitle {{
                font-size: {self.dm.scale_font_size(20)}px;
                font-weight: 700;
                padding-bottom: {self.dm.scale(10)}px;
                margin-bottom: {self.dm.scale(8)}px;
            }}
            
            #TipBox {{
                background-color: #2c2c2e;
                border-radius: {self.dm.scale(10)}px;
                padding: {self.dm.scale(22)}px;
                color: #dddddd;
                font-style: italic;
                font-size: {self.dm.scale_font_size(15)}px;
                border: 1px solid #444444;
                line-height: 1.5;
            }}
            
            .Label {{
                color: #ffffff;
                font-size: {self.dm.scale_font_size(15)}px;
                font-weight: 700;
            }}
            
            .Value {{
                color: #00ff00;
                font-size: {self.dm.scale_font_size(15)}px;
                font-weight: 600;
            }}
            
            #Separator {{ 
                background-color: rgba(255, 255, 255, 0.12); 
            }}
            
            QPushButton#HistoryButton {{
                background-color: #00bfff;
                color: white;
                border: none;
                border-radius: {self.dm.scale(8)}px;
                padding: {self.dm.scale(14)}px {self.dm.scale(24)}px;
                font-weight: 700;
                font-size: {self.dm.scale_font_size(14)}px;
            }}
            QPushButton#HistoryButton:hover {{ background-color: #1e90ff; }}
            QPushButton#HistoryButton:pressed {{ background-color: #0080c0; }}
            
            QPushButton#AiButton {{
                background-color: #3e3e3e;
                color: #00bfff;
                border: 2px solid #00bfff;
                border-radius: {self.dm.scale(8)}px;
                padding: {self.dm.scale(14)}px {self.dm.scale(24)}px;
                font-weight: 700;
                font-size: {self.dm.scale_font_size(14)}px;
            }}
            QPushButton#AiButton:hover {{
                background-color: #4a4a4a;
                border-color: #1e90ff;
            }}
            QPushButton#AiButton:pressed {{ background-color: #2e2e2e; }}
            
            QPushButton#NavCircleButton {{
                background-color: #3e3e3e;
                color: #cccccc;
                border: 2px solid #555555;
                border-radius: {self.dm.scale(20)}px;
                font-size: {self.dm.scale_font_size(18)}px;
                font-weight: bold;
            }}
            QPushButton#NavCircleButton:hover {{
                background-color: #4a4a4a;
                color: #00ff00;
                border-color: #00ff00;
            }}
            QPushButton#NavCircleButton:pressed {{ background-color: #2e2e2e; }}
            
            QProgressBar {{
                border: none;
                background-color: #2a2a2a;
                border-radius: {self.dm.scale(12)}px;
                text-align: center;
                color: #00ff00;
                font-weight: bold;
                font-size: {self.dm.scale_font_size(12)}px;
                min-height: {self.dm.scale(24)}px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00ff00, stop:0.5 #00dd00, stop:1 #00aa00);
                border-radius: {self.dm.scale(12)}px;
            }}
            
            QMenuBar {{
                background-color: #2a2a2a;
                color: #ffffff;
                padding: {self.dm.scale(5)}px;
            }}
            QMenuBar::item {{
                padding: {self.dm.scale(8)}px {self.dm.scale(15)}px;
                background: transparent;
            }}
            QMenuBar::item:selected {{
                background-color: #00bfff;
                border-radius: {self.dm.scale(4)}px;
            }}
            
            QMenu {{
                background-color: #2a2a2a;
                color: #ffffff;
                border: 1px solid #444444;
                border-radius: {self.dm.scale(6)}px;
                padding: {self.dm.scale(5)}px;
            }}
            QMenu::item {{
                padding: {self.dm.scale(8)}px {self.dm.scale(25)}px;
                border-radius: {self.dm.scale(4)}px;
            }}
            QMenu::item:selected {{
                background-color: #00bfff;
            }}
            
            QMessageBox {{
                background-color: #2a2a2a;
            }}
            QMessageBox QLabel {{
                color: #ffffff;
                font-size: {self.dm.scale_font_size(14)}px;
            }}
            QMessageBox QPushButton {{
                background-color: #00bfff;
                color: white;
                border: none;
                border-radius: {self.dm.scale(6)}px;
                padding: {self.dm.scale(10)}px {self.dm.scale(20)}px;
                font-weight: 600;
                min-width: {self.dm.scale(80)}px;
            }}
            QMessageBox QPushButton:hover {{ background-color: #1e90ff; }}
            
            QDialog {{
                background-color: #1a1a1a;
            }}
            
            QInputDialog {{
                background-color: #2a2a2a;
            }}
            QInputDialog QLabel {{
                color: #ffffff;
            }}
            QInputDialog QLineEdit, QInputDialog QSpinBox {{
                background-color: #3a3a3a;
                color: #ffffff;
                border: 1px solid #555555;
                border-radius: {self.dm.scale(4)}px;
                padding: {self.dm.scale(8)}px;
            }}
        """

class SystemTrayManager:
    """
    Manages the application's system tray icon, its context menu, and its
    associated actions. PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, parent_window):

# ============================================================================
# PART 5
# ============================================================================
        # The __init__ method is the constructor for the class.
        self.parent = parent_window
        # Initialize the tray icon attribute to None.
        self.tray_icon = None
        
        # Use a try-except block to handle potential errors during tray icon creation.
        try:
            # Create an instance of QSystemTrayIcon, associating it with the main window.
            self.tray_icon = QSystemTrayIcon(parent_window)
            
            # Check if the specified icon file exists.
            if os.path.exists(LOGO_ICO_PATH):
                # If it exists, create a QIcon from the file.
                icon = QIcon(LOGO_ICO_PATH)
            # If the icon file is missing, create a fallback icon programmatically.
            else:
                # Create a blank QPixmap to draw on.
                pixmap = QPixmap(32, 32)
                # Fill it with a transparent background.
                pixmap.fill(Qt.GlobalColor.transparent)
                # Create a QPainter to draw the custom icon.
                painter = QPainter(pixmap)
                # Enable antialiasing for smooth shapes.
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                # Set the brush to a solid green color.
                painter.setBrush(QBrush(QColor(0, 255, 0)))
                # Disable the pen to avoid an outline.
                painter.setPen(Qt.PenStyle.NoPen)
                # Draw a simple circle as the fallback icon.
                painter.drawEllipse(2, 2, 28, 28)
                # Finalize the drawing operations.
                painter.end()
                # Create a QIcon from the generated pixmap.
                icon = QIcon(pixmap)
            
            # Set the icon for the system tray item.
            self.tray_icon.setIcon(icon)
            # Set the tooltip text that appears when the user hovers over the tray icon.
            self.tray_icon.setToolTip(f"{APP_NAME} - Battery Monitoring")
            
            # Create the right-click context menu for the tray icon.
            self._create_context_menu()
            
            # Connect the 'activated' signal of the tray icon to a handler method.
            # This signal is emitted when the user clicks or double-clicks the icon.
            self.tray_icon.activated.connect(self._on_tray_activated)
            
            # Make the tray icon visible in the system tray.
            self.tray_icon.show()
            
            # Log the successful initialization.
            logging.info("System tray manager initialized successfully.")
            
        # Catch any exception that occurs during initialization.
        except Exception as e:
            # Log the failure. The application can still run without a tray icon.
            logging.error("System tray initialization failed: %s", e)

    def _create_context_menu(self):
        """Creates the right-click menu for the system tray icon."""
        # Create a QMenu instance.
        menu = QMenu()
        
        # --- Create Actions ---
        # Action to show the main window.
        show_action = QAction("Show Battery-Z", menu)
        # Action to refresh all battery data.
        refresh_action = QAction("🔄 Refresh Data", menu)
        # Placeholder actions to display live data. These are disabled as they are for display only.
        self.battery_info_action = QAction("Battery: ---%", menu)
        self.battery_info_action.setEnabled(False)
        self.health_info_action = QAction("Health: ---%", menu)
        self.health_info_action.setEnabled(False)
        self.temp_info_action = QAction("Temp: --°C", menu)
        self.temp_info_action.setEnabled(False)
        # Action to show the "About" dialog.
        about_action = QAction("About Battery-Z", menu)
        # Action to exit the application.
        exit_action = QAction("❌ Exit", menu)
        
        # --- Connect Actions to Methods ---
        # When an action is triggered (clicked), it will call the corresponding method.
        show_action.triggered.connect(self.parent.show)
        refresh_action.triggered.connect(self.parent.refresh_application)
        about_action.triggered.connect(self.parent.show_about_dialog)
        exit_action.triggered.connect(self.parent.close)
        
        # --- Assemble the Menu ---
        # Add actions and separators to the menu to structure it logically.
        menu.addAction(show_action)
        menu.addSeparator()
        menu.addAction(refresh_action)
        menu.addSeparator()
        menu.addAction(self.battery_info_action)
        menu.addAction(self.health_info_action)
        menu.addAction(self.temp_info_action)
        menu.addSeparator()
        menu.addAction(about_action)
        menu.addSeparator()
        menu.addAction(exit_action)
        
        # Set the created menu as the context menu for the tray icon.
        self.tray_icon.setContextMenu(menu)

    def _on_tray_activated(self, reason):
        """Handles clicks on the system tray icon."""
        # Check if the activation reason was a double-click.
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            # If the window is already visible, hide it.
            if self.parent.isVisible():
                self.parent.hide()
            # If the window is hidden, show it and bring it to the front.
            else:
                self.parent.show()
                self.parent.activateWindow()

    def update_battery_info(self, battery_data: BatteryData, soh_percentage: float):
        """Updates the text of the placeholder actions in the context menu with live data."""
        try:
            # Update the battery percentage text.
            if battery_data.current_percentage is not None:
                # Use a different icon depending on the charging state.
                charging_icon = "⚡" if battery_data.is_charging else "🔋"
                self.battery_info_action.setText(
                    f"{charging_icon} Battery: {battery_data.current_percentage}%"
                )
            
            # Update the health percentage text.
            if soh_percentage > 0:
                self.health_info_action.setText(f"Health: {soh_percentage:.0f}%")
            
            # Update the temperature text.
            if battery_data.temperature_celsius is not None:
                self.temp_info_action.setText(
                    f"Temp: {battery_data.temperature_celsius:.1f}°C"
                )
        except Exception as e:
            logging.error("Failed to update tray menu text: %s", e)

    def show_notification(self, title: str, message: str, icon: QSystemTrayIcon.MessageIcon = QSystemTrayIcon.MessageIcon.Information):
        """Displays a desktop notification from the system tray icon."""
        try:
            # Check if the tray icon is valid and visible.
            if self.tray_icon and self.tray_icon.isVisible():
                # Call the showMessage method to display the notification bubble.
                self.tray_icon.showMessage(title, message, icon, 5000) # 5000ms timeout
        except Exception as e:
            logging.error("Failed to show tray notification: %s", e)

    def cleanup(self):
        """Hides the tray icon upon application exit."""
        try:
            if self.tray_icon:
                self.tray_icon.hide()
        except Exception as e:
            logging.warning("Error during tray icon cleanup: %s", e)

class NotificationManager:
    """
    Manages the logic for sending desktop notifications, including cooldowns
    to prevent spamming the user with alerts.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, tray_manager: Optional[SystemTrayManager] = None):
        # Store a reference to the SystemTrayManager to show notifications.
        self.tray_manager = tray_manager
        # Timestamps for the last time a notification of each type was sent.
        self.last_high_temp_notification = None
        self.last_low_battery_notification = None
        # Cooldown period in seconds (5 minutes) to prevent notification spam.
        self.notification_cooldown = 300
        logging.info("Notification manager initialized.")
        
    def _can_send_notification(self, notification_type: str) -> bool:
        """Checks if a notification of a given type can be sent based on the cooldown."""
        # Get the current time.
        now = datetime.datetime.now()
        
        # Check cooldown for high temperature alerts.
        if notification_type == "high_temp":
            if self.last_high_temp_notification:
                # Calculate time elapsed since the last notification.
                elapsed = (now - self.last_high_temp_notification).total_seconds()
                # If cooldown period has not passed, return False.
                if elapsed < self.notification_cooldown:
                    return False
            # If we can send, update the timestamp to the current time.
            self.last_high_temp_notification = now
        
        # Check cooldown for low battery alerts.
        elif notification_type == "low_battery":
            if self.last_low_battery_notification:
                elapsed = (now - self.last_low_battery_notification).total_seconds()
                if elapsed < self.notification_cooldown:
                    return False
            self.last_low_battery_notification = now
            
        # If cooldown has passed, return True.
        return True

    def send_high_temperature_alert(self, temperature: float):
        """Sends a high temperature notification if the cooldown allows."""
        # Check if we are allowed to send this notification.
        if not self._can_send_notification("high_temp"):
            return
        
        # Define the notification title and message.
        title = "⚠️ Battery Temperature Alert"
        message = (f"Battery temperature is critical: {temperature:.1f}°C!\n"
                   f"High temperatures can permanently damage your battery. "
                   f"Please ensure good ventilation and close intensive applications.")
        
        # Use the tray manager to show the notification.
        if self.tray_manager:
            self.tray_manager.show_notification(title, message, QSystemTrayIcon.MessageIcon.Warning)
        
        logging.warning("High temperature alert sent for %.1f°C", temperature)

    def send_low_battery_alert(self, percentage: int):
        """Sends a low battery notification if the cooldown allows."""
        if not self._can_send_notification("low_battery"):
            return
            
        title = "🔋 Low Battery Warning"
        message = f"Battery is at {percentage}%. Please connect your charger soon."
        
        if self.tray_manager:
            self.tray_manager.show_notification(title, message, QSystemTrayIcon.MessageIcon.Warning)
            
        logging.warning("Low battery alert sent for %d%%", percentage)

# ============================================================================
# SECTION 9: CUSTOM UI WIDGETS (PART 1)
# Description: This section contains the custom PyQt5 widgets that make up the
#              application's unique user interface. Their code is preserved
#              exactly from the original main.py to maintain the visual design.
#              Extensive comments are added to explain their function.
# ============================================================================

class SplashScreen(QSplashScreen):
    """
    A highly customized, animated splash screen that displays during the initial
    data loading phase of the application. It provides visual feedback and a
    professional loading experience. PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, dm: DisplayManager):
        # --- Initialization ---
        # Calculate the scaled size of the splash screen pixmap.
        width = dm.scale(800)
        height = dm.scale(550)
        # Create a transparent QPixmap to serve as the canvas for our custom-drawn splash screen.
        pixmap = QPixmap(width, height)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        # Call the parent constructor with the transparent pixmap.
        super().__init__(pixmap)
        # Store a reference to the DisplayManager for scaling.
        self.dm = dm
        
        # --- Window Flags and Attributes ---
        # Set flags to make the window frameless, stay on top, and function as a splash screen.
        self.setWindowFlags(Qt.WindowType.SplashScreen | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        # Enable translucent background to allow for custom shapes and transparency.
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # --- Positioning ---
        # Get the primary screen's geometry.
        screen_geometry = QApplication.primaryScreen().availableGeometry()
        # Move the splash screen to the center of the screen.
        self.move(screen_geometry.center() - self.rect().center())
        
        # --- UI and Animation Setup ---
        # Build the internal widgets of the splash screen.
        self._setup_ui()
        
        # Initialize animation properties.
        self._rotation_angle = 0
        self._glow_opacity = 0.3
        self._pulse_scale = 1.0
        
        # --- Timers and Animations ---
        # A QTimer to drive the rotation of the loading spinner.
        self.rotation_timer = QTimer(self)
        self.rotation_timer.timeout.connect(self._animate_loader)
        self.rotation_timer.start(30) # Update every 30ms for a smooth animation.
        
        # A QPropertyAnimation for the pulsing glow effect.
        self.glow_animation = QPropertyAnimation(self, b"glow_opacity")
        self.glow_animation.setDuration(2000)
        self.glow_animation.setStartValue(0.3)
        self.glow_animation.setEndValue(1.0)
        self.glow_animation.setLoopCount(-1) # Loop indefinitely.
        self.glow_animation.setEasingCurve(QEasingCurve.Type.InOutSine)
        self.glow_animation.start()
        
        # A QPropertyAnimation for the subtle pulsing scale of the loader.
        self.pulse_animation = QPropertyAnimation(self, b"pulse_scale")
        self.pulse_animation.setDuration(1500)
        self.pulse_animation.setStartValue(0.95)
        self.pulse_animation.setEndValue(1.05)
        self.pulse_animation.setLoopCount(-1)
        self.pulse_animation.setEasingCurve(QEasingCurve.Type.InOutSine)
        self.pulse_animation.start()
    
    # --- PyQt Properties for Animation ---
    # These properties allow QPropertyAnimation to animate custom attributes.
    
    # Glow Opacity Property
    @pyqtProperty(float)
    def glow_opacity(self):
        return self._glow_opacity
    
    @glow_opacity.setter
    def glow_opacity(self, opacity: float):
        self._glow_opacity = opacity
        self.update() # Trigger a repaint when the property changes.
        
    # Pulse Scale Property
    @pyqtProperty(float)
    def pulse_scale(self):
        return self._pulse_scale
    
    @pulse_scale.setter
    def pulse_scale(self, scale: float):
        self._pulse_scale = scale
        self.update() # Trigger a repaint.

    def _setup_ui(self):
        """Builds the internal layout and widgets of the splash screen."""
        # Create a container widget that will hold all other elements.
        container = QWidget(self)
        container.setGeometry(self.rect())
        
        # Main vertical layout for the container.
        layout = QVBoxLayout(container)
        layout.setContentsMargins(
            self.dm.scale(50), self.dm.scale(50), self.dm.scale(50), self.dm.scale(50)
        )
        layout.setSpacing(self.dm.scale(30))
        
        # --- Header Section (Logo and Title) ---
        logo_title_layout = QHBoxLayout()
        logo_title_layout.setSpacing(self.dm.scale(20))
        
        # Add the custom logo widget.
        logo = LogoWidget(self.dm, size=70)
        logo_title_layout.addWidget(logo)
        
        # Vertical layout for the title and subtitle.
        title_container = QVBoxLayout()
        title_container.setSpacing(self.dm.scale(5))
        
        # Application Title Label.
        title = QLabel("Battery-Z")
        title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        title.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(56)}px; font-weight: 900;
            color: #00ff00; letter-spacing: 2px; background: transparent;
        """)
        
        # Add a glow effect to the title text.
        glow = QGraphicsDropShadowEffect()
        glow.setBlurRadius(40)
        glow.setColor(QColor(0, 255, 0, 220))
        glow.setOffset(0, 0)
        title.setGraphicsEffect(glow)
        
        # Subtitle Label.
        subtitle = QLabel(APP_SUBTITLE)
        subtitle.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        subtitle.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(16)}px; color: #bbbbbb;
            font-weight: 600; background: transparent;
        """)
        
        title_container.addWidget(title)
        title_container.addWidget(subtitle)
        logo_title_layout.addLayout(title_container)
        logo_title_layout.addStretch()
        layout.addLayout(logo_title_layout)
        
        layout.addSpacing(self.dm.scale(20))
        
        # --- Loading Spinner Section ---
        self.loading_label = QLabel(self)
        self.loading_label.setFixedSize(self.dm.scale(100), self.dm.scale(100))
        self.loading_label.setStyleSheet("background: transparent;")
        
        loading_layout = QHBoxLayout()
        loading_layout.addStretch()
        loading_layout.addWidget(self.loading_label)
        loading_layout.addStretch()
        layout.addLayout(loading_layout)
        
        # --- Progress Bar Section ---
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setStyleSheet(f"""
            QProgressBar {{
                border: none; background: rgba(42, 42, 42, 0.8);
                border-radius: {self.dm.scale(15)}px; text-align: center;
                color: #00ff00; font-weight: bold;
                font-size: {self.dm.scale_font_size(13)}px;
                min-height: {self.dm.scale(30)}px;
            }}
            QProgressBar::chunk {{
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 rgba(0, 255, 0, 0.9), stop:1 rgba(0, 170, 0, 0.9));
                border-radius: {self.dm.scale(15)}px;
            }}
        """)
        layout.addWidget(self.progress_bar)
        
        # --- Progress Text Label ---
        self.progress_label = QLabel("Initializing...")
        self.progress_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_label.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(15)}px; color: #00bfff;
            font-weight: 700; background: transparent;
        """)
        layout.addWidget(self.progress_label)
        
        layout.addStretch()
        
        # --- Footer Section (Version Info) ---
        version_info = QLabel(f"Version {APP_VERSION} | Build {BUILD_DATE}")
        version_info.setAlignment(Qt.AlignmentFlag.AlignCenter)
        version_info.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(11)}px;
            color: #666666; background: transparent;
        """)
        layout.addWidget(version_info)

# ============================================================================
# PART 6
# ============================================================================
    
    def _animate_loader(self):
        """This method is called by a timer to update the rotation angle of the spinner."""
        # Increment the rotation angle. Using a modulo operator ensures it wraps around 360 degrees.
        self._rotation_angle = (self._rotation_angle + 8) % 360
        # Call the drawing function to repaint the loader with the new angle.
        self._draw_3d_glassy_loader()

    def _draw_3d_glassy_loader(self):
        """Performs the custom drawing of the animated loading spinner onto a QPixmap."""
        # Get the size of the label widget that will display the spinner.
        size = self.loading_label.size()
        # Create a new transparent pixmap to draw on for this frame of the animation.
        pixmap = QPixmap(size)
        pixmap.fill(Qt.GlobalColor.transparent)
        
        # Create a QPainter to perform the drawing operations on our pixmap.
        painter = QPainter(pixmap)
        # Enable antialiasing for smooth, high-quality graphics.
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # --- Drawing Calculations ---
        # Define the center point and radius for the spinner, with a margin.
        center = QPointF(size.width() / 2, size.height() / 2)
        radius = min(size.width(), size.height()) / 2 - self.dm.scale(10)
        
        # --- Draw Background Glow ---
        # Create a radial gradient for the soft glow effect behind the spinner.
        glow_gradient = QRadialGradient(center, radius * 1.2)
        # The color alpha is animated via the _glow_opacity property.
        glow_gradient.setColorAt(0, QColor(0, 255, 0, int(80 * self._glow_opacity)))
        glow_gradient.setColorAt(0.7, QColor(0, 255, 0, int(40 * self._glow_opacity)))
        glow_gradient.setColorAt(1, QColor(0, 255, 0, 0)) # Fades to transparent.
        
        # Apply the gradient brush and draw the ellipse for the glow.
        painter.setBrush(QBrush(glow_gradient))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(center, radius * 1.2, radius * 1.2)
        
        # Apply the pulsing scale animation value to the radius.
        scaled_radius = radius * self._pulse_scale
        
        # --- Draw the Spinning Arc ---
        # Create a conical gradient to give the arc its fading-in-and-out look.
        # The gradient's angle is rotated by the _rotation_angle property.
        conical_gradient = QConicalGradient(center, self._rotation_angle)
        conical_gradient.setColorAt(0.0, QColor(0, 255, 0, 0))   # Transparent start
        conical_gradient.setColorAt(0.3, QColor(0, 255, 0, 255)) # Opaque middle
        conical_gradient.setColorAt(0.7, QColor(0, 255, 0, 255)) # Opaque middle
        conical_gradient.setColorAt(1.0, QColor(0, 255, 0, 0))   # Transparent end
        
        # Create a thick pen using the conical gradient as its brush.
        pen = QPen(QBrush(conical_gradient), self.dm.scale(8))
        pen.setCapStyle(Qt.PenCapStyle.RoundCap) # Use round caps for a softer look.
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush)
        
        # Define the bounding rectangle for the arc, using the scaled radius.
        arc_rect = QRectF(
            center.x() - scaled_radius, center.y() - scaled_radius,
            scaled_radius * 2, scaled_radius * 2
        )
        
        # Draw the arc. Angles are specified in 1/16th of a degree.
        start_angle = 0
        span_angle = 270 * 16 # Draw a 270-degree arc.
        painter.drawArc(arc_rect, start_angle, span_angle)
        
        # --- Draw Inner Core ---
        # Draw a subtle inner glow to add depth.
        inner_radius = scaled_radius * 0.6
        inner_gradient = QRadialGradient(center, inner_radius)
        inner_gradient.setColorAt(0, QColor(0, 255, 0, 30))
        inner_gradient.setColorAt(0.8, QColor(0, 255, 0, 80))
        inner_gradient.setColorAt(1, QColor(0, 255, 0, 0))
        
        painter.setBrush(QBrush(inner_gradient))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(center, inner_radius, inner_radius)
        
        # Finalize painting operations for this frame.
        painter.end()
        
        # Set the newly drawn pixmap as the content of the loading label.
        self.loading_label.setPixmap(pixmap)

    def paintEvent(self, event):
        """Overrides the default paint event to draw the custom splash screen background."""
        # Create a QPainter for the splash screen widget itself.
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Get the widget's rectangle.
        rect = self.rect()
        
        # --- Draw Main Background ---
        # Create a linear gradient for the main body of the splash screen.
        bg_gradient = QLinearGradient(rect.topLeft(), rect.bottomRight())
        bg_gradient.setColorAt(0, QColor(26, 26, 26, 250))
        bg_gradient.setColorAt(0.5, QColor(35, 35, 35, 250))
        bg_gradient.setColorAt(1, QColor(26, 26, 26, 250))
        
        # Set the brush and a subtle border pen.
        painter.setBrush(QBrush(bg_gradient))
        painter.setPen(QPen(QColor(0, 255, 0, 100), 2))
        # Draw the main rounded rectangle background.
        painter.drawRoundedRect(rect, self.dm.scale(20), self.dm.scale(20))
        
        # --- Draw Glossy Overlay ---
        # Create a second, subtle gradient to simulate a glossy reflection.
        overlay_gradient = QLinearGradient(rect.topLeft(), rect.bottomLeft())
        overlay_gradient.setColorAt(0, QColor(255, 255, 255, 20))
        overlay_gradient.setColorAt(0.5, QColor(255, 255, 255, 5))
        overlay_gradient.setColorAt(1, QColor(0, 0, 0, 10))
        
        # Draw the overlay with no border.
        painter.setBrush(QBrush(overlay_gradient))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(rect, self.dm.scale(20), self.dm.scale(20))

    def update_progress(self, message: str, progress: int = None):
        """Public method to update the progress text and bar from the main thread."""
        # Set the text on the progress label.
        self.progress_label.setText(message)
        # If a progress value is provided, update the progress bar.
        if progress is not None:
            self.progress_bar.setValue(progress)
        # Process events to ensure the UI updates immediately.
        QApplication.processEvents()

    def finish_with_fade(self, main_window):
        """Stops animations and schedules the final fade-out and close."""
        # Stop all running timers and animations.
        self.rotation_timer.stop()
        self.glow_animation.stop()
        self.pulse_animation.stop()
        # Schedule the finish() method to be called after a short delay (400ms).
        # This allows the last progress update to be visible before closing.
        QTimer.singleShot(400, lambda: self.finish(main_window))

class LogoWidget(QWidget):
    """
    A dedicated widget for displaying the application logo. It handles loading
    from SVG or ICO files with a programmatic fallback.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, dm: DisplayManager, size: int = 55):
        super().__init__()
        self.dm = dm
        # Set the widget's size, scaled by the DPI factor.
        self.logo_size = dm.scale(size)
        self.setFixedSize(self.logo_size, self.logo_size)
        
        # Initialize pixmap and renderer to None.
        self.logo_pixmap = None
        self.svg_renderer = None
        
        # --- Logo Loading Logic ---
        # Method 1: Try to load the SVG logo first (preferred for scalability).
        if os.path.exists(LOGO_SVG_PATH):
            try:
                self.svg_renderer = QSvgRenderer(LOGO_SVG_PATH)
                print(f"[✓] SVG logo loaded successfully from: {LOGO_SVG_PATH}")
            except Exception as e:
                print(f"[!] Failed to load SVG logo: {e}")
        
        # Method 2: If SVG fails or is missing, try to load the ICO logo.
        if not self.svg_renderer and os.path.exists(LOGO_ICO_PATH):
            try:
                self.logo_pixmap = QPixmap(LOGO_ICO_PATH)
                print(f"[✓] ICO logo loaded successfully from: {LOGO_ICO_PATH}")
            except Exception as e:
                print(f"[!] Failed to load ICO logo: {e}")
                
        # If both fail, a fallback will be drawn in the paintEvent.
        if not self.svg_renderer and not self.logo_pixmap:
            print("[!] No logo files found. A fallback design will be drawn.")

    def paintEvent(self, event):
        """Draws the logo (from file or fallback) onto the widget."""
        painter = QPainter(self)
        # Enable high-quality rendering hints.
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        
        rect_f = QRectF(self.rect())
        
        # If an SVG renderer is available, use it to render the logo.
        if self.svg_renderer:
            self.svg_renderer.render(painter, rect_f)
        
        # If a pixmap (from ICO) is available, draw it scaled to fit.
        elif self.logo_pixmap:
            scaled_pixmap = self.logo_pixmap.scaled(
                self.logo_size, self.logo_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            # Center the pixmap within the widget area.
            x = (self.logo_size - scaled_pixmap.width()) // 2
            y = (self.logo_size - scaled_pixmap.height()) // 2
            painter.drawPixmap(x, y, scaled_pixmap)
            
        # If no logo files were loaded, draw a programmatic fallback logo.
        else:
            rect = self.rect()
            center = rect.center()
            # Draw a green radial gradient as a background.
            gradient = QRadialGradient(QPointF(center), rect.width() / 2)
            gradient.setColorAt(0, QColor(0, 255, 0, 220))
            gradient.setColorAt(1, QColor(0, 255, 0, 0))
            painter.setBrush(QBrush(gradient))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(rect)
            # Draw a stylized 'Z' or lightning bolt shape.
            path = QPainterPath()
            w, h = rect.width(), rect.height()
            path.moveTo(w * 0.55, h * 0.2)
            path.lineTo(w * 0.4, h * 0.5)
            path.lineTo(w * 0.5, h * 0.5)
            path.lineTo(w * 0.35, h * 0.8)
            path.lineTo(w * 0.5, h * 0.55)
            path.lineTo(w * 0.45, h * 0.55)
            path.closeSubpath()
            painter.setBrush(QColor(20, 20, 20))
            painter.setPen(QPen(QColor(0, 255, 0), 1))
            painter.drawPath(path)

class SeparatorLine(QFrame):
    """A simple, styled horizontal line used as a separator in the UI cards."""
    def __init__(self):
        super().__init__()
        # Set the frame shape to a horizontal line.
        self.setFrameShape(QFrame.Shape.HLine)
        # Set a fixed height of 1 pixel.
        self.setFixedHeight(1)
        # Set an object name so it can be styled by the global stylesheet.
        self.setObjectName("Separator")

class InfoRowWidget(QWidget):
    """
    A compound widget representing a single row in an info card, consisting of
    a left-aligned label and a right-aligned value.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, label_text: str, value_text: str, dm: DisplayManager):
        super().__init__()
        
        # Use a horizontal layout for the label and value.
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(dm.scale(10))
        
        # Create the label for the data point (e.g., "Manufacturer").
        self.label = QLabel(label_text)
        self.label.setProperty("class", "Label") # Set a CSS class for styling.
        self.label.setWordWrap(True)
        
        # Create the label for the value (e.g., "Dell Inc.").
        self.value = QLabel(value_text)
        self.value.setProperty("class", "Value") # Set a CSS class for styling.
        self.value.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        self.value.setWordWrap(True)
        
        # Add widgets to the layout. The stretch ensures the value is pushed to the right.
        layout.addWidget(self.label)
        layout.addStretch(1)
        layout.addWidget(self.value)

    def setValue(self, text: str):
        """Public method to update the value text."""
        self.value.setText(text)

    def setStyle(self, style: str):
        """Public method to apply a custom stylesheet to the value label (e.g., for color changes)."""
        self.value.setStyleSheet(style)

class HealthCircleWidget(QWidget):
    """
    A complex custom widget for displaying battery health in a circular gauge
    with animations and a pulsing glow effect.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, dm: DisplayManager, parent=None):
        super().__init__(parent)
        
        # --- Initialization ---
        self._health_percentage = 0     # The current displayed health value.
        self._glow_opacity = 0.5        # The current opacity for the glow animation.
        self._status_color = "#888888"  # The current color of the gauge and text.
        self._animation_progress = 0.0  # The target value for the number animation.
        self.dm = dm                    # Store the DisplayManager for scaling.
        
        # Set a minimum size to ensure the widget is visible.
        self.setMinimumSize(dm.scale(240), dm.scale(240))
        
        # --- Child Widgets ---
        # The large percentage label in the center.
        self.percentage_label = QLabel("0%", self)
        self.percentage_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.percentage_label.setStyleSheet(f"""
            font-size: {dm.scale_font_size(56)}px; font-weight: 900; color: #888888;
        """)
        
        # The smaller status text label below the percentage.
        self.status_label = QLabel("Calculating...", self)
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(f"""
            font-size: {dm.scale_font_size(16)}px; color: #888888; font-weight: 700;
        """)
        
        # --- Layout ---
        # Use a QVBoxLayout to center the labels vertically within the circle.
        layout = QVBoxLayout(self)
        layout.addStretch(2)
        layout.addWidget(self.percentage_label)
        layout.addWidget(self.status_label)
        layout.addStretch(2)
        layout.setSpacing(dm.scale(5))
        
        # --- Animation Setup ---
        self._setup_animations()

    def _setup_animations(self):
        """Initializes the QPropertyAnimation objects for the widget."""
        # Animation for the pulsing background glow opacity.
        self.glow_animation = QPropertyAnimation(self, b"glow_opacity")
        self.glow_animation.setDuration(2500)
        self.glow_animation.setStartValue(0.3)
        self.glow_animation.setEndValue(1.0)
        self.glow_animation.setLoopCount(-1) # Loop forever.
        self.glow_animation.setEasingCurve(QEasingCurve.Type.InOutSine)
        self.glow_animation.start()
        
        # Animation for the "counting up" of the percentage number and gauge arc.
        self.percentage_animation = QPropertyAnimation(self, b"animation_progress")
        self.percentage_animation.setDuration(800)
        self.percentage_animation.setEasingCurve(QEasingCurve.Type.OutCubic)

    def _get_color_for_percentage(self, percentage: float) -> str:
        """Helper method to get the correct color for a given health percentage."""
        # This logic is duplicated from the global helper but kept here for widget encapsulation.
        if percentage >= 95: return "#00ff00"
        elif percentage >= 85: return "#7cfc00"
        elif percentage >= 75: return "#9a67ea"
        elif percentage >= 60: return "#ffa500"
        elif percentage >= 40: return "#ff8c00"
        else: return "#ff0000"

# ============================================================================
# PART 7
# ============================================================================
    
    def setHealth(self, percentage: float, status_text: str, status_color: str, animate: bool = True):
        """
        Public method to update the health value displayed by the widget. It can
        either animate the change or set it instantly.

        Args:
            percentage (float): The new health percentage to display.
            status_text (str): The descriptive text (e.g., "Good").
            status_color (str): The color for the text and gauge.
            animate (bool): If True, the change will be animated.
        """
        # Store the previous percentage value for the animation's start point.
        old_percentage = self._health_percentage
        # Clamp the new percentage between 0 and 100 for safety.
        new_percentage = max(0.0, min(100.0, percentage))
        
        # Get the appropriate color for the new percentage value.
        actual_color = self._get_color_for_percentage(new_percentage)
        
        # Store the target color for use in the paintEvent.
        self._status_color = actual_color
        
        # Update the status text label and its color.
        self.status_label.setText(status_text)
        self.status_label.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(16)}px;
            color: {actual_color};
            font-weight: 700;
        """)
        
        # If animation is requested and the change is significant...
        if animate and abs(new_percentage - old_percentage) > 0.5:
            # Stop any ongoing animation.
            self.percentage_animation.stop()
            # Set the start and end values for the animation.
            self.percentage_animation.setStartValue(old_percentage)
            self.percentage_animation.setEndValue(new_percentage)
            # Start the animation.
            self.percentage_animation.start()
        # If no animation is requested...
        else:
            # Instantly set the new health value.
            self._health_percentage = new_percentage
            self._animation_progress = new_percentage
            # Update the percentage text label and its color.
            self.percentage_label.setText(f"{int(new_percentage)}%")
            self.percentage_label.setStyleSheet(f"""
                font-size: {self.dm.scale_font_size(56)}px;
                font-weight: 900;
                color: {actual_color};
            """)
            # Trigger a repaint to show the changes immediately.
            self.update()

    # --- PyQt Properties for Animation ---

    # Property for the glow opacity animation.
    @pyqtProperty(float)
    def glow_opacity(self):
        return self._glow_opacity
    
    @glow_opacity.setter
    def glow_opacity(self, opacity: float):
        self._glow_opacity = opacity
        self.update() # Repaint on change.

    # Property for the health percentage animation.
    @pyqtProperty(float)
    def animation_progress(self):
        return self._animation_progress

    @animation_progress.setter
    def animation_progress(self, progress: float):
        """
        This setter is called by the QPropertyAnimation on every frame of the
        health percentage animation.
        """
        # Update the internal progress value.
        self._animation_progress = progress
        # This is the actual value that the paintEvent will use to draw the arc.
        self._health_percentage = progress
        
        # Get the color corresponding to the current progress value.
        actual_color = self._get_color_for_percentage(progress)
        self._status_color = actual_color
        
        # Update the text and color of the main percentage label.
        self.percentage_label.setText(f"{int(progress)}%")
        self.percentage_label.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(56)}px;
            font-weight: 900;
            color: {actual_color};
        """)
        
        # Trigger a repaint to draw the updated arc and text.
        self.update()

    def paintEvent(self, event):
        """Overrides the paint event to draw the custom circular gauge."""
        # Create a QPainter for this widget.
        painter = QPainter(self)
        # Enable antialiasing for smooth circles and arcs.
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # --- Drawing Calculations ---
        rect = self.rect()
        center = rect.center()
        
        # Define the dimensions for the gauge.
        pen_width = self.dm.scale(14)    # The thickness of the gauge arc.
        glow_margin = self.dm.scale(18)  # The space for the glow effect.
        # Calculate the radius of the main arc.
        radius = (min(rect.width(), rect.height()) / 2) - pen_width - glow_margin
        
        # --- 1. Draw the Background Glow ---
        # The radius of the glow effect is animated via the glow_opacity property.
        glow_spread = glow_margin * self._glow_opacity
        glow_radius = radius + pen_width / 2 + glow_spread
        glow_color = QColor(self._status_color)
        
        # Create a radial gradient for the glow.
        radial_gradient = QRadialGradient(QPointF(center), glow_radius)
        start_point = (radius + pen_width / 2) / glow_radius
        glow_color.setAlphaF(0.7 * self._glow_opacity) # Animate the alpha.
        radial_gradient.setColorAt(start_point, glow_color) # Glow starts from the edge of the circle.
        radial_gradient.setColorAt(1.0, QColor(0, 0, 0, 0)) # Fades to transparent.
        
        # Draw the glow ellipse.
        painter.setBrush(QBrush(radial_gradient))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(QPointF(center), glow_radius, glow_radius)
        
        # --- 2. Draw the Inner Dark Circle ---
        # This circle provides the background for the text and gauge.
        painter.setBrush(QColor("#2c2c2e"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(QPointF(center), radius + pen_width / 2, radius + pen_width / 2)
        
        # --- 3. Draw the Gauge Background Track ---
        # This is the gray track that the colored arc is drawn on top of.
        pen = QPen(QColor("#333333"), pen_width)
        pen.setCapStyle(Qt.PenCapStyle.RoundCap) # Use round caps for the ends of the arc.
        painter.setPen(pen)
        painter.drawEllipse(QPointF(center), radius, radius)
        
        # --- 4. Draw the Health Arc ---
        # Set the pen color to the current status color.
        pen.setColor(QColor(self._status_color))
        painter.setPen(pen)
        
        # Define the start angle (top of the circle) and the span angle.
        # Angles are in 1/16th of a degree. A full circle is 360 * 16.
        start_angle = 90 * 16
        # The span angle is calculated from the animated health percentage.
        span_angle = -int(self._health_percentage / 100 * 360 * 16)
        
        # Draw the arc over the background track.
        painter.drawArc(
            QRectF(center.x() - radius, center.y() - radius, radius * 2, radius * 2),
            start_angle, span_angle
        )

class BatteryIconWidget(QWidget):
    """
    A custom widget for displaying the battery charge level as a stylized
    battery icon with an animated fill and charging indicator.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, dm: DisplayManager, parent=None):
        super().__init__(parent)
        # --- Initialization ---
        self._charge_percentage = 0
        self._is_charging = False
        self.dm = dm
        self.setMinimumSize(dm.scale(200), dm.scale(100))
        
        # --- Animation Setup ---
        # This animation controls the position of the "shine" effect when charging.
        self._charging_animation_value = 0
        self.charging_animation = QPropertyAnimation(self, b"charging_animation_value")
        self.charging_animation.setDuration(2000)
        self.charging_animation.setStartValue(0)
        self.charging_animation.setEndValue(100)
        self.charging_animation.setLoopCount(-1) # Loop indefinitely while charging.
        self.charging_animation.setEasingCurve(QEasingCurve.Type.InOutSine)

    def setCharge(self, percentage: int, is_charging: bool = False):
        """Public method to update the charge percentage and charging status."""
        self._charge_percentage = max(0, min(100, percentage))
        self._is_charging = is_charging
        
        # Start or stop the charging animation based on the status.
        if is_charging and not self.charging_animation.state() == QPropertyAnimation.State.Running:
            self.charging_animation.start()
        elif not is_charging:
            self.charging_animation.stop()
        
        # Trigger a repaint to show the new state.
        self.update()

    # --- PyQt Properties for Animation ---
    @pyqtProperty(int)
    def charging_animation_value(self):
        return self._charging_animation_value

    @charging_animation_value.setter
    def charging_animation_value(self, value: int):
        self._charging_animation_value = value
        self.update() # Repaint on every animation frame.

    def _get_color_for_percentage(self) -> QColor:
        """Helper method to determine the fill color based on charge level."""
        if self._charge_percentage > 80: return QColor("#00ff00")
        elif self._charge_percentage > 60: return QColor("#7cfc00")
        elif self._charge_percentage > 40: return QColor("#ffeb3b")
        elif self._charge_percentage > 20: return QColor("#ffa500")
        elif self._charge_percentage > 10: return QColor("#ff8c00")
        else: return QColor("#ff0000")

    def paintEvent(self, event):
        """Overrides the paint event to draw the custom battery icon."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # --- Drawing Calculations ---
        charge_color = self._get_color_for_percentage()
        border_thickness = self.dm.scale(4)
        terminal_width = self.dm.scale(10)
        terminal_height = self.dm.scale(35)
        border_radius = self.dm.scale(12)
        
        # --- 1. Draw Battery Outline ---
        # Define the rectangle for the main body of the battery.
        body_rect = QRectF(0, 0, self.width() - terminal_width - border_thickness, self.height())
        # Set a gray pen for the outline.
        pen = QPen(QColor(120, 120, 120), border_thickness)
        painter.setPen(pen)
        painter.setBrush(Qt.BrushStyle.NoBrush) # No fill for the outline.
        painter.drawRoundedRect(body_rect, border_radius, border_radius)
        
        # --- 2. Draw Battery Terminal ---
        # Define the rectangle for the positive terminal.
        terminal_rect = QRectF(
            body_rect.right() + 2, (self.height() - terminal_height) / 2,
            terminal_width, terminal_height
        )
        # Draw the terminal with a solid gray fill.
        painter.setBrush(QBrush(QColor(120, 120, 120)))
        painter.drawRoundedRect(terminal_rect, border_radius / 2, border_radius / 2)
        
        # --- 3. Draw Battery Fill Level ---
        # Calculate the width of the fill rectangle based on the charge percentage.
        fill_width = (body_rect.width() - border_thickness) * (self._charge_percentage / 100.0)
        fill_rect = QRectF(
            body_rect.left() + border_thickness / 2, body_rect.top() + border_thickness / 2,
            fill_width, body_rect.height() - border_thickness
        )
        
        # Create a gradient for the fill to give it a 3D look.
        gradient = QLinearGradient(fill_rect.topLeft(), fill_rect.bottomLeft())
        gradient.setColorAt(0, charge_color.lighter(120))
        gradient.setColorAt(0.5, charge_color)
        gradient.setColorAt(1, charge_color.darker(110))
        
        # Draw the fill rectangle.
        painter.setBrush(QBrush(gradient))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawRoundedRect(fill_rect, border_radius, border_radius)
        
        # --- 4. Draw Charging "Shine" Animation ---
        # If the battery is charging, draw a moving highlight.
        if self._is_charging:
            shine_width = fill_rect.width() * 0.3
            # The X position of the shine is controlled by the animation property.
            shine_x = fill_rect.left() + (fill_rect.width() - shine_width) * (self._charging_animation_value / 100.0)
            
            # Create a gradient for the shine effect.
            shine_gradient = QLinearGradient(shine_x, fill_rect.top(), shine_x + shine_width, fill_rect.top())
            shine_gradient.setColorAt(0, QColor(255, 255, 255, 0))   # Fade in
            shine_gradient.setColorAt(0.5, QColor(255, 255, 255, 100))# Bright center
            shine_gradient.setColorAt(1, QColor(255, 255, 255, 0))   # Fade out
            
            # Draw the shine rectangle over the fill.
            painter.setBrush(QBrush(shine_gradient))
            painter.drawRoundedRect(
                QRectF(shine_x, fill_rect.top(), shine_width, fill_rect.height()),
                border_radius, border_radius
            )
            
        # --- 5. Draw Percentage Text ---
        # Set a large, bold font for the percentage text.
        font = QFont("Segoe UI", self.dm.scale_font_size(36), QFont.Weight.ExtraBold)
        painter.setFont(font)
        
        # Draw a subtle shadow for the text to make it pop.
        painter.setPen(QPen(QColor(0, 0, 0, 100)))
        text_rect_shadow = QRectF(body_rect.left() + 2, body_rect.top() + 2, body_rect.width(), body_rect.height())
        painter.drawText(text_rect_shadow, Qt.AlignmentFlag.AlignCenter, f"{self._charge_percentage}%")
        
        # Draw the main white text.
        painter.setPen(QPen(Qt.GlobalColor.white))
        text_rect = QRectF(body_rect.left(), body_rect.top(), body_rect.width(), body_rect.height())
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, f"{self._charge_percentage}%")

class AnimatedCard(QFrame):
    """
    A custom QFrame that provides a container for other widgets. It features a
    subtle background gradient, a drop shadow, and an animated border that
    glows on mouse hover. PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, parent=None):
        super().__init__(parent)
        # Set an object name for styling via the global stylesheet.
        self.setObjectName("Card")
        
        # Set a default minimum height, scaled by DPI.
        try:
            dm = QApplication.instance().display_manager
            self.setMinimumHeight(dm.scale(220))
        except: # Fallback if display manager isn't available yet.
            self.setMinimumHeight(220)
        
        # --- Colors and Animation Setup ---
        # Define the default and hover colors for the animated border.
        self._default_color = QColor(255, 255, 255, 40)  # Semi-transparent white
        self._hover_color = QColor(0, 255, 0, 200)       # Semi-transparent green
        self._border_color = self._default_color
        
        # Create a QPropertyAnimation to animate the 'borderColor' property.
        self.animation = QPropertyAnimation(self, b"borderColor")
        self.animation.setDuration(300) # Fast, 300ms animation.
        self.animation.setEasingCurve(QEasingCurve.Type.InOutCubic)
        
        # --- Effects ---
        # Create a drop shadow effect for the card.
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(35)
        shadow.setColor(QColor(0, 0, 0, 100))
        shadow.setOffset(0, 6)
        self.setGraphicsEffect(shadow)
        
        # Initialize the border color property.
        self.borderColor = self._default_color

    # --- Event Handlers for Hover Animation ---
    
    def enterEvent(self, event):
        """Called when the mouse cursor enters the widget's area."""
        # Set the animation's end value to the hover color and start it.
        self.animation.setEndValue(self._hover_color)
        self.animation.start()
        super().enterEvent(event)

    def leaveEvent(self, event):
        """Called when the mouse cursor leaves the widget's area."""
        # Set the animation's end value back to the default color and start it.
        self.animation.setEndValue(self._default_color)
        self.animation.start()
        super().leaveEvent(event)

    # --- PyQt Property for Border Color Animation ---
    
    @pyqtProperty(QColor)
    def borderColor(self):
        return self._border_color

    @borderColor.setter
    def borderColor(self, color: QColor):
        """
        This setter is called by the QPropertyAnimation. It updates the border
        color by dynamically setting the stylesheet for the card.
        """
        self._border_color = color
        # Update the stylesheet with the new animated color.
        self.setStyleSheet(f"""
            #Card {{
                background: qlineargradient(
                    spread:pad, x1:0.5, y1:0, x2:0.5, y2:1,
                    stop:0 rgba(50, 50, 50, 0.92),
                    stop:1 rgba(40, 40, 40, 0.97)
                );
                border: 2px solid {color.name(QColor.NameFormat.HexArgb)};
                border-radius: 16px;
            }}
        """)

# ============================================================================
# PART 8
# ============================================================================

class AboutDialog(QDialog):
    """
    A custom, frameless "About" dialog with a modern design, animations, and
    links to the developer's profiles.
    PRESERVED FROM ORIGINAL MAIN.PY.
    """
    def __init__(self, parent, dm: DisplayManager):
        super().__init__(parent)
        # Store the DisplayManager for scaling.
        self.dm = dm
        # Set the window title (though it won't be visible on the frameless window).
        self.setWindowTitle("About Battery-Z")
        # Set the dialog to be modal, blocking interaction with the main window.
        self.setModal(True)
        
        # --- Window Sizing and Styling ---
        # Calculate and set the scaled size of the dialog.
        width = dm.scale(850)
        height = dm.scale(550)
        self.setMinimumSize(width, height)
        # Set a max size to prevent it from becoming too large on huge screens.
        self.setMaximumSize(dm.scale(1100), dm.scale(750))
        
        # Apply a base stylesheet for the dialog's background and border.
        self.setStyleSheet(f"""
            QDialog {{
                background-color: #000000;
                border: 2px solid rgba(0, 255, 0, 0.3);
                border-radius: {dm.scale(15)}px;
            }}
            QLabel {{ color: #ffffff; background-color: transparent; }}
            QScrollArea {{ background-color: transparent; border: none; }}
            QScrollArea > QWidget > QWidget {{ background-color: transparent; }}
        """)
        
        # --- UI Construction and Effects ---
        # Build the internal widgets and layout.
        self._setup_ui()
        
        # Add a soft outer glow effect to the entire dialog window.
        glow = QGraphicsDropShadowEffect(self)
        glow.setBlurRadius(60)
        glow.setColor(QColor(0, 255, 0, 100))
        glow.setOffset(0, 0)
        self.setGraphicsEffect(glow)
    
    def _setup_ui(self):
        """Builds the internal layout and widgets of the About dialog."""
        # Main layout for the dialog.
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Use a QScrollArea to ensure content is viewable on smaller screens.
        scroll = QScrollArea(self)
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        # Create a container widget for the scroll area's content.
        scroll_widget = QWidget()
        scroll.setWidget(scroll_widget)
        
        # Main layout for the content within the scroll area.
        layout = QVBoxLayout(scroll_widget)
        layout.setContentsMargins(
            self.dm.scale(60), self.dm.scale(40), self.dm.scale(60), self.dm.scale(40)
        )
        layout.setSpacing(self.dm.scale(20))
        
        # --- Top Section Layout ---
        top_section = QHBoxLayout()
        top_section.setSpacing(self.dm.scale(50))
        
        # --- Left Column (App Info) ---
        left_column = QVBoxLayout()
        left_column.setSpacing(self.dm.scale(12))
        
        # Add the app logo.
        logo_container = QHBoxLayout()
        logo = LogoWidget(self.dm, size=60)
        logo_container.addWidget(logo)
        logo_container.addStretch()
        left_column.addLayout(logo_container)
        
        # Add the app name label with a glow effect.
        app_name = QLabel(APP_NAME)
        app_name.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(38)}px; font-weight: 900;
            color: #00ff00; letter-spacing: 2px;
        """)
        glow = QGraphicsDropShadowEffect()
        glow.setBlurRadius(30)
        glow.setColor(QColor(0, 255, 0, 200))
        glow.setOffset(0, 0)
        app_name.setGraphicsEffect(glow)
        left_column.addWidget(app_name)
        
        # Add version, subtitle, and description labels.
        version = QLabel(f"Version {APP_VERSION}")
        version.setStyleSheet(f"font-size: {self.dm.scale_font_size(14)}px; color: #bbbbbb; font-weight: 600;")
        left_column.addWidget(version)
        
        subtitle = QLabel(APP_SUBTITLE)
        subtitle.setStyleSheet(f"font-size: {self.dm.scale_font_size(13)}px; color: #888888; font-style: italic;")
        subtitle.setWordWrap(True)
        left_column.addWidget(subtitle)
        
        left_column.addSpacing(self.dm.scale(15))
        
        description = QLabel(
            "Advanced battery health monitoring system with multi-factor "
            "degradation analysis, real-time telemetry, and predictive analytics. "
            "Powered by a robust, multi-layered data retrieval engine and "
            "reliable empirical algorithms for health and lifespan calculation."
        )
        description.setStyleSheet(f"font-size: {self.dm.scale_font_size(12)}px; color: #aaaaaa; line-height: 1.6;")
        description.setWordWrap(True)
        left_column.addWidget(description)
        
        left_column.addSpacing(self.dm.scale(10))
        
        # Add developer information.
        dev_label = QLabel("Developed by")
        dev_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(11)}px; color: #888888;")
        left_column.addWidget(dev_label)
        
        author = QLabel(AUTHOR_NAME)
        author.setStyleSheet(f"font-size: {self.dm.scale_font_size(20)}px; font-weight: 700; color: #00bfff;")
        left_column.addWidget(author)
        
        left_column.addStretch()
        
        # --- Right Column (Links) ---
        right_column = QVBoxLayout()
        right_column.setSpacing(self.dm.scale(15))
        
        follow_label = QLabel("Follow Me")
        follow_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(16)}px; font-weight: 700; color: #ffffff;")
        right_column.addWidget(follow_label)
        
        social_layout = QVBoxLayout()
        social_layout.setSpacing(self.dm.scale(10))
        
        # GitHub button.
        github_btn = QPushButton("🐙 GitHub")
        github_btn.setMinimumSize(self.dm.scale(200), self.dm.scale(45))
        github_btn.setStyleSheet(f"""
            QPushButton {{ background-color: #333333; color: white; border: none;
                           border-radius: {self.dm.scale(10)}px; padding: {self.dm.scale(12)}px;
                           font-weight: 700; font-size: {self.dm.scale_font_size(13)}px; }}
            QPushButton:hover {{ background-color: #444444; border: 2px solid #00ff00; }}
        """)
        github_btn.clicked.connect(lambda: self._open_link(AUTHOR_GITHUB))
        
        # LinkedIn button.
        linkedin_btn = QPushButton("💼 LinkedIn")
        linkedin_btn.setMinimumSize(self.dm.scale(200), self.dm.scale(45))
        linkedin_btn.setStyleSheet(f"""
            QPushButton {{ background-color: #0077b5; color: white; border: none;
                           border-radius: {self.dm.scale(10)}px; padding: {self.dm.scale(12)}px;
                           font-weight: 700; font-size: {self.dm.scale_font_size(13)}px; }}
            QPushButton:hover {{ background-color: #0088cc; border: 2px solid #00ff00; }}
        """)
        linkedin_btn.clicked.connect(lambda: self._open_link(AUTHOR_LINKEDIN))
        
        social_layout.addWidget(github_btn)
        social_layout.addWidget(linkedin_btn)
        right_column.addLayout(social_layout)
        right_column.addSpacing(self.dm.scale(15))
        
        # "Buy Me a Coffee" button.
        support_label = QLabel("Support Development 💖")
        support_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(16)}px; font-weight: 700; color: #ffffff;")
        right_column.addWidget(support_label)
        
        coffee_btn = QPushButton("☕ Buy Me a Coffee")
        coffee_btn.setMinimumSize(self.dm.scale(200), self.dm.scale(50))
        coffee_btn.setStyleSheet(f"""
            QPushButton {{ background-color: #FFDD00; color: #000000; border: none;
                           border-radius: {self.dm.scale(12)}px; padding: {self.dm.scale(14)}px;
                           font-weight: 900; font-size: {self.dm.scale_font_size(15)}px; }}
            QPushButton:hover {{ background-color: #FFE633; border: 3px solid #00ff00; }}
        """)
        coffee_btn.clicked.connect(lambda: self._open_link(BUY_ME_COFFEE))
        right_column.addWidget(coffee_btn)
        
        right_column.addStretch()
        
        # --- Assemble Top Section ---
        top_section.addLayout(left_column, 1)
        top_section.addLayout(right_column, 1)
        layout.addLayout(top_section)
        layout.addSpacing(self.dm.scale(20))
        
        # --- Close Button ---
        close_btn = QPushButton("Close")
        close_btn.setMinimumSize(self.dm.scale(140), self.dm.scale(42))
        close_btn.setStyleSheet(f"""
            QPushButton {{ background-color: #3e3e3e; color: white; border: 2px solid #555555;
                        border-radius: {self.dm.scale(10)}px; padding: {self.dm.scale(10)}px;
                        font-weight: 700; font-size: {self.dm.scale_font_size(14)}px; }}
            QPushButton:hover {{ background-color: #4a4a4a; border-color: #00ff00; }}
        """)
        close_btn.clicked.connect(self.accept) # 'accept' is the standard way to close a QDialog.
        
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_layout.addWidget(close_btn)
        close_layout.addStretch()
        layout.addLayout(close_layout)
        
        # Add the scroll area to the main dialog layout.
        main_layout.addWidget(scroll)

    def _open_link(self, url: str):
        """Safely opens a URL in the user's default web browser."""
        try:
            # Import webbrowser locally to keep it contained.
            import webbrowser
            # Open the URL.
            webbrowser.open(url)
        except Exception as e:
            # Log an error if opening the URL fails.
            logging.error("Failed to open URL '%s': %s", url, e)


# ============================================================================
# SECTION 10: MAIN APPLICATION WINDOW
# Description: This is the main QMainWindow class that orchestrates the entire
#              user interface. Its visual and structural code is preserved
#              identically from the original main.py. The modifications are in
#              the backend logic, connecting UI elements to the new, reliable
#              data from the BatteryIntelligence class.
# ============================================================================

class MainWindow(QMainWindow):
    """
    The main application window. It constructs the UI from various custom
    widgets, connects signals to slots, and handles the presentation of all
    battery data.
    """
    def __init__(self, dm: DisplayManager, battery_data: BatteryData,
                 soh_result: Dict[str, Any], rul_prediction: Dict[str, Any]):
        """
        Initializes the MainWindow.

        Args:
            dm (DisplayManager): The instance for handling DPI scaling.
            battery_data (BatteryData): The fully populated data object from the initial load.
            soh_result (Dict[str, Any]): The dictionary containing the calculated SOH.
            rul_prediction (Dict[str, Any]): The dictionary containing the predicted RUL.
        """
        super().__init__()
        # Store the initial data objects.
        self.dm = dm
        self.battery_data = battery_data
        self.soh_result = soh_result
        self.rul_prediction = rul_prediction
        
        # Set an object name for styling.
        self.setObjectName("MainWindow")
        
        # --- Initialize UI State Variables ---
        self.cards = [] # A list to hold the animated card widgets for startup animation.
        self.basic_labels = {} # A dictionary to easily access the InfoRowWidgets for updates.
        
        # --- Initialize Manager Classes ---
        # The SystemTrayManager handles the icon in the Windows system tray.
        self.tray_manager = SystemTrayManager(self)
        # The NotificationManager handles showing desktop notifications.
        self.notification_manager = NotificationManager(self.tray_manager)
        
        # --- Setup Methods ---
        # Configure the main window properties (title, size, icon).
        self._setup_window()
        # Create the top menu bar (File, About, etc.).
        self._setup_menu_bar()
        # Construct the entire UI layout and all its widgets.
        self._setup_ui()
        
        # --- Initial Data Population ---
        # Take the initial data and populate all the UI fields.
        self._populate_initial_data()
        
        # --- Start Background Tasks ---
        # If a battery is present, start the real-time polling worker thread.
        if battery_data.battery_present:
            self._start_realtime_updates()
        else:
            logging.info("Real-time updates disabled (no battery detected).")
        
        # --- Final UI Steps ---
        # Trigger the startup animations for the cards.
        self._start_startup_animations()
        # Center the window on the screen.
        self.center_on_screen()
        
        # Perform an initial update of the system tray menu with live data.
        if self.tray_manager:
            self.tray_manager.update_battery_info(
                self.battery_data, 
                self.soh_result.get("soh_percentage", 0)
            )
    
    def center_on_screen(self):
        """Centers the main window on the primary display."""
        # Get the available geometry of the screen (excluding taskbar).
        screen_geometry = self.dm.geometry
        # Move the window's center to the screen's center.
        self.move(screen_geometry.center() - self.rect().center())

    def _setup_window(self):
        """Sets up the main window's properties."""
        # Set a different window title if no battery is detected.
        if not self.battery_data.battery_present:
            self.setWindowTitle(f"{APP_NAME} v{APP_VERSION} - No Battery Detected")
        else:
            self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        
        # Set the window size using the DisplayManager for proper scaling.
        self.resize(self.dm.get_optimal_window_size())
        
        # Set the window icon.
        if os.path.exists(LOGO_ICO_PATH):
            try:
                self.setWindowIcon(QIcon(LOGO_ICO_PATH))
                logging.info("Window icon set successfully.")
            except Exception as e:
                logging.error("Failed to set window icon: %s", e)
        
        # Apply the global stylesheet.
        stylesheet_manager = GlobalStylesheet(self.dm)
        self.setStyleSheet(stylesheet_manager.get())
        
        # Set the default font for the application, scaled correctly.
        font = QFont("Segoe UI", self.dm.scale_font_size(10))
        self.setFont(font)

    def _setup_menu_bar(self):
        """Creates and configures the main menu bar."""
        menubar = self.menuBar()
        
        # --- File Menu ---
        file_menu = menubar.addMenu("File")
        
        # Refresh Action
        refresh_action = QAction("🔄 Refresh", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(lambda: self.refresh_application(hard_reset=True))
        file_menu.addAction(refresh_action)
        
        file_menu.addSeparator()
        
        # Custom Cycle Count Action
        custom_cycle_action = QAction("⚙️ Custom Cycle Count", self)
        custom_cycle_action.triggered.connect(self.show_custom_cycle_dialog)
        file_menu.addAction(custom_cycle_action)
        
        file_menu.addSeparator()
        
        # Exit Action
        exit_action = QAction("❌ Exit", self)
        exit_action.setShortcut("Alt+F4")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # --- Direct Actions ---
        # Buy Me a Coffee Action
        coffee_action = menubar.addAction("☕ Buy Me a Coffee")
        coffee_action.triggered.connect(self.open_buy_me_coffee)
        
        # About Action
        about_action = menubar.addAction("About")
        about_action.triggered.connect(self.show_about_dialog)

# ============================================================================
# PART 9
# ============================================================================
    
    def _setup_ui(self):
        """
        Constructs the entire UI layout by creating and arranging all the custom
        widgets. This method is preserved from the original main.py to maintain
        the exact visual layout.
        """
        # --- Main Scroll Area and Container ---
        # A QScrollArea is the main container, allowing the content to be scrollable on smaller screens.
        self.scroll_area = QScrollArea(self)
        self.scroll_area.setObjectName("ScrollArea")
        self.scroll_area.setWidgetResizable(True) # Allows the inner widget to resize with the scroll area.
        self.setCentralWidget(self.scroll_area) # Set this as the main window's central widget.
        
        # This QWidget will contain all other UI elements and will be placed inside the scroll area.
        container_widget = QWidget()
        # Set a minimum size to ensure the layout doesn't collapse on very small window sizes.
        container_widget.setMinimumSize(self.dm.scale(1200), self.dm.scale(850))
        
        # This is the top-level layout for all content.
        main_layout = QVBoxLayout(container_widget)
        # Set generous margins to provide spacing around the content.
        main_layout.setContentsMargins(
            self.dm.scale(40), self.dm.scale(25), self.dm.scale(40), self.dm.scale(30)
        )
        
        # Place the container widget inside the scroll area.
        self.scroll_area.setWidget(container_widget)
        
        # --- Header ---
        # Create the header section (logo, title, welcome message).
        header_layout = self._create_header()
        main_layout.addLayout(header_layout)
        
        # Add vertical spacing between the header and the main grid.
        main_layout.addSpacing(self.dm.scale(25))
        
        # --- Main Grid Layout ---
        # A QGridLayout is used to arrange the info cards in a 2x4 grid.
        self.grid_layout = QGridLayout()
        self.grid_layout.setSpacing(self.dm.scale(28)) # Set spacing between cards.
        main_layout.addLayout(self.grid_layout)
        
        # Create the individual card widgets by calling their respective setup methods.
        print("[UI] Constructing the main 6-card UI layout...")
        self.cards = [
            self._create_battery_basics_card(),
            self._create_charging_state_card(),
            self._create_battery_health_card(),
            self._create_sensors_card(),
            self._create_summary_card(),
            self._create_tips_card()
        ]
        
        # --- Add Cards to the Grid ---
        # The grid layout allows for spanning rows and columns.
        # Card 0 (Basics) spans 2 rows in the first column.
        self.grid_layout.addWidget(self.cards[0], 0, 0, 2, 1)
        # Card 1 (Charging) is in row 0, col 1.
        self.grid_layout.addWidget(self.cards[1], 0, 1)
        # Card 2 (Health) is in row 0, col 2.
        self.grid_layout.addWidget(self.cards[2], 0, 2)
        # Card 3 (Sensors) is in row 0, col 3.
        self.grid_layout.addWidget(self.cards[3], 0, 3)
        # Card 4 (Summary) spans 2 columns in the second row.
        self.grid_layout.addWidget(self.cards[4], 1, 1, 1, 2)
        # Card 5 (Tips) is in row 1, col 3.
        self.grid_layout.addWidget(self.cards[5], 1, 3)
        
        # --- Configure Grid Column Stretching ---
        # Set stretch factors to control how columns resize relative to each other.
        self.grid_layout.setColumnStretch(0, 35) # Basics card column is widest.
        self.grid_layout.setColumnStretch(1, 22)
        self.grid_layout.setColumnStretch(2, 22)
        self.grid_layout.setColumnStretch(3, 21)
        
        print("[✓] UI layout construction complete.")

    def _create_header(self) -> QHBoxLayout:
        """Creates the header section of the UI."""
        header_layout = QHBoxLayout()
        
        logo_title_layout = QHBoxLayout()
        logo_title_layout.setSpacing(self.dm.scale(18))
        
        logo = LogoWidget(self.dm)
        logo_title_layout.addWidget(logo)
        
        title_slogan_layout = QVBoxLayout()
        title_slogan_layout.setSpacing(0)
        
        title_label = QLabel(APP_NAME)
        title_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        title_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(40)}px; font-weight: 900; color: #00ff00; letter-spacing: 1px;")
        
        glow = QGraphicsDropShadowEffect()
        glow.setBlurRadius(self.dm.scale(30))
        glow.setColor(QColor(0, 255, 0, 170))
        glow.setOffset(0, 0)
        title_label.setGraphicsEffect(glow)
        
        slogan_label = QLabel(get_random_quote())
        slogan_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        slogan_label.setStyleSheet(f"color: #bbbbbb; font-size: {self.dm.scale_font_size(15)}px; font-weight: 600; padding-top: {self.dm.scale(6)}px;")
        
        title_slogan_layout.addWidget(title_label)
        title_slogan_layout.addWidget(slogan_label)
        
        logo_title_layout.addLayout(title_slogan_layout)
        header_layout.addLayout(logo_title_layout)
        
        header_layout.addStretch(1)
        
        username = get_user_display_name()
        
        welcome_container = QVBoxLayout()
        welcome_container.setSpacing(self.dm.scale(6))
        
        welcome_label = QLabel(f"Welcome, <span style='color: #00ff00; font-weight:700;'>{username}</span>")
        welcome_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        welcome_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(17)}px; color: #cccccc;")
        
        welcome_container.addWidget(welcome_label)
        header_layout.addLayout(welcome_container)
        
        return header_layout

    def _create_card(self, title: str, title_color: str = "#00bfff") -> Tuple[AnimatedCard, QVBoxLayout]:
        """A factory method to create a styled AnimatedCard with a title."""
        card = AnimatedCard()
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(self.dm.scale(32), self.dm.scale(28), self.dm.scale(32), self.dm.scale(28))
        card_layout.setSpacing(self.dm.scale(16))
        
        card_title = QLabel(title)
        card_title.setObjectName("CardTitle")
        card_title.setStyleSheet(f"color: {title_color}; border-bottom: 2px solid {QColor(title_color).lighter(120).name(QColor.NameFormat.HexArgb)};")
        card_layout.addWidget(card_title)
        
        return card, card_layout

    def _create_battery_basics_card(self) -> AnimatedCard:
        """Creates the 'Battery Basics' card UI."""
        card, layout = self._create_card("Battery Basics", "#00bfff")
        
        grid = QGridLayout()
        grid.setVerticalSpacing(self.dm.scale(14))
        layout.addLayout(grid)
        
        # Define the fields to be displayed in this card.
        basics_fields = [
            ("System", "Unknown"),
            ("Battery Model", "Not Available"),
            ("Manufacturer", "Not Available"),
            ("Serial Number", "Not Available"),
            ("Chemistry", "Unknown"),
            ("Design Capacity", "N/A"),
            ("Full Charge Capacity", "N/A"),
            ("Current Health", "N/A"),
            ("Cycle Count", "N/A"),
            ("Rated Cycle Life", "N/A"),
            ("Battery Age", "Unknown")
        ]
        
        row = 0
        for i, (label_text, value_text) in enumerate(basics_fields):
            # Create a custom InfoRowWidget for each field.
            info_row = InfoRowWidget(label_text, value_text, self.dm)
            # Store a reference to the widget in a dictionary for easy updates.
            self.basic_labels[label_text] = info_row
            grid.addWidget(info_row, row, 0)
            row += 1
            
            # Add a separator line between each row.
            if i < len(basics_fields) - 1:
                grid.addWidget(SeparatorLine(), row, 0)
                row += 1
        
        layout.addStretch(1)
        return card

    def _create_charging_state_card(self) -> AnimatedCard:
        """Creates the 'Charging State' card UI."""
        card, layout = self._create_card("Charging State", "#00bfff")
        
        layout.addStretch(1)
        
        # Add the custom battery icon widget.
        self.battery_icon_widget = BatteryIconWidget(self.dm)
        layout.addWidget(self.battery_icon_widget)
        
        layout.addStretch(1)
        
        # Label for displaying time remaining/to full.
        self.time_display_label = QLabel("Calculating...")
        self.time_display_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.time_display_label.setStyleSheet(f"color: #cccccc; font-size: {self.dm.scale_font_size(16)}px; font-weight: 600;")
        layout.addWidget(self.time_display_label)
        
        # Label for displaying the current status (e.g., "Charging").
        self.charging_status_label = QLabel("Detecting...")
        self.charging_status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.charging_status_label.setStyleSheet(f"font-size: {self.dm.scale_font_size(19)}px; font-weight: 700;")
        layout.addWidget(self.charging_status_label)
        
        return card

    def _create_battery_health_card(self) -> AnimatedCard:
        """Creates the 'Battery Health' card UI."""
        card, layout = self._create_card("Battery Health", "#00bfff")
        
        # Add the custom health circle gauge widget.
        self.health_circle_widget = HealthCircleWidget(self.dm)
        layout.addWidget(self.health_circle_widget)
        
        # Label for displaying detailed cycle count information.
        self.cycle_info_label = QLabel("Calculating cycle count...")
        self.cycle_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.cycle_info_label.setStyleSheet(f"color: #cccccc; font-size: {self.dm.scale_font_size(14)}px;")
        layout.addWidget(self.cycle_info_label)
        
        return card

    def _create_sensors_card(self) -> AnimatedCard:
        """Creates the 'Sensors' card UI."""
        card, layout = self._create_card("Sensors", "#00bfff")
        
        grid = QGridLayout()
        grid.setVerticalSpacing(self.dm.scale(14))
        layout.addLayout(grid)
        
        # Create InfoRowWidgets for each sensor reading.
        self.temp_value_row = InfoRowWidget("Temperature", "N/A", self.dm)
        grid.addWidget(self.temp_value_row, 0, 0)
        
        grid.addWidget(SeparatorLine(), 1, 0)
        
        self.volt_value_row = InfoRowWidget("Voltage", "N/A", self.dm)
        grid.addWidget(self.volt_value_row, 2, 0)
        
        grid.addWidget(SeparatorLine(), 3, 0)
        
        self.power_value_row = InfoRowWidget("Power Draw", "N/A", self.dm)
        grid.addWidget(self.power_value_row, 4, 0)
        
        layout.addStretch(1)
        return card

    def _create_summary_card(self) -> AnimatedCard:
        """Creates the 'Your Battery Summary' card UI."""
        card, layout = self._create_card("Your Battery Summary", "#00ff00")
        
        # A single QLabel that will be populated with a rich-text summary.
        self.summary_text = QLabel("Analyzing your battery health...")
        self.summary_text.setWordWrap(True)
        self.summary_text.setStyleSheet(f"color: #dddddd; line-height: 1.7; font-size: {self.dm.scale_font_size(15)}px;")
        layout.addWidget(self.summary_text)
        
        layout.addStretch(1)
        return card

    def _create_tips_card(self) -> AnimatedCard:
        """Creates the 'Quick Battery Tips' card UI."""
        card, layout = self._create_card("Quick Battery Tips", "#00bfff")
        
        # A list of predefined tips.
        self.tips = BATTERY_TIPS
        self.current_tip_index = 0
        
        # The main label to display the tip content.
        self.tip_content = QLabel(self.tips[self.current_tip_index])
        self.tip_content.setWordWrap(True)
        self.tip_content.setObjectName("TipBox")
        layout.addWidget(self.tip_content)
        
        # --- Tip Navigation ---
        nav_layout = QHBoxLayout()
        nav_layout.setSpacing(self.dm.scale(18))
        
        prev_button = QPushButton("◀")
        prev_button.setObjectName("NavCircleButton")
        prev_button.setFixedSize(self.dm.scale(40), self.dm.scale(40))
        prev_button.clicked.connect(self.show_previous_tip)
        
        self.tip_counter = QLabel(f"{self.current_tip_index + 1} / {len(self.tips)}")
        self.tip_counter.setStyleSheet(f"color: #888888; font-size: {self.dm.scale_font_size(14)}px; font-weight: bold;")
        
        next_button = QPushButton("▶")
        next_button.setObjectName("NavCircleButton")
        next_button.setFixedSize(self.dm.scale(40), self.dm.scale(40))
        next_button.clicked.connect(self.show_next_tip)
        
        nav_layout.addStretch(1)
        nav_layout.addWidget(prev_button)
        nav_layout.addWidget(self.tip_counter)
        nav_layout.addWidget(next_button)
        nav_layout.addStretch(1)
        
        layout.addLayout(nav_layout)
        
        # --- Action Buttons ---
        ai_button = QPushButton("✨ Generate AI Tip")
        ai_button.setObjectName("AiButton")
        ai_button.clicked.connect(self.generate_ai_tip)
        layout.addWidget(ai_button)
        
        history_button = QPushButton("📊 View Battery History")
        history_button.setObjectName("HistoryButton")
        history_button.clicked.connect(self.show_battery_history)
        layout.addWidget(history_button)
        
        return card

    def _populate_initial_data(self):
        """
        Populates all UI widgets with the data that was fetched during the
        splash screen loading phase. This is a critical integration point
        where backend data is passed to the frontend.
        """
        logging.info("Populating UI with initial battery data...")
        
        # Use try-except blocks for each UI section to ensure that an error
        # in one section does not prevent others from loading.
        try:
            # Update the 'Battery Basics' card.
            self._update_battery_basics()
        except Exception as e:
            logging.error("Failed to populate 'Battery Basics' card: %s", traceback.format_exc())
            
        try:
            # Update the 'Charging State' card.
            self._update_charging_state()
        except Exception as e:
            logging.error("Failed to populate 'Charging State' card: %s", traceback.format_exc())
            
        try:
            # Update the 'Battery Health' card.
            self._update_battery_health()
        except Exception as e:
            logging.error("Failed to populate 'Battery Health' card: %s", traceback.format_exc())
            
        try:
            # Update the 'Sensors' card.
            self._update_sensors()
        except Exception as e:
            logging.error("Failed to populate 'Sensors' card: %s", traceback.format_exc())
            
        try:
            # Update the 'Summary' card.
            self._update_summary()
        except Exception as e:
            logging.error("Failed to populate 'Summary' card: %s", traceback.format_exc())
            
        print("[✓] UI population with initial data complete.")

    def _update_battery_basics(self):
        """Updates all the labels in the 'Battery Basics' card."""
        # Create a local reference to the main data object for convenience.
        data = self.battery_data
        
        # Format the system info string.
        system_info = f"{data.laptop_manufacturer or 'Unknown'} {data.laptop_model or 'System'}"
        self.basic_labels["System"].setValue(system_info)
        
        # If no battery is present, populate all fields with "N/A" and return.
        if not data.battery_present:
            self.basic_labels["Battery Model"].setValue("No Battery Detected")
            self.basic_labels["Manufacturer"].setValue("N/A - Desktop PC")
            self.basic_labels["Serial Number"].setValue("N/A")
            self.basic_labels["Chemistry"].setValue("N/A")
            self.basic_labels["Design Capacity"].setValue("N/A")
            self.basic_labels["Full Charge Capacity"].setValue("N/A")
            self.basic_labels["Current Health"].setValue("N/A")
            self.basic_labels["Cycle Count"].setValue("N/A")
            self.basic_labels["Rated Cycle Life"].setValue("N/A")
            self.basic_labels["Battery Age"].setValue("N/A")
            return
        
        # Populate the rest of the fields with data from the BatteryData object.
        # FIX: Use the new helper function to clean the model name before displaying it
        cleaned_model = self._clean_battery_model_name(data.battery_name, data.laptop_manufacturer)
        self.basic_labels["Battery Model"].setValue(cleaned_model)
        self.basic_labels["Manufacturer"].setValue(data.battery_manufacturer or "Unknown")
        self.basic_labels["Serial Number"].setValue(data.battery_serial or "Not Available")
        
        # Use the already normalized chemistry name.
        self.basic_labels["Chemistry"].setValue(data.battery_chemistry or "Unknown")
        
        # Format capacity values with commas and a Wh conversion.
        if data.design_capacity_mwh:
            design_mwh = data.design_capacity_mwh
            self.basic_labels["Design Capacity"].setValue(f"{format_with_commas(design_mwh)} mWh ({design_mwh / 1000:.2f} Wh)")
        else:
            self.basic_labels["Design Capacity"].setValue("Not Available")
            
        if data.full_charge_capacity_mwh:
            fcc_mwh = data.full_charge_capacity_mwh
            self.basic_labels["Full Charge Capacity"].setValue(f"{format_with_commas(fcc_mwh)} mWh ({fcc_mwh / 1000:.2f} Wh)")
        else:
            self.basic_labels["Full Charge Capacity"].setValue("Not Available")
            
        # Update health with color coding.
        health_pct = self.soh_result.get("soh_percentage", 0)
        if health_pct > 0:
            _, status_color = get_health_status(health_pct)
            self.basic_labels["Current Health"].setValue(f"{health_pct:.1f}%")
            self.basic_labels["Current Health"].setStyle(f"color: {status_color}; font-weight: 700;")
        else:
            self.basic_labels["Current Health"].setValue("Calculating...")
            self.basic_labels["Current Health"].setStyle("") # Reset style
            
        # Update cycle count, adding a "(Custom)" note if it's user-defined.
        if data.cycle_count is not None and data.cycle_count >= 0:
            cycle_text = str(data.cycle_count)
            if CUSTOM_CYCLE_COUNT is not None:
                cycle_text += " (Custom)"
            self.basic_labels["Cycle Count"].setValue(cycle_text)
        else:
            self.basic_labels["Cycle Count"].setValue("Not Available")
            
        # Update rated cycle life.
        self.basic_labels["Rated Cycle Life"].setValue(f"{format_with_commas(data.rated_cycle_life)} cycles")
        
        # Estimate and display battery age.
        install_date = get_windows_install_date()
        if install_date:
            age_days = (datetime.datetime.now() - install_date).days
            years = age_days // 365
            months = (age_days % 365) // 30
            age_text = f"~{years}y {months}m" if years > 0 else f"~{months} months"
            self.basic_labels["Battery Age"].setValue(f"{age_text} (OS install)")
        elif data.cycle_count and data.cycle_count > 0:
            # Fallback to estimating age from cycle count.
            age_days = int(data.cycle_count / 0.5) # Assuming 0.5 cycles per day
            years = age_days // 365
            months = (age_days % 365) // 30
            age_text = f"~{years}y {months}m" if years > 0 else f"~{months} months"
            self.basic_labels["Battery Age"].setValue(f"{age_text} (estimated)")
        else:
            self.basic_labels["Battery Age"].setValue("Unknown")
    
    def _clean_battery_model_name(self, model_name: str, manufacturer: str) -> str:
        """
        [NEW] A helper function to clean up concatenated or prefixed battery model
        strings provided by system APIs, ensuring only the model number is shown.
        """
        if not model_name:
            return "Not Available"

        clean_name = model_name

        # Heuristic 1: If the main laptop manufacturer name is a prefix, remove it.
        # This handles cases like "LENOVOL23B4PK4" -> "L23B4PK4"
        if manufacturer and clean_name.upper().startswith(manufacturer.upper()):
            clean_name = clean_name[len(manufacturer):].strip()

        # Heuristic 2: If there are spaces, the last part is almost always the model.
        # This handles cases like "DELL 71R31C7" -> "71R31C7"
        if ' ' in clean_name:
            return clean_name.split(' ')[-1]
        
        # Heuristic 3: Handle known cell manufacturer codes inside the string.
        # This specifically fixes your example: "6007BYDL23B4PK4" -> "L23B4PK4"
        cell_makers = ['BYD', 'SDI', 'LGC', 'PANASONIC', 'SANYO']
        for maker in cell_makers:
            if maker in clean_name:
                # Split the string by the maker code and take the last part
                parts = clean_name.split(maker)
                if len(parts) > 1 and parts[-1]:
                    return parts[-1]

        return clean_name # Return the cleaned (or original) name if no rules apply

# ============================================================================
# PART 10
# ============================================================================

    def _update_charging_state(self):
        """Updates the 'Charging State' card with the latest real-time data."""
        # Create a local reference to the main data object for convenience.
        data = self.battery_data
        
        # Handle the case where no battery is present.
        if not data.battery_present:
            # Set the icon to show 0% and not charging.
            self.battery_icon_widget.setCharge(0, False)
            # Update the status labels to reflect the desktop PC status.
            self.charging_status_label.setText("No Battery")
            self.charging_status_label.setStyleSheet(f"""
                font-size: {self.dm.scale_font_size(19)}px;
                font-weight: 700;
                color: #888888;
            """)
            self.time_display_label.setText("Desktop PC - AC Power Only")
            return
        
        # Get the current percentage and charging status, with defaults.
        percentage = data.current_percentage or 0
        is_charging = data.is_charging or False
        # Update the battery icon widget with the new data.
        self.battery_icon_widget.setCharge(percentage, is_charging)
        
        # Get the dynamic color from the battery icon to use for the status label.
        dynamic_color = self.battery_icon_widget._get_color_for_percentage()
        # Apply the dynamic color to the status label.
        self.charging_status_label.setStyleSheet(f"""
            font-size: {self.dm.scale_font_size(19)}px;
            font-weight: 700;
            color: {dynamic_color.name()};
        """)
        
        # --- Logic to determine the correct status text ---
        # If the battery is currently charging...
        if data.is_charging:
            # Set the status text to "Charging".
            self.charging_status_label.setText("⚡ Charging")
            
            # Check if there's an estimated time to full.
            if data.time_to_full_seconds and data.time_to_full_seconds > 0:
                # Format the seconds into a human-readable string.
                time_str = format_time_duration(data.time_to_full_seconds)
                self.time_display_label.setText(f"{time_str} until full")
            else:
                # If no time is available, provide a contextual message.
                if percentage >= 99:
                    self.time_display_label.setText("Almost Full")
                else:
                    self.time_display_label.setText("Calculating charge time...")
        
        # If the AC adapter is plugged in but the battery is not charging (i.e., it's full).
        elif data.ac_online:
            self.charging_status_label.setText("✓ Fully Charged")
            self.time_display_label.setText("Connected to AC Power")
        
        # If not charging and not on AC power, it must be discharging.
        else:
            self.charging_status_label.setText("🔋 On Battery")
            
            # Check if there's an estimated time to empty.
            if data.time_to_empty_seconds and data.time_to_empty_seconds > 0:
                # Format and display the time remaining.
                time_str = format_time_duration(data.time_to_empty_seconds)
                self.time_display_label.setText(f"{time_str} remaining")
            else:
                self.time_display_label.setText("Calculating time remaining...")

    def _update_battery_health(self, animate: bool = True):
        """
        [MILITARY-GRADE ALGORITHM] This function implements the advanced
        'Tri-Factor Health Matrix' for the main circle display. This complex
        algorithm synthesizes three distinct health vectors to generate a single,
        highly logical, and dynamic score that reacts correctly to all battery
        parameters and user-defined cycle counts.
        """
        data = self.battery_data
        
        if not data.battery_present:
            self.health_circle_widget.setHealth(0, "No Battery", "#888888", animate=False)
            self.cycle_info_label.setText("No battery detected in this system.")
            return

        # --- TRI-FACTOR HEALTH MATRIX ALGORITHM ---
        
        # === VECTOR 1: PHYSICAL CONDITION SCORE (SoH) ===
        # This is the true, physically measured state of the battery, based on
        # capacity fade and cycle wear. It represents the "here and now".
        # This value is now correctly calculated by the fixed `calculate_health` function.
        physical_soh_score = self.soh_result.get("soh_percentage", 0)

        # === VECTOR 2: LIFESPAN UTILIZATION SCORE ===
        # This score is a direct measure of how much of the battery's rated
        # life has been consumed by usage (cycles). It answers "How much have I used?".
        # A non-linear (power of 1.2) curve makes the score drop faster as cycles increase.
        cycle_count = float(data.cycle_count if data.cycle_count is not None else 0)
        rated_cycles = float(data.rated_cycle_life if data.rated_cycle_life > 0 else 1000)
        utilization_ratio = min(cycle_count / rated_cycles, 1.5)
        lifespan_utilization_score = 100 * (1 - (utilization_ratio ** 1.2) * 0.7) # Cap loss at 70% from this factor
        lifespan_utilization_score = max(0.0, min(100.0, lifespan_utilization_score))

        # === VECTOR 3: FUTURE VIABILITY SCORE ===
        # This score represents the remaining lifespan before the battery becomes
        # critically degraded (60% health). It answers "How much useful life is left?".
        # First, we calculate the total predicted lifespan in days until 60% health.
        CRITICAL_THRESHOLD = 60.0
        cycles_per_day = 0.7 # Assume a default if age is unknown
        install_date = get_windows_install_date()
        if install_date and (datetime.datetime.now() - install_date).days > 0:
            age_days = (datetime.datetime.now() - install_date).days
            if age_days > 0 and cycle_count > 0:
                cycles_per_day = cycle_count / age_days
        
        days_to_critical = -1
        if cycles_per_day > 0.01:
            for day in range(1, 7300): # Project up to 20 years
                proj_cycles = cycle_count + (day * cycles_per_day)
                cycle_ratio = proj_cycles / rated_cycles
                # Project SoH using the same formula as the backend
                proj_soh = 100.0 * (1 - (0.2 * (cycle_ratio ** 1.5)))
                if proj_soh < CRITICAL_THRESHOLD:
                    days_to_critical = day
                    break
        
        # Now, calculate the percentage of future viability remaining
        total_lifespan_days = (age_days + days_to_critical) if days_to_critical != -1 else (age_days / (cycle_count / rated_cycles) if cycle_count > 0 else 5475)
        future_viability_score = 100.0
        if total_lifespan_days > 0:
            future_viability_score = (days_to_critical / total_lifespan_days) * 100 if days_to_critical != -1 else 100.0
        future_viability_score = max(0.0, min(100.0, future_viability_score))

        # === FUSION & FINAL SCORE CALCULATION ===
        # We blend the three scores with dynamic weights. As the battery ages (cycle_ratio increases),
        # we put more emphasis on the future viability and less on the raw physical state.
        cycle_ratio_for_weighting = cycle_count / rated_cycles
        
        if cycle_ratio_for_weighting < 0.1: # New Battery
            phys_weight = 0.60
            util_weight = 0.30
            viability_weight = 0.10
        elif cycle_ratio_for_weighting < 0.8: # Mid-Life Battery
            phys_weight = 0.30
            util_weight = 0.40
            viability_weight = 0.30
        else: # Old Battery
            phys_weight = 0.15
            util_weight = 0.35
            viability_weight = 0.50

        holistic_score = (physical_soh_score * phys_weight) + \
                         (lifespan_utilization_score * util_weight) + \
                         (future_viability_score * viability_weight)
        
        final_pct = max(0.0, min(100.0, holistic_score))

        # === ADVANCED STATUS TEXT AND COLOR MAPPING ===
        status_text, status_color = "Unknown", "#888888"
        if final_pct >= 95: status_text, status_color = "Optimal", "#00ff00"
        elif final_pct >= 85: status_text, status_color = "Excellent", "#7cfc00"
        elif final_pct >= 75: status_text, status_color = "Good", "#9a67ea"
        elif final_pct >= 60: status_text, status_color = "Fair", "#ffeb3b" # Using Yellow now
        elif final_pct >= 40: status_text, status_color = "Degraded", "#ffa500"
        elif final_pct >= 20: status_text, status_color = "Poor", "#ff8c00"
        else: status_text, status_color = "Critical", "#ff0000"

        # Update the main health circle with our powerful new score
        self.health_circle_widget.setHealth(final_pct, status_text, status_color, animate=animate)

        # The cycle info label below the circle remains factual and unchanged
        if cycle_count is not None and cycle_count >= 0:
            custom_note = " (custom)" if CUSTOM_CYCLE_COUNT is not None else ""
            usage_pct = (cycle_count / rated_cycles * 100) if rated_cycles > 0 else 0
            self.cycle_info_label.setText(
                f"<strong>{format_with_commas(int(cycle_count))}</strong> of <strong>{format_with_commas(int(rated_cycles))}</strong> rated cycles used{custom_note} "
                f"({usage_pct:.1f}% of rated life)."
            )
        else:
            self.cycle_info_label.setText("Cycle count data not available from your system.")

    def _update_sensors(self):
        """Updates the 'Sensors' card with the latest real-time sensor data."""
        # Create a local reference to the data object.
        data = self.battery_data
        
        # Handle the no-battery case.
        if not data.battery_present:
            self.temp_value_row.setValue("N/A")
            self.volt_value_row.setValue("N/A")
            self.power_value_row.setValue("N/A")
            return
            
        # --- Update Temperature ---
        if data.temperature_celsius is not None:
            temp_c = data.temperature_celsius
            
            # Determine the color for the temperature value based on thresholds.
            if temp_c < 15: temp_color = "#00bfff"      # Cold
            elif 15 <= temp_c < 25: temp_color = "#00ff00"   # Optimal
            elif 25 <= temp_c < 40: temp_color = "#7cfc00"   # Warm
            elif 40 <= temp_c < 50: temp_color = "#ffa500"   # Hot
            elif 50 <= temp_c < 60: temp_color = "#ff8c00"   # Very Hot
            else: temp_color = "#ff0000"                     # Critical
            
            # Set the value and the dynamic style.
            self.temp_value_row.setValue(f"{temp_c:.1f} °C")
            self.temp_value_row.setStyle(f"color: {temp_color}; font-weight: 700; font-size: {self.dm.scale_font_size(15)}px;")
        else:
            self.temp_value_row.setValue("Not Available")
            self.temp_value_row.setStyle("") # Reset style.

        # --- Update Voltage ---
        if data.current_voltage_mv:
            voltage_v = data.current_voltage_mv / 1000.0
            self.volt_value_row.setValue(f"{voltage_v:.2f} V ({format_with_commas(data.current_voltage_mv)} mV)")
            # Use a simple color logic for voltage.
            voltage_color = "#00ff00" if 10.0 < voltage_v < 20.0 else "#ffa500"
            self.volt_value_row.setStyle(f"color: {voltage_color}; font-weight: 600; font-size: {self.dm.scale_font_size(15)}px;")
        elif data.design_voltage_mv:
            # Fallback to showing the design voltage if real-time voltage is unavailable.
            voltage_v = data.design_voltage_mv / 1000.0
            self.volt_value_row.setValue(f"~{voltage_v:.2f} V (rated)")
            self.volt_value_row.setStyle(f"color: #888888; font-size: {self.dm.scale_font_size(15)}px;")
        else:
            self.volt_value_row.setValue("Not Available")
            self.volt_value_row.setStyle("")

        # --- Update Power Draw ---
        if data.power_draw_watts is not None:
            power_w = data.power_draw_watts
            
            # Determine color based on power consumption level.
            if power_w < 5: power_color = "#00ff00"     # Idle
            elif 5 <= power_w < 15: power_color = "#7cfc00"  # Light Use
            elif 15 <= power_w < 30: power_color = "#ffeb3b" # Moderate Use
            elif 30 <= power_w < 50: power_color = "#ffa500" # Heavy Use
            else: power_color = "#ff0000"                    # Gaming/Intensive
            
            # Distinguish between charging and discharging in the label.
            rate_label = "Charge Rate" if data.is_charging else "Power Draw"
            self.power_value_row.label.setText(rate_label)
            self.power_value_row.setValue(f"{power_w:.2f} W")
            self.power_value_row.setStyle(f"color: {power_color}; font-weight: 700; font-size: {self.dm.scale_font_size(15)}px;")
        else:
            self.power_value_row.label.setText("Power Draw")
            self.power_value_row.setValue("Not Available")
            self.power_value_row.setStyle("")

# ============================================================================
    def _update_summary(self):
        """
        [ALGORITHM ENHANCED] Generates a much more advanced, multi-part summary.
        It now includes a secondary forecast to a 'critical' 60% health threshold,
        identifies the user's likely usage profile, and provides tailored advice.
        """
        data = self.battery_data
        soh = self.soh_result.get("soh_percentage", 0)
        rul = self.rul_prediction
        
        if not data.battery_present:
            summary = "🖥️ <strong>Desktop PC Detected</strong><br><br>"
            summary += "<p>This system does not have a battery. This application is monitoring general system information only.</p>"
            self.summary_text.setText(summary)
            return

        if not rul.get("success") or soh <= 0:
            summary = "⚠️ <strong>Analysis Incomplete:</strong><br><br>Could not generate a reliable battery summary due to missing or inconsistent data from your system."
            self.summary_text.setText(summary)
            return

        # --- Generate the Advanced Summary ---
        summary = ""
        health_pct = soh
        _, status_color = get_health_status(health_pct)

        # 1. Headline
        summary += f"<strong style='font-size: 16px; color: {status_color};'>{get_health_status(health_pct)[0]} Condition</strong><br>"
        summary += f"Your battery is operating at <strong>{health_pct:.1f}%</strong> of its original design capacity.<br><br>"
        
        # 2. Primary Lifespan Forecast (to 80%)
        years, months = rul["years"], rul["months"]
        summary += "📊 <strong>Primary Lifespan Forecast:</strong><br>"
        if years > 0 or months > 0:
            summary += f"Based on current trends, you have approximately <strong>{years} years and {months} months</strong> remaining until your battery health degrades to the 80% replacement threshold.<br><br>"
        else:
            summary += "Your battery is at or below the 80% replacement threshold. Reduced runtime and performance are expected.<br><br>"
        
        # 3. NEW: Critical Lifespan Forecast (to 60%)
        install_date = get_windows_install_date()
        age_days = max(1, (datetime.datetime.now() - install_date).days if install_date else int((data.cycle_count or 0) / 0.7) or 1)
        cycles_per_day = (data.cycle_count or 0) / age_days

        days_to_critical = -1
        if cycles_per_day > 0.01 and health_pct > 60:
            for day in range(1, 5475):
                proj_cycles = (data.cycle_count or 0) + (day * cycles_per_day)
                cycle_ratio = proj_cycles / data.rated_cycle_life
                proj_soh = 100.0 * (1 - (0.2 * (cycle_ratio ** 1.5)))
                if proj_soh < 60.0:
                    days_to_critical = day
                    break
        
        if days_to_critical != -1:
            crit_years = days_to_critical // 365
            crit_months = (days_to_critical % 365) // 30
            summary += "📉 <strong>Critical Lifespan Forecast:</strong><br>"
            summary += f"You have roughly <strong>{crit_years} years and {crit_months} months</strong> of total predicted life until the battery reaches a critical 60% health, where runtime will be severely impacted.<br><br>"

        # 4. NEW: User Profile Analysis
        summary += f"💡 <strong>Analysis & Top Recommendation:</strong><br>"
        chem_info = f"Your <strong>{data.battery_chemistry}</strong> battery is rated for ~<strong>{data.rated_cycle_life}</strong> cycles. "
        
        if cycles_per_day > 0.9:
            summary += chem_info + f"Your usage of <strong>{cycles_per_day:.1f} cycles/day</strong> identifies you as a <strong>Mobile Power User</strong>. The primary aging factor is from frequent charge/discharge cycles. "
            summary += "<strong><u>Recommendation:</u></strong> To maximize remaining life, perform shallower cycles (recharge at 30-40% instead of waiting for <10%)."
        elif cycles_per_day < 0.3 and data.ac_online:
            summary += chem_info + "Your usage identifies you as a <strong>Docked User</strong>, with the battery often kept at or near 100% charge. The primary aging factor is 'calendar aging' from high voltage stress. "
            summary += "<strong><u>Recommendation:</u></strong> Your top priority is to enable your laptop's 'Battery Care' mode (in Dell/HP/Lenovo software) to limit the maximum charge to 80%."
        else:
            summary += chem_info + "Your usage appears <strong>Balanced</strong>, with a mix of mobile and docked use. "
            summary += "<strong><u>Recommendation:</u></strong> For you, the most effective strategy is to avoid temperature extremes (e.g., direct sunlight, poor ventilation) which dramatically accelerate degradation."

        self.summary_text.setText(summary)
        
    def _start_realtime_updates(self):
        """
        Initializes and starts the background thread for real-time data polling.
        """
        # Create an instance of the backend class for the worker to use.
        intelligence_instance = BatteryIntelligence()
        
        # Create the thread and the worker.
        self.realtime_thread = QThread()
        self.realtime_worker = RealtimeUpdateWorker(intelligence_instance)
        
        # Move the worker object to the new thread.
        self.realtime_worker.moveToThread(self.realtime_thread)
        
        # --- Connect Signals and Slots ---
        # When the worker emits data_updated, call the _on_realtime_data_update slot.
        self.realtime_worker.data_updated.connect(self._on_realtime_data_update)
        # Connect the alert signals to their respective handler slots.
        self.realtime_worker.high_temp_alert.connect(self._on_high_temp_alert)
        self.realtime_worker.low_battery_alert.connect(self._on_low_battery_alert)
        # When the thread starts, it will call the worker's main polling method.
        self.realtime_thread.started.connect(self.realtime_worker.poll_realtime_data)
        
        # Start the thread.
        self.realtime_thread.start()
        logging.info("Real-time update thread started.")

    def _on_realtime_data_update(self, realtime_data: Dict):
        """
        This is the SLOT that is called on the main UI thread whenever the
        RealtimeUpdateWorker emits new data. It updates the UI with the fresh data.
        """
        try:
            # Update the main BatteryData object with the new real-time values.
            if realtime_data.get('percent') is not None:
                self.battery_data.current_percentage = realtime_data['percent']
            if realtime_data.get('is_charging') is not None:
                self.battery_data.is_charging = realtime_data['is_charging']
            if realtime_data.get('ac_online') is not None:
                self.battery_data.ac_online = realtime_data['ac_online']
            if realtime_data.get('time_remaining') is not None:
                self.battery_data.time_to_empty_seconds = realtime_data['time_remaining']
            if realtime_data.get('temperature_celsius') is not None:
                self.battery_data.temperature_celsius = realtime_data['temperature_celsius']
            if realtime_data.get('voltage_mv') is not None:
                self.battery_data.current_voltage_mv = realtime_data['voltage_mv']
            if realtime_data.get('power_watts') is not None:
                self.battery_data.power_draw_watts = realtime_data['power_watts']
            
            # Call the specific UI update methods for the affected cards.
            self._update_charging_state()
            self._update_sensors()
            
            # Update the system tray icon's context menu as well.
            if self.tray_manager:
                self.tray_manager.update_battery_info(
                    self.battery_data, self.soh_result.get("soh_percentage", 0)
                )
        except Exception as e:
            logging.error("Failed to process real-time UI update: %s", e)

    def _on_high_temp_alert(self, temperature: float):
        """Slot to handle high temperature alerts from the worker thread."""
        logging.warning("Received high temperature alert signal for %.1f°C", temperature)
        self.notification_manager.send_high_temperature_alert(temperature)

    def _on_low_battery_alert(self, percentage: int):
        """Slot to handle low battery alerts from the worker thread."""
        logging.warning("Received low battery alert signal for %d%%", percentage)
        self.notification_manager.send_low_battery_alert(percentage)

    def show_previous_tip(self):
        """Cycles to the previous tip in the 'Quick Battery Tips' card."""
        self.current_tip_index = (self.current_tip_index - 1 + len(self.tips)) % len(self.tips)
        self.tip_content.setText(self.tips[self.current_tip_index])
        self.tip_counter.setText(f"{self.current_tip_index + 1} / {len(self.tips)}")

    def show_next_tip(self):
        """Cycles to the next tip in the 'Quick Battery Tips' card."""
        self.current_tip_index = (self.current_tip_index + 1) % len(self.tips)
        self.tip_content.setText(self.tips[self.current_tip_index])
        self.tip_counter.setText(f"{self.current_tip_index + 1} / {len(self.tips)}")

# ============================================================================
# PART 11
# ============================================================================

    def generate_ai_tip(self):
        """
        Generates a context-aware "AI" tip. In reality, it selects a relevant
        tip from a predefined list based on the current battery status.
        The logic is preserved from the original main.py but updated to use
        the new, reliable data model.
        """
        # Create a local reference to the data object.
        data = self.battery_data
        
        # A list to hold tips that are relevant to the current situation.
        possible_tips = []
        
        # --- Tip Selection Logic ---
        
        # If no battery is present, provide a tip about desktop PCs.
        if not data.battery_present:
            tip = "💻 <strong>Desktop PC Info</strong><br><br>" \
                  "Your system doesn't have a battery. For protection against power outages, " \
                  "consider using an Uninterruptible Power Supply (UPS)."
            possible_tips.append(tip)
        else:
            # If a battery is present, check various conditions to find relevant tips.
            health_pct = self.soh_result.get("soh_percentage", 0)
            temp = data.temperature_celsius
            is_charging = data.is_charging
            percentage = data.current_percentage or 50
            
            # Add tips based on temperature.
            if temp and temp > 55:
                possible_tips.append(f"🔥 <strong>CRITICAL Temperature: {temp:.1f}°C!</strong><br><br>High heat causes permanent damage. Stop intensive tasks and ensure ventilation NOW.")
            elif temp and temp > 45:
                possible_tips.append(f"🌡️ <strong>High Temperature: {temp:.1f}°C.</strong><br><br>Your battery is running hot, which accelerates aging. Try to improve airflow.")
                
            # Add tips based on charging status and level.
            if is_charging and percentage > 90:
                possible_tips.append("⚡ <strong>Optimize Charging:</strong><br><br>Leaving your battery at 100% while plugged in can cause stress. Unplug when full or use a charge-limiting feature.")
            if not is_charging and percentage < 10:
                possible_tips.append("🚨 <strong>CRITICAL Low Battery:</strong><br><br>Charge immediately! Deep discharges below 10% are very stressful for the battery.")
            elif not is_charging and percentage < 20:
                possible_tips.append("🪫 <strong>Low Battery:</strong><br><br>Charge soon to avoid a deep discharge cycle. It's best to keep the battery above 20%.")
            
            # Add tips based on overall health.
            if health_pct > 0 and health_pct < 60:
                possible_tips.append(f"🔧 <strong>Replacement Advisory:</strong><br><br>With health at {health_pct:.1f}%, your battery is significantly degraded. Plan for a replacement for best performance.")
            elif health_pct > 0 and health_pct < 80:
                possible_tips.append(f"⚠️ <strong>Health is Degrading:</strong><br><br>At {health_pct:.1f}%, you'll notice shorter runtimes. Use battery-saving features to maximize its remaining life.")
            elif health_pct >= 90:
                possible_tips.append(f"✅ <strong>Excellent Health:</strong><br><br>At {health_pct:.1f}%, your battery is in great shape. Keep up the good habits!")
        
        # Add a generic tip from the main list as a fallback.
        possible_tips.append(random.choice(BATTERY_TIPS))
        
        # --- Display the Tip ---
        # Select a random tip from the list of relevant options.
        selected_tip = random.choice(possible_tips)
        # Update the UI.
        self.tip_content.setText(selected_tip)
        self.tip_counter.setText("AI Generated")
        logging.info("Generated a new contextual 'AI' tip.")

    def show_battery_history(self):
        """
        [NEW FEATURE] This function, which was a placeholder in the original
        main.py, is now fully implemented. It launches a dialog that shows a
        graph of battery percentage over time.
        """
        # This feature is only available if a battery is present.
        if not self.battery_data.battery_present:
            QMessageBox.information(self, "Feature Not Available", "Battery history is not applicable for desktop PCs without a battery.")
            return

        # Create and execute the HistoryDialog.
        # This will be defined later in the file.
        # try:
        #     dialog = HistoryDialog(self)
        #     dialog.exec_()
        # except NameError:
        # Placeholder for now, as HistoryDialog is defined later.
        QMessageBox.information(self, "Coming Soon", "The history plotting feature is being integrated.")

    def show_about_dialog(self):
        """Creates and shows the custom 'About' dialog."""
        dialog = AboutDialog(self, self.dm)
        dialog.exec_()

    def open_buy_me_coffee(self):
        """Opens the developer's support link in the default web browser."""
        try:
            import webbrowser
            webbrowser.open(BUY_ME_COFFEE)
        except Exception as e:
            logging.error("Failed to open support URL: %s", e)

    def show_custom_cycle_dialog(self):
        """
        Allows the user to manually input a cycle count for testing or calibration.
        This will re-run the analysis with the new value.
        """
        if not self.battery_data.battery_present:
            QMessageBox.information(self, "Not Applicable", "Custom cycle count can only be set for systems with a battery.")
            return
        
        current_cycles = self.battery_data.cycle_count or 0
        
        # Use a QInputDialog to get an integer from the user.
        cycles, ok = QInputDialog.getInt(
            self, "Custom Cycle Count",
            f"Enter a custom cycle count for analysis (0 to reset):\n\nCurrent (detected): {current_cycles}",
            value=current_cycles, min=0, max=10000
        )
        
        # If the user clicked OK and provided a value...
        if ok:
            global CUSTOM_CYCLE_COUNT
            # A value of 0 is used to reset to the auto-detected value.
            if cycles == 0:
                CUSTOM_CYCLE_COUNT = None
                logging.info("Custom cycle count reset. Reverting to auto-detected value.")
            else:
                CUSTOM_CYCLE_COUNT = cycles
                logging.info("Custom cycle count set to %d.", cycles)
            
            # Immediately refresh all application data to reflect the change.
            self.refresh_application("Custom cycle count applied.", hard_reset=False)

# ============================================================================
    def refresh_application(self, completion_message: str = "All data has been refreshed and custom settings reset.", hard_reset: bool = False):
        """
        [ALGORITHM FIXED] Re-runs the data pipeline. Now accepts a 'hard_reset'
        flag. If True (from F5/menu), it clears custom settings. If False
        (from a dialog), it uses the custom settings for the analysis.
        """
        logging.info(f"Application refresh requested. Hard Reset: {hard_reset}")

        # --- FIX: Only reset custom cycle count on a hard reset ---
        if hard_reset:
            global CUSTOM_CYCLE_COUNT
            CUSTOM_CYCLE_COUNT = None
            logging.info("Hard reset: Custom cycle count has been cleared.")

        status_msg = QLabel("Refreshing all battery data...\nPlease wait...")
        status_msg.setAlignment(Qt.AlignmentFlag.AlignCenter)
        status_msg.setStyleSheet(f"""
            background-color: #00bfff; color: white; padding: {self.dm.scale(20)}px;
            border-radius: {self.dm.scale(10)}px; font-weight: 700;
            font-size: {self.dm.scale_font_size(16)}px;
        """)
        status_msg.setWindowFlags(Qt.WindowType.SplashScreen | Qt.WindowType.FramelessWindowHint)
        status_msg.setFixedSize(self.dm.scale(350), self.dm.scale(100))
        status_msg.move(self.geometry().center() - status_msg.rect().center())
        status_msg.show()
        QApplication.processEvents()
        
        try:
            intelligence = BatteryIntelligence()
            # The get_all_data() function will automatically use CUSTOM_CYCLE_COUNT if it's set
            self.battery_data = intelligence.get_all_data()
            
            self.soh_result['soh_percentage'] = intelligence.calculate_health(self.battery_data)
            
            self.rul_prediction = intelligence.estimate_remaining_life(self.battery_data)
            self.rul_prediction['success'] = True
            self.rul_prediction['is_desktop'] = not self.battery_data.battery_present
            self.rul_prediction['cycle_count_available'] = self.battery_data.cycle_count is not None
            self.rul_prediction['cycles_used'] = self.battery_data.cycle_count
            self.rul_prediction['rated_cycles'] = self.battery_data.rated_cycle_life
            self.rul_prediction['using_custom_cycles'] = (CUSTOM_CYCLE_COUNT is not None)

            self._populate_initial_data()
            
            status_msg.close()
            
            logging.info("Application refresh complete.")
            msg_text = completion_message if isinstance(completion_message, str) else "Battery data has been refreshed successfully!"
            QMessageBox.information(self, "Refresh Complete", msg_text)
            
        except Exception as e:
            status_msg.close()
            logging.error("Failed to refresh application data: %s", traceback.format_exc())
            QMessageBox.critical(self, "Refresh Failed", f"An error occurred while refreshing: {e}")

    def _start_startup_animations(self):
        """Starts the fade-in animations for the UI cards on startup."""
        # Iterate through the card widgets.
        for i, card in enumerate(self.cards):
            # Make the card initially transparent.
            card.setWindowOpacity(0.0)
            
            # Create a property animation for the 'windowOpacity' property.
            anim = QPropertyAnimation(card, b"windowOpacity")
            anim.setDuration(700)
            anim.setStartValue(0.0)
            anim.setEndValue(1.0)
            anim.setEasingCurve(QEasingCurve.Type.OutCubic)
            
            # Use QTimer.singleShot to stagger the start of each card's animation.
            QTimer.singleShot(120 + i * 100, anim.start)

    def closeEvent(self, event):
        """
        Overrides the default close event to ensure a clean shutdown, especially
        for the background threads.
        """
        logging.info("Close event triggered. Shutting down application.")
        
        # --- Stop Background Threads ---
        # Check if the realtime worker and thread exist.
        if hasattr(self, 'realtime_worker'):
            try:
                # Tell the worker to stop its loop.
                self.realtime_worker.stop()
                logging.info("Real-time worker stop signal sent.")
            except Exception as e:
                logging.error("Error stopping real-time worker: %s", e)
        
        if hasattr(self, 'realtime_thread'):
            try:
                # Tell the thread to quit its event loop.
                self.realtime_thread.quit()
                # Wait for the thread to finish cleanly (with a 2-second timeout).
                if not self.realtime_thread.wait(2000):
                    # If it doesn't finish in time, forcefully terminate it.
                    logging.warning("Real-time thread timed out. Terminating.")
                    self.realtime_thread.terminate()
                else:
                    logging.info("Real-time thread terminated gracefully.")
            except Exception as e:
                logging.error("Error cleaning up real-time thread: %s", e)
        
        # --- Cleanup UI Elements ---
        # Clean up the system tray icon.
        if hasattr(self, 'tray_manager') and self.tray_manager:
            try:
                self.tray_manager.cleanup()
                logging.info("System tray icon cleaned up.")
            except Exception as e:
                logging.error("Error cleaning up tray icon: %s", e)
        
        logging.info("Shutdown sequence complete. Exiting.")
        # Accept the close event to allow the window to close.
        event.accept()


# ============================================================================
# SECTION 11: MAIN EXECUTION BLOCK
# Description: This is the entry point of the application. It handles command-
#              line arguments, initializes the Qt application, runs the splash
#              screen and data worker, and launches the main window.
# ============================================================================

def main():
    """
    The main entry point for the GUI application.
    """
    # Use a try-except block to catch all unhandled exceptions for a graceful exit.
    try:
        logging.info("========================================================")
        logging.info("Battery-Z Application Starting...")
        logging.info("========================================================")
        
        # --- Step 1: Initialize QApplication ---
        app = QApplication(sys.argv)
        app.setApplicationName(APP_NAME)
        app.setApplicationVersion(APP_VERSION)
        
        # Set the window icon for the entire application.
        if os.path.exists(LOGO_ICO_PATH):
            app.setWindowIcon(QIcon(LOGO_ICO_PATH))
        
        # Enable High DPI scaling attributes.
        app.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
        app.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)
        
        # --- Step 2: Initialize Display Manager and Splash Screen ---
        primary_screen = app.primaryScreen()
        display_manager = DisplayManager(primary_screen)
        # Store the display manager on the app instance for global access.
        app.display_manager = display_manager
        
        splash = SplashScreen(display_manager)
        splash.show()
        
        # --- Step 3: Run Data Fetching in Background Thread ---
        # This is the core of the optimized startup sequence. The heavy data lifting
        # happens here while the user sees the animated splash screen.
        fetch_thread = QThread()
        fetch_worker = BatteryDataWorker()
        fetch_worker.moveToThread(fetch_thread)
        
        # Use lists as mutable containers to store results from the thread.
        battery_data_result = [None]
        fetch_error = [None]
        fetch_complete = [False]
        
        # --- Define Slots to handle signals from the worker ---
        def on_fetch_progress(message, progress):
            """Updates the splash screen."""
            splash.update_progress(message, progress)
        
        def on_fetch_finished(battery_data):
            """Stores the result when fetching is complete."""
            battery_data_result[0] = battery_data
            fetch_complete[0] = True
        
        def on_fetch_error(error_message):
            """Stores an error message if fetching fails."""
            fetch_error[0] = error_message
            fetch_complete[0] = True
            
        # Connect the worker's signals to the handler slots.
        fetch_worker.progress.connect(on_fetch_progress)
        fetch_worker.finished.connect(on_fetch_finished)
        fetch_worker.error.connect(on_fetch_error)
        fetch_thread.started.connect(fetch_worker.fetch_data)
        
        # Start the thread. This will trigger the fetch_data method.
        fetch_thread.start()
        
        # --- Event Loop while loading ---
        # Keep the UI responsive while the worker thread is running.
        while not fetch_complete[0]:
            app.processEvents()
            time.sleep(0.05) # Small sleep to prevent a busy-wait loop.
            
        # Clean up the fetcher thread.
        fetch_thread.quit()
        fetch_thread.wait()
        
        # Check if an error occurred.
        if fetch_error[0]:
            splash.close()
            QMessageBox.critical(None, "Fatal Error", f"Could not start application:\n{fetch_error[0]}")
            return 1
            
        battery_data = battery_data_result[0]
        
        # --- Step 4: Perform Final Calculations ---
        # Now that we have the data, we can do the final calculations before showing the main window.
        splash.update_progress("Calculating health...", 90)
        intelligence = BatteryIntelligence()
        health_percent = intelligence.calculate_health(battery_data)
        soh_result = {"soh_percentage": health_percent}
        
        splash.update_progress("Predicting lifespan...", 95)
        rul_prediction = intelligence.estimate_remaining_life(battery_data)
        # Add extra fields to the RUL dict to match the UI's expected format.
        rul_prediction['success'] = True
        rul_prediction['is_desktop'] = not battery_data.battery_present
        rul_prediction['cycle_count_available'] = battery_data.cycle_count is not None
        rul_prediction['cycles_used'] = battery_data.cycle_count
        rul_prediction['rated_cycles'] = battery_data.rated_cycle_life
        
        # --- Step 5: Launch Main Window ---
        splash.update_progress("Launching UI...", 98)
        main_window = MainWindow(display_manager, battery_data, soh_result, rul_prediction)
        
        # Fade out the splash screen and show the main window.
        splash.finish_with_fade(main_window)
        main_window.show()
        
        logging.info("Application startup successful. Main window is now visible.")
        
        # Start the main application event loop.
        exit_code = app.exec_()
        logging.info("Application exited with code %d.", exit_code)
        return exit_code
        
    except Exception as e:
        # Global exception handler to catch any unhandled errors.
        logging.critical("An unhandled exception occurred: %s", traceback.format_exc())
        QMessageBox.critical(None, "Unhandled Exception", f"A fatal error occurred:\n{e}")
        return 1

# --- Command-Line Argument Handling ---

def print_version():
    """Prints the application's version and build info to the console."""
    print(f"\n{APP_NAME} Version {APP_VERSION}")
    print(f"Build Date: {BUILD_DATE}")
    print(f"Author: {AUTHOR_NAME}\n")

def print_help():
    """Prints the command-line help message."""
    print(f"\n{APP_NAME} v{APP_VERSION} - {APP_SUBTITLE}\n")
    print("USAGE: python main.py [OPTIONS]\n")
    print("OPTIONS:")
    print("  --help, -h       Show this help message.")
    print("  --version, -v    Show version information.")
    print("  --no-gui         Run in console-only mode for a quick data dump.")

def run_console_mode():
    """
    Runs the application in a command-line only mode, printing a summary
    of battery data without launching the GUI.
    """
    print("\n" + "="*80)
    print(f"{APP_NAME} v{APP_VERSION} - CONSOLE MODE")
    print("="*80 + "\n")
    
    # Initialize the backend.
    intelligence = BatteryIntelligence()
    
    print("[→] Fetching all battery data...")
    battery_data = intelligence.get_all_data()
    
    if not battery_data.battery_present:
        print("\n--- NO BATTERY DETECTED ---")
        print("This appears to be a desktop PC or the battery is not connected.")
        print("\n" + "="*80 + "\n")
        return
        
    print("[→] Calculating health and RUL...")
    health = intelligence.calculate_health(battery_data)
    rul = intelligence.estimate_remaining_life(battery_data)
    
    print("\n" + "="*80)
    print("BATTERY ANALYSIS REPORT")
    print("="*80)
    
    print("\n--- SYSTEM & BATTERY ---")
    print(f"  System:         {battery_data.laptop_manufacturer} {battery_data.laptop_model}")
    print(f"  Battery Model:  {battery_data.battery_name}")
    print(f"  Manufacturer:   {battery_data.battery_manufacturer}")
    print(f"  Serial:         {battery_data.battery_serial}")
    print(f"  Chemistry:      {battery_data.battery_chemistry}")
    
    print("\n--- HEALTH & CAPACITY ---")
    print(f"  Health:         {health:.1f}%")
    print(f"  Design Capacity:{format_with_commas(battery_data.design_capacity_mwh)} mWh")
    print(f"  Full Charge:    {format_with_commas(battery_data.full_charge_capacity_mwh)} mWh")
    
    print("\n--- USAGE & LIFESPAN ---")
    print(f"  Cycle Count:    {battery_data.cycle_count} / {battery_data.rated_cycle_life}")
    print(f"  Est. Lifespan:  {rul['years']} years, {rul['months']} months")
    
    print("\n--- REAL-TIME STATUS ---")
    print(f"  Charge Level:   {battery_data.current_percentage}%")
    print(f"  Status:         {'Charging' if battery_data.is_charging else 'Discharging' if not battery_data.ac_online else 'Full'}")
    print(f"  Power Draw:     {battery_data.power_draw_watts or 'N/A'} W")
    print(f"  Temperature:    {battery_data.temperature_celsius or 'N/A'} °C")
    
    print("\n" + "="*80 + "\n")

# This is the main entry point when the script is executed.
if __name__ == "__main__":
    # Get command-line arguments, excluding the script name itself.
    args = sys.argv[1:]
    
    # Check for help or version flags.
    if "--help" in args or "-h" in args:
        print_help()
        sys.exit(0)
    
    if "--version" in args or "-v" in args:
        print_version()
        sys.exit(0)
    
    # Check for the no-gui flag.
    if "--no-gui" in args:
        run_console_mode()
        sys.exit(0)
    
    # If no flags are present, run the main GUI application.
    # sys.exit() ensures the application's exit code is passed to the shell.
    sys.exit(main())
# ======================================================================================================================
# ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
# ======================================================================================================================
