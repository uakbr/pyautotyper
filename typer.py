#!/usr/bin/env python3
"""
VM Keystroke Simulator
A simple Python application that runs on the host machine and sends keystrokes 
to a virtual machine without leaving any traces on the VM.
"""
import os
import sys
import time
import random
import logging
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, filedialog, messagebox
import platform
import pyautogui
import keyboard
import re
import subprocess

# -----------------------------------------------------------------------------
# Configuration and Setup
# -----------------------------------------------------------------------------

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("keystroke_simulator")

# Configure PyAutoGUI
pyautogui.PAUSE = 0.01  # Short pause between PyAutoGUI commands
pyautogui.FAILSAFE = True  # Move mouse to upper-left to abort

# -----------------------------------------------------------------------------
# Utility Functions
# -----------------------------------------------------------------------------

def get_platform_info():
    """
    Get information about the current platform.
    
    Returns:
        dict: Platform information
    """
    return {
        "system": platform.system(),
        "release": platform.release(),
        "version": platform.version(),
        "machine": platform.machine(),
        "processor": platform.processor()
    }

def estimate_typing_time(text, speed_factor):
    """
    Estimate the time needed to type the given text.
    
    Args:
        text (str): The text to be typed
        speed_factor (float): Speed factor from 1 (slow) to 10 (fast)
        
    Returns:
        float: Estimated time in seconds
    """
    char_count = len(text)
    
    # Approximate base typing speed based on speed factor
    base_speed = {
        1: 0.300,  # Very slow
        2: 0.250,
        3: 0.200,
        4: 0.150,
        5: 0.120,  # Medium
        6: 0.100,
        7: 0.080,
        8: 0.060,
        9: 0.040,
        10: 0.020  # Very fast
    }
    
    base_delay = base_speed.get(int(speed_factor), 0.120)
    
    # Rough estimate with some adjustments for special characters
    return char_count * base_delay * 1.1  # Add 10% for special characters and pauses

def clear_clipboard():
    """Clear the system clipboard to avoid leaving sensitive data."""
    try:
        import pyperclip
        pyperclip.copy('')
        logger.debug("Clipboard cleared")
    except ImportError:
        logger.warning("pyperclip not installed, cannot clear clipboard")
        pass

def secure_random():
    """
    Generate a cryptographically secure random number between 0 and 1.
    
    Returns:
        float: Random number between 0 and 1
    """
    try:
        import secrets
        return secrets.randbelow(1000) / 1000.0
    except ImportError:
        # Fall back to less secure but still acceptable random
        return random.random()

def get_active_window_title():
    """Get the title of the currently active window on macOS."""
    try:
        if platform.system() == 'Darwin':  # macOS
            script = 'tell application "System Events" to get name of first application process whose frontmost is true'
            result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
            return result.stdout.strip()
        else:
            # Default fallback for other platforms
            return None
    except Exception as e:
        logger.warning(f"Could not determine active window: {e}")
        return None

# -----------------------------------------------------------------------------
# VM Integration Functions
# -----------------------------------------------------------------------------

def check_vm_focus(vm_window_patterns=None):
    """
    Check if the VM window is in focus based on window title.
    
    Args:
        vm_window_patterns (list): List of strings or regex patterns that match VM window titles
        
    Returns:
        bool: True if the VM appears to be in focus
    """
    if not vm_window_patterns:
        # If no patterns provided, rely on user verification
        return True
        
    current_window = get_active_window_title()
    if not current_window:
        return False
        
    # Check if current window matches any of the VM patterns
    for pattern in vm_window_patterns:
        if re.search(pattern, current_window, re.IGNORECASE):
            return True
            
    return False

def send_key(key, max_retries=3):
    """
    Send a single key press to the VM with retry capability.
    
    Args:
        key (str): The key to press
        max_retries (int): Maximum number of retry attempts
    
    Returns:
        bool: True if successful, False if failed after retries
    """
    retries = 0
    last_error = None
    
    while retries <= max_retries:
        try:
            pyautogui.press(key)
            logger.debug(f"Sent key: {key}")
            return True
        except Exception as e:
            last_error = e
            retries += 1
            logger.warning(f"Error sending key '{key}' (attempt {retries}/{max_retries}): {str(e)}")
            time.sleep(0.2)  # Small delay before retry
    
    logger.error(f"Failed to send key '{key}' after {max_retries} attempts: {str(last_error)}")
    return False

def send_key_combination(modifier, key, max_retries=3):
    """
    Send a key combination (e.g., Shift+1 for !) to the VM with retry capability.
    
    Args:
        modifier (str): The modifier key (e.g., 'shift', 'ctrl')
        key (str): The key to press with the modifier
        max_retries (int): Maximum number of retry attempts
    
    Returns:
        bool: True if successful, False if failed after retries
    """
    retries = 0
    last_error = None
    
    while retries <= max_retries:
        try:
            pyautogui.hotkey(modifier, key)
            logger.debug(f"Sent key combination: {modifier}+{key}")
            return True
        except Exception as e:
            last_error = e
            retries += 1
            logger.warning(f"Error sending key combination '{modifier}+{key}' (attempt {retries}/{max_retries}): {str(e)}")
            time.sleep(0.2)  # Small delay before retry
    
    logger.error(f"Failed to send key combination '{modifier}+{key}' after {max_retries} attempts: {str(last_error)}")
    return False

def send_text(text, max_retries=3):
    """
    Send a text string to the VM all at once with retry capability.
    Note: This is faster but less human-like than character-by-character.
    
    Args:
        text (str): The text to type
        max_retries (int): Maximum number of retry attempts
    
    Returns:
        bool: True if successful, False if failed after retries
    """
    retries = 0
    last_error = None
    
    while retries <= max_retries:
        try:
            pyautogui.write(text)
            logger.debug(f"Sent text: '{text[:10]}...' ({len(text)} chars)")
            return True
        except Exception as e:
            last_error = e
            retries += 1
            logger.warning(f"Error sending text (attempt {retries}/{max_retries}): {str(e)}")
            time.sleep(0.2)  # Small delay before retry
    
    logger.error(f"Failed to send text after {max_retries} attempts: {str(last_error)}")
    return False

# -----------------------------------------------------------------------------
# Timing Functions
# -----------------------------------------------------------------------------

# Base typing speed in seconds per character
BASE_SPEED = {
    1: 0.300,  # Very slow
    2: 0.250,
    3: 0.200,
    4: 0.150,
    5: 0.120,  # Medium
    6: 0.100,
    7: 0.080,
    8: 0.060,
    9: 0.040,
    10: 0.020  # Very fast
}

# Character groups that affect typing speed
EASY_CHARS = set('etaoinsrhldcu ')
MEDIUM_CHARS = set('mfpgwybvkj')
HARD_CHARS = set('xqz')
PUNCTUATION = set('.,;:\'"-!?')
SPECIAL_CHARS = set('@#$%^&*()_+{}|:<>?~')

# Characters that are supported directly by PyAutoGUI and our engine
SUPPORTED_CHARS = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')
SUPPORTED_CHARS.update(' \t\n')  # Add space, tab, and newline
SUPPORTED_CHARS.update(PUNCTUATION)
SUPPORTED_CHARS.update(SPECIAL_CHARS)

def calculate_delay(char, speed_factor):
    """
    Calculate the delay before typing the next character.
    
    Args:
        char (str): The character to type
        speed_factor (float): Speed factor from 1 (slow) to 10 (fast)
        
    Returns:
        float: Delay in seconds
    """
    # Get the base speed from the speed factor
    base_delay = BASE_SPEED.get(int(speed_factor), 0.120)
    
    # Apply character-specific adjustments
    if char in EASY_CHARS:
        delay_factor = 0.9
    elif char in MEDIUM_CHARS:
        delay_factor = 1.0
    elif char in HARD_CHARS:
        delay_factor = 1.2
    elif char in PUNCTUATION:
        delay_factor = 1.1
    elif char in SPECIAL_CHARS:
        delay_factor = 1.3
    elif char == '\n':  # Newline (Enter)
        delay_factor = 1.5
    elif char == '\t':  # Tab
        delay_factor = 1.2
    else:
        delay_factor = 1.0
    
    # Calculate final delay
    final_delay = base_delay * delay_factor
    
    logger.debug(f"Character '{char}': base_delay={base_delay}, factor={delay_factor}, final={final_delay}")
    return final_delay

def add_human_variance(base_delay):
    """
    Add human-like variance to typing timing.
    
    Args:
        base_delay (float): Base delay in seconds
        
    Returns:
        float: Delay with added human-like variance
    """
    # Add random variance between -15% and +20%
    variance_factor = random.uniform(0.85, 1.20)
    return base_delay * variance_factor

def is_supported_character(char):
    """
    Check if a character is supported for typing.
    
    Args:
        char (str): The character to check
        
    Returns:
        bool: True if supported, False otherwise
    """
    # Basic ASCII characters, punctuation, and our defined special chars are supported
    if char in SUPPORTED_CHARS:
        return True
    
    # Standard control characters might be supported
    if ord(char) < 32:  # ASCII control characters
        return False
        
    # Non-printable or exotic unicode might cause issues
    if ord(char) > 127:
        return False
        
    return False

def validate_text(text):
    """
    Validate if all characters in the text can be typed reliably.
    
    Args:
        text (str): The text to validate
        
    Returns:
        tuple: (is_valid, unsupported_chars, positions)
            is_valid (bool): True if all characters are supported
            unsupported_chars (set): Set of unsupported characters
            positions (list): List of (position, char) tuples for unsupported characters
    """
    unsupported_chars = set()
    positions = []
    
    for i, char in enumerate(text):
        if not is_supported_character(char):
            unsupported_chars.add(char)
            positions.append((i, char))
    
    return len(unsupported_chars) == 0, unsupported_chars, positions

# -----------------------------------------------------------------------------
# Keystroke Engine
# -----------------------------------------------------------------------------

class KeystrokeEngine:
    """Engine for handling keystroke simulation with natural timing."""
    
    def __init__(self):
        """Initialize the keystroke engine."""
        self.special_chars = {
            # Map of special characters to their key combinations
            '!': ('shift', '1'),
            '@': ('shift', '2'),
            '#': ('shift', '3'),
            '$': ('shift', '4'),
            '%': ('shift', '5'),
            '^': ('shift', '6'),
            '&': ('shift', '7'),
            '*': ('shift', '8'),
            '(': ('shift', '9'),
            ')': ('shift', '0'),
            '_': ('shift', '-'),
            '+': ('shift', '='),
            '{': ('shift', '['),
            '}': ('shift', ']'),
            '|': ('shift', '\\'),
            ':': ('shift', ';'),
            '"': ('shift', "'"),
            '<': ('shift', ','),
            '>': ('shift', '.'),
            '?': ('shift', '/'),
            '~': ('shift', '`'),
        }
        # Keep track of failed keys
        self.failed_keys = []
    
    def type_text(self, text, speed_factor, stop_event, pause_event):
        """
        Type the given text at the specified speed with natural timing.
        
        Args:
            text (str): The text to type
            speed_factor (float): Speed factor from 1 (slow) to 10 (fast)
            stop_event (threading.Event): Event to signal stopping
            pause_event (threading.Event): Event to signal pausing
            
        Yields:
            tuple: (progress_fraction, current_character, success)
        """
        total_chars = len(text)
        chars_typed = 0
        self.failed_keys = []  # Reset failed keys tracking
        
        # Main typing loop
        for i, char in enumerate(text):
            # Check if we should stop
            if stop_event.is_set():
                logger.info("Typing stopped")
                break
            
            # Check if we should pause
            while pause_event.is_set() and not stop_event.is_set():
                time.sleep(0.1)  # Sleep briefly while paused
            
            # Calculate delay based on character and context
            delay = calculate_delay(char, speed_factor)
            
            # Add human-like variance to timing
            actual_delay = add_human_variance(delay)
            
            # Type the character
            success = False
            try:
                if char == '\n':
                    success = send_key('enter')
                elif char == '\t':
                    success = send_key('tab')
                elif char in self.special_chars:
                    modifier, key = self.special_chars[char]
                    success = send_key_combination(modifier, key)
                else:
                    success = send_key(char)
                
                # Track failed keys
                if not success:
                    self.failed_keys.append((i, char))
                
                chars_typed += 1
                progress = chars_typed / total_chars
                
                # Wait before typing the next character
                time.sleep(actual_delay)
                
                # Yield progress, current character, and success status
                yield progress, char, success
            
            except Exception as e:
                logger.error(f"Error typing character '{char}': {str(e)}")
                self.failed_keys.append((i, char))
                yield chars_typed / total_chars, char, False
        
        logger.info(f"Typing completed with {len(self.failed_keys)} failed keys")

# -----------------------------------------------------------------------------
# GUI Implementation
# -----------------------------------------------------------------------------

class KeystrokeSimulatorGUI:
    """Main GUI class for the Keystroke Simulator."""
    
    def __init__(self):
        """Initialize the GUI components."""
        self.root = tk.Tk()
        self.root.title("VM Keystroke Simulator")
        self.root.geometry("600x550")  # Increased height for new VM focus controls
        self.root.minsize(500, 450)
        
        # Make window always on top
        self.root.attributes('-topmost', True)
        
        # Create a flag for stopping the typing process
        self.stop_typing = threading.Event()
        self.paused = threading.Event()
        self.typing_thread = None
        self.keystroke_engine = KeystrokeEngine()
        
        # Set up the emergency stop hotkey (Esc) - This will work even when window doesn't have focus
        keyboard.add_hotkey('esc', self.emergency_stop, suppress=True)
        
        # Track if we're currently typing
        self.is_typing = False
        
        # VM window patterns for focus detection
        self.vm_window_patterns = []
        self.check_focus_enabled = tk.BooleanVar(value=False)
        
        self._create_widgets()
        self._create_layout()
        self._create_bindings()
    
    def _create_widgets(self):
        """Create all the widgets for the application."""
        # Add emergency stop indicator
        self.emergency_frame = ttk.Frame(self.root)
        self.emergency_label = ttk.Label(
            self.emergency_frame, 
            text="PRESS ESC FOR EMERGENCY STOP", 
            foreground="red",
            font=("Arial", 12, "bold")
        )
        
        # Text input area
        self.text_frame = ttk.LabelFrame(self.root, text="Text to Type")
        self.text_input = scrolledtext.ScrolledText(self.text_frame, wrap=tk.WORD, width=60, height=15)
        
        # Control frame
        self.control_frame = ttk.LabelFrame(self.root, text="Controls")
        
        # VM Focus Detection controls
        self.focus_frame = ttk.LabelFrame(self.control_frame, text="VM Focus Detection")
        self.focus_check = ttk.Checkbutton(
            self.focus_frame, 
            text="Enable VM window focus detection",
            variable=self.check_focus_enabled
        )
        self.vm_title_label = ttk.Label(self.focus_frame, text="VM Window Title Contains:")
        self.vm_title_entry = ttk.Entry(self.focus_frame, width=30)
        self.vm_title_button = ttk.Button(
            self.focus_frame, 
            text="Add",
            command=self.add_vm_window_pattern
        )
        self.vm_patterns_display = ttk.Label(
            self.focus_frame, 
            text="No VM window patterns added"
        )
        
        # Speed control
        self.speed_frame = ttk.Frame(self.control_frame)
        self.speed_label = ttk.Label(self.speed_frame, text="Typing Speed:")
        self.speed_slider = ttk.Scale(self.speed_frame, from_=1, to=10, orient=tk.HORIZONTAL, length=200)
        self.speed_slider.set(5)  # Default mid-range speed
        self.speed_value_label = ttk.Label(self.speed_frame, text="5 (Medium)")
        
        # Buttons
        self.button_frame = ttk.Frame(self.control_frame)
        self.load_button = ttk.Button(self.button_frame, text="Load Text", command=self.load_text)
        self.start_button = ttk.Button(self.button_frame, text="Start Typing (F9)", command=self.start_typing)
        self.pause_button = ttk.Button(self.button_frame, text="Pause/Resume (F10)", command=self.toggle_pause)
        self.stop_button = ttk.Button(self.button_frame, text="Stop (ESC)", command=self.stop_typing_process)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        
        # Progress indicator
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        
        # Register hotkeys for start and pause
        keyboard.add_hotkey('f9', self.start_typing)
        keyboard.add_hotkey('f10', self.toggle_pause)
    
    def _create_layout(self):
        """Arrange the widgets in the layout."""
        # Emergency stop indicator at the top
        self.emergency_frame.pack(fill=tk.X, padx=10, pady=5)
        self.emergency_label.pack(side=tk.TOP, pady=5)
        self.emergency_label.config(background=self.root.cget('background'))
        
        # Text frame
        self.text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.text_input.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Control frame
        self.control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # VM Focus Detection layout
        self.focus_frame.pack(fill=tk.X, padx=5, pady=5)
        self.focus_check.pack(anchor=tk.W, padx=5, pady=2)
        
        focus_entry_frame = ttk.Frame(self.focus_frame)
        focus_entry_frame.pack(fill=tk.X, padx=5, pady=2)
        self.vm_title_label.pack(side=tk.LEFT, padx=2)
        self.vm_title_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        self.vm_title_button.pack(side=tk.LEFT, padx=2)
        
        self.vm_patterns_display.pack(anchor=tk.W, padx=5, pady=2)
        
        # Speed control
        self.speed_frame.pack(fill=tk.X, padx=5, pady=5)
        self.speed_label.pack(side=tk.LEFT, padx=5)
        self.speed_slider.pack(side=tk.LEFT, padx=5)
        self.speed_value_label.pack(side=tk.LEFT, padx=5)
        
        # Buttons
        self.button_frame.pack(fill=tk.X, padx=5, pady=5)
        self.load_button.pack(side=tk.LEFT, padx=5)
        self.start_button.pack(side=tk.LEFT, padx=5)
        self.pause_button.pack(side=tk.LEFT, padx=5)
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # Status and progress
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.progress_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
    
    def _create_bindings(self):
        """Set up event bindings."""
        self.speed_slider.bind("<Motion>", self.update_speed_label)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
        # Focus event bindings for transparency change
        self.root.bind("<FocusOut>", self.on_focus_out)
        self.root.bind("<FocusIn>", self.on_focus_in)
    
    def add_vm_window_pattern(self):
        """Add a VM window pattern to the list."""
        pattern = self.vm_title_entry.get().strip()
        if pattern:
            self.vm_window_patterns.append(pattern)
            self.vm_title_entry.delete(0, tk.END)
            self.update_vm_patterns_display()
    
    def update_vm_patterns_display(self):
        """Update the display of VM window patterns."""
        if not self.vm_window_patterns:
            self.vm_patterns_display.config(text="No VM window patterns added")
        else:
            patterns_text = "VM windows: " + ", ".join(
                f'"{p}"' for p in self.vm_window_patterns[:3]
            )
            if len(self.vm_window_patterns) > 3:
                patterns_text += f" and {len(self.vm_window_patterns)-3} more"
            self.vm_patterns_display.config(text=patterns_text)
    
    def on_focus_out(self, event):
        """Handle focus out event - make window more transparent."""
        if self.is_typing:
            self.root.attributes('-alpha', 0.3)  # More transparent when typing and lost focus
        else:
            self.root.attributes('-alpha', 0.7)  # Slightly transparent when not typing and lost focus
    
    def on_focus_in(self, event):
        """Handle focus in event - restore opacity."""
        if self.is_typing:
            self.root.attributes('-alpha', 0.7)  # Partially transparent when typing
        else:
            self.root.attributes('-alpha', 1.0)  # Fully opaque when not typing
    
    def update_speed_label(self, event=None):
        """Update the speed label when the slider moves."""
        speed = int(self.speed_slider.get())
        speed_text = {
            1: "1 (Very Slow)",
            2: "2 (Slow)",
            3: "3 (Moderate Slow)",
            4: "4 (Below Medium)",
            5: "5 (Medium)",
            6: "6 (Above Medium)",
            7: "7 (Moderate Fast)",
            8: "8 (Fast)",
            9: "9 (Very Fast)",
            10: "10 (Ultra Fast)"
        }
        self.speed_value_label.config(text=speed_text.get(speed, f"{speed}"))
    
    def load_text(self):
        """Load text from a file."""
        file_path = filedialog.askopenfilename(
            title="Select Text File",
            filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as file:
                    self.text_input.delete(1.0, tk.END)
                    self.text_input.insert(tk.END, file.read())
                self.status_var.set(f"Loaded: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not load file: {str(e)}")
    
    def start_typing(self):
        """Start the typing process in a separate thread."""
        if self.typing_thread and self.typing_thread.is_alive():
            # Already typing
            return
        
        text = self.text_input.get(1.0, tk.END)
        if not text.strip():
            messagebox.showinfo("Info", "Please enter some text to type.")
            return
        
        # Validate text for unsupported characters
        is_valid, unsupported_chars, positions = validate_text(text)
        if not is_valid:
            # Show the unsupported characters with options to continue or cancel
            char_list = ', '.join([f"'{c}'" for c in unsupported_chars])
            positions_sample = ', '.join([f"pos {p+1}: '{c}'" for p, c in positions[:5]])
            if len(positions) > 5:
                positions_sample += f", and {len(positions)-5} more..."
            
            message = (
                f"Your text contains {len(positions)} characters that may not be supported:\n\n"
                f"Characters: {char_list}\n"
                f"Examples: {positions_sample}\n\n"
                f"These characters might not type correctly. Do you want to continue anyway?"
            )
            
            response = messagebox.askyesno("Unsupported Characters", message)
            if not response:
                return
        
        # Check VM focus if enabled before starting
        if self.check_focus_enabled.get() and self.vm_window_patterns:
            if not check_vm_focus(self.vm_window_patterns):
                messagebox.showwarning(
                    "VM Focus Warning",
                    "The VM window does not appear to be in focus. Make sure to focus the VM window before typing."
                )
                return
        
        # Tell user to focus the VM window
        messagebox.showinfo(
            "Prepare to Type",
            "Click OK, then focus your cursor in the VM window where typing should begin.\n"
            "Typing will start after 3 seconds."
        )
        
        # Reset flags
        self.stop_typing.clear()
        self.paused.clear()
        
        # Get typing speed
        speed_factor = self.speed_slider.get()
        
        # Create and start the typing thread
        self.typing_thread = threading.Thread(
            target=self.typing_process,
            args=(text, speed_factor)
        )
        self.typing_thread.daemon = True
        self.typing_thread.start()
        
        # Set typing flag and make window semi-transparent
        self.is_typing = True
        self.root.attributes('-alpha', 0.7)
        
        # Make emergency label blink by changing its background color
        self.start_emergency_indicator()
        
        self.status_var.set("Typing in progress...")
        self.start_button.config(state=tk.DISABLED)
        self.load_button.config(state=tk.DISABLED)
    
    def start_emergency_indicator(self):
        """Start blinking the emergency indicator."""
        if not self.is_typing:
            return
            
        # Toggle between red and normal background
        current_bg = self.emergency_label.cget('background')
        new_bg = "red" if current_bg != "red" else self.root.cget('background')
        self.emergency_label.config(background=new_bg)
        
        # Schedule next toggle if still typing
        if self.is_typing:
            self.root.after(500, self.start_emergency_indicator)
    
    def typing_process(self, text, speed_factor):
        """Run the typing process in a background thread."""
        total_chars = len(text)
        chars_typed = 0
        failed_chars = 0
        
        try:
            # Give the user time to focus the VM window
            for i in range(3, 0, -1):
                if self.stop_typing.is_set():
                    return
                self.status_var.set(f"Starting in {i}...")
                time.sleep(1)
            
            # Start typing using the keystroke engine
            for progress, char, success in self.keystroke_engine.type_text(text, speed_factor, self.stop_typing, self.paused):
                # Check VM focus if enabled (every 10 characters)
                if self.check_focus_enabled.get() and self.vm_window_patterns and chars_typed % 10 == 0:
                    if not check_vm_focus(self.vm_window_patterns):
                        # Pause typing if focus is lost
                        if not self.paused.is_set():
                            self.paused.set()
                            self.status_var.set("Typing paused: VM window lost focus")
                            
                            # Create a popup notification
                            self.root.after(0, lambda: messagebox.showwarning(
                                "Focus Lost",
                                "VM window has lost focus. Typing is paused.\n\n"
                                "Focus the VM window to continue or press ESC to stop."
                            ))
                            
                            # Wait for focus to return or user to cancel
                            while not check_vm_focus(self.vm_window_patterns) and not self.stop_typing.is_set():
                                time.sleep(0.5)
                                
                            # Resume if focus returned and not stopped
                            if not self.stop_typing.is_set():
                                self.paused.clear()
                                self.status_var.set("Typing resumed: VM window focus restored")
                
                chars_typed += 1
                if not success:
                    failed_chars += 1
                    
                self.progress_var.set(progress * 100)
                self.status_var.set(f"Typing in progress... ({failed_chars} failed keys)")
                self.root.update_idletasks()
                
        except Exception as e:
            messagebox.showerror("Error", f"Typing error: {str(e)}")
        finally:
            # Reset UI
            self.is_typing = False
            self.root.attributes('-alpha', 1.0)  # Restore full opacity
            self.emergency_label.config(background=self.root.cget('background'))
            
            # Show final status with failed keys summary
            if failed_chars > 0:
                self.status_var.set(f"Typing completed with {failed_chars} failed keys")
                # Show a summary of failed keys if any
                if self.keystroke_engine.failed_keys:
                    failed_msg = "Some keys failed to type. The following characters failed:\n\n"
                    for pos, char in self.keystroke_engine.failed_keys[:10]:  # Show first 10 failed keys
                        failed_msg += f"Position {pos+1}: '{char}'\n"
                    if len(self.keystroke_engine.failed_keys) > 10:
                        failed_msg += f"...and {len(self.keystroke_engine.failed_keys) - 10} more."
                    messagebox.showwarning("Failed Keys", failed_msg)
            else:
                self.status_var.set("Typing completed successfully")
                
            self.start_button.config(state=tk.NORMAL)
            self.load_button.config(state=tk.NORMAL)
    
    def toggle_pause(self):
        """Pause or resume the typing process."""
        if not self.typing_thread or not self.typing_thread.is_alive():
            return
        
        if self.paused.is_set():
            # Check VM focus before resuming if enabled
            if self.check_focus_enabled.get() and self.vm_window_patterns:
                if not check_vm_focus(self.vm_window_patterns):
                    messagebox.showwarning(
                        "VM Focus Warning",
                        "The VM window does not appear to be in focus. Make sure to focus the VM window before resuming."
                    )
                    return
            
            self.paused.clear()
            self.status_var.set("Typing resumed")
        else:
            self.paused.set()
            self.status_var.set("Typing paused")
    
    def stop_typing_process(self):
        """Stop the typing process."""
        self.emergency_stop()
    
    def emergency_stop(self):
        """Emergency stop for the typing process."""
        if self.typing_thread and self.typing_thread.is_alive():
            self.stop_typing.set()
            self.paused.clear()  # Clear pause if paused
            self.is_typing = False
            self.root.attributes('-alpha', 1.0)  # Restore full opacity
            self.emergency_label.config(background=self.root.cget('background'))
            self.status_var.set("Typing STOPPED by emergency stop")
            self.start_button.config(state=tk.NORMAL)
            self.load_button.config(state=tk.NORMAL)
            
            # Make the window flash to indicate emergency stop was triggered
            for i in range(5):
                self.root.attributes('-alpha', 0.3)
                self.root.update()
                time.sleep(0.1)
                self.root.attributes('-alpha', 1.0)
                self.root.update()
                time.sleep(0.1)
    
    def on_close(self):
        """Handle window close event."""
        if self.typing_thread and self.typing_thread.is_alive():
            self.emergency_stop()
            self.typing_thread.join(1.0)  # Wait for thread to finish
        
        # Remove keyboard hooks
        keyboard.unhook_all()
        
        # Clear any sensitive data
        self.text_input.delete(1.0, tk.END)
        clear_clipboard()
        
        # Close the window
        self.root.destroy()
    
    def run(self):
        """Run the main application loop."""
        self.root.mainloop()

# -----------------------------------------------------------------------------
# Main Entry Point
# -----------------------------------------------------------------------------

def main():
    """Main entry point for the application."""
    logger.info("Starting VM Keystroke Simulator")
    
    # Initialize and run the GUI
    app = KeystrokeSimulatorGUI()
    app.run()
    
    logger.info("VM Keystroke Simulator closed")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Unexpected error: {str(e)}")
        sys.exit(1)