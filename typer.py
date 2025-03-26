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

# -----------------------------------------------------------------------------
# VM Integration Functions
# -----------------------------------------------------------------------------

def check_vm_focus():
    """
    Check if the VM window is in focus.
    
    Returns:
        bool: True if the VM appears to be in focus
    """
    # This is a basic implementation. In a real application, you would add
    # more sophisticated focus detection for your specific VM software.
    # For now, we'll just return True and rely on the user to ensure focus.
    return True

def send_key(key):
    """
    Send a single key press to the VM.
    
    Args:
        key (str): The key to press
    """
    try:
        pyautogui.press(key)
        logger.debug(f"Sent key: {key}")
    except Exception as e:
        logger.error(f"Error sending key '{key}': {str(e)}")
        raise

def send_key_combination(modifier, key):
    """
    Send a key combination (e.g., Shift+1 for !) to the VM.
    
    Args:
        modifier (str): The modifier key (e.g., 'shift', 'ctrl')
        key (str): The key to press with the modifier
    """
    try:
        pyautogui.hotkey(modifier, key)
        logger.debug(f"Sent key combination: {modifier}+{key}")
    except Exception as e:
        logger.error(f"Error sending key combination '{modifier}+{key}': {str(e)}")
        raise

def send_text(text):
    """
    Send a text string to the VM all at once.
    Note: This is faster but less human-like than character-by-character.
    
    Args:
        text (str): The text to type
    """
    try:
        pyautogui.write(text)
        logger.debug(f"Sent text: '{text[:10]}...' ({len(text)} chars)")
    except Exception as e:
        logger.error(f"Error sending text: {str(e)}")
        raise

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
    
    def type_text(self, text, speed_factor, stop_event, pause_event):
        """
        Type the given text at the specified speed with natural timing.
        
        Args:
            text (str): The text to type
            speed_factor (float): Speed factor from 1 (slow) to 10 (fast)
            stop_event (threading.Event): Event to signal stopping
            pause_event (threading.Event): Event to signal pausing
            
        Yields:
            tuple: (progress_fraction, current_character)
        """
        total_chars = len(text)
        chars_typed = 0
        
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
            try:
                if char == '\n':
                    send_key('enter')
                elif char == '\t':
                    send_key('tab')
                elif char in self.special_chars:
                    modifier, key = self.special_chars[char]
                    send_key_combination(modifier, key)
                else:
                    send_key(char)
                
                chars_typed += 1
                progress = chars_typed / total_chars
                
                # Wait before typing the next character
                time.sleep(actual_delay)
                
                # Yield progress and current character
                yield progress, char
            
            except Exception as e:
                logger.error(f"Error typing character '{char}': {str(e)}")
                raise
        
        logger.info("Typing completed")

# -----------------------------------------------------------------------------
# GUI Implementation
# -----------------------------------------------------------------------------

class KeystrokeSimulatorGUI:
    """Main GUI class for the Keystroke Simulator."""
    
    def __init__(self):
        """Initialize the GUI components."""
        self.root = tk.Tk()
        self.root.title("VM Keystroke Simulator")
        self.root.geometry("600x500")
        self.root.minsize(500, 400)
        
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
        
        try:
            # Give the user time to focus the VM window
            for i in range(3, 0, -1):
                if self.stop_typing.is_set():
                    return
                self.status_var.set(f"Starting in {i}...")
                time.sleep(1)
            
            # Start typing using the keystroke engine
            for progress, char in self.keystroke_engine.type_text(text, speed_factor, self.stop_typing, self.paused):
                chars_typed += 1
                self.progress_var.set(progress * 100)
                self.root.update_idletasks()
        except Exception as e:
            messagebox.showerror("Error", f"Typing error: {str(e)}")
        finally:
            # Reset UI
            self.is_typing = False
            self.root.attributes('-alpha', 1.0)  # Restore full opacity
            self.emergency_label.config(background=self.root.cget('background'))
            self.status_var.set("Typing completed")
            self.start_button.config(state=tk.NORMAL)
            self.load_button.config(state=tk.NORMAL)
    
    def toggle_pause(self):
        """Pause or resume the typing process."""
        if not self.typing_thread or not self.typing_thread.is_alive():
            return
        
        if self.paused.is_set():
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