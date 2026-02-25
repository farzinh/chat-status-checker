"""
Chat Status Monitor v2 - Region Selection + System Tray
========================================================
A more reliable version that:
1. Lets you select a specific region (the chat list area)
2. Runs in system tray (no UI interference)
3. Only monitors the selected region

Requirements:
    pip install pyautogui opencv-python numpy pytesseract mss Pillow pystray tzdata

Also install Tesseract OCR:
    Download from: https://github.com/UB-Mannheim/tesseract/wiki
"""

import cv2
import numpy as np
import pytesseract
from mss import mss
from PIL import Image, ImageDraw, ImageTk
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import time
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from typing import Optional, Tuple
from datetime import datetime, timedelta
import sys

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

# Try to import pystray for system tray
try:
    import pystray
    from pystray import MenuItem as item
    TRAY_AVAILABLE = True
except ImportError:
    TRAY_AVAILABLE = False
    print("Note: pystray not installed. System tray feature disabled.")
    print("Install with: pip install pystray")


class RegionSelector:
    """Fullscreen overlay to select a region - supports multiple monitors"""
    
    def __init__(self, callback):
        self.callback = callback
        self.start_x = None
        self.start_y = None
        self.current_rect = None
        self.is_dragging = False
        
    def select(self):
        """Show overlay on ALL monitors and let user draw a rectangle"""
        # Get all monitors info
        with mss() as sct:
            # monitors[0] is the "all monitors" combined, monitors[1], [2], etc are individual
            all_monitors = sct.monitors[0]  # Combined bounding box of all monitors
            self.offset_x = all_monitors['left']  # Can be negative if monitor is to the left
            self.offset_y = all_monitors['top']
            self.total_width = all_monitors['width']
            self.total_height = all_monitors['height']
            
            # Take screenshot of ALL monitors
            screenshot = sct.grab(all_monitors)
            self.screenshot = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
        
        # Create window that spans ALL monitors
        self.root = tk.Toplevel()
        self.root.title("Select Region")
        self.root.attributes('-topmost', True)
        self.root.config(cursor="cross")
        
        # Position window to cover all monitors (including negative coordinates)
        self.root.geometry(f"{self.total_width}x{self.total_height}+{self.offset_x}+{self.offset_y}")
        self.root.overrideredirect(True)  # Remove window decorations
        
        # Create canvas
        self.canvas = tk.Canvas(
            self.root, 
            width=self.total_width,
            height=self.total_height,
            highlightthickness=0,
            cursor="cross"
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Show screenshot as background
        self.bg_image = ImageTk.PhotoImage(self.screenshot)
        self.canvas.create_image(0, 0, image=self.bg_image, anchor='nw')
        
        # Add semi-transparent dark overlay
        self.canvas.create_rectangle(
            0, 0, self.total_width, self.total_height,
            fill='black', stipple='gray25', tags='overlay'
        )
        
        # Instructions - show on each monitor
        with mss() as sct:
            for i, mon in enumerate(sct.monitors[1:], 1):  # Skip the "all" monitor
                # Calculate position relative to combined screenshot
                text_x = mon['left'] - self.offset_x + mon['width'] // 2
                text_y = mon['top'] - self.offset_y + 50
                
                # Background for text
                self.canvas.create_rectangle(
                    text_x - 350, text_y - 30,
                    text_x + 350, text_y + 30,
                    fill='black', outline='white'
                )
                self.canvas.create_text(
                    text_x, text_y,
                    text=f"Monitor {i}: 🖱️ CLICK and DRAG to select chat list area  |  ESC to cancel",
                    fill='white', font=('Arial', 14, 'bold')
                )
        
        # Bind mouse events
        self.canvas.bind('<ButtonPress-1>', self.on_mouse_down)
        self.canvas.bind('<B1-Motion>', self.on_mouse_move)
        self.canvas.bind('<ButtonRelease-1>', self.on_mouse_up)
        self.root.bind('<Escape>', self.on_cancel)
        
        self.root.focus_force()
        self.root.grab_set()
        
    def on_mouse_down(self, event):
        """Mouse button pressed - start drawing"""
        self.start_x = event.x
        self.start_y = event.y
        self.is_dragging = True
        
        # Remove old rectangle if exists
        if self.current_rect:
            self.canvas.delete(self.current_rect)
            self.canvas.delete('size_text')
        
        # Create new rectangle
        self.current_rect = self.canvas.create_rectangle(
            self.start_x, self.start_y, self.start_x, self.start_y,
            outline='#00FF00', width=3, dash=(5, 5)
        )
        
    def on_mouse_move(self, event):
        """Mouse dragged - update rectangle size"""
        if self.is_dragging and self.current_rect:
            # Update rectangle
            self.canvas.coords(self.current_rect, self.start_x, self.start_y, event.x, event.y)
            
            # Show size
            width = abs(event.x - self.start_x)
            height = abs(event.y - self.start_y)
            
            self.canvas.delete('size_text')
            self.canvas.create_text(
                (self.start_x + event.x) // 2,
                (self.start_y + event.y) // 2,
                text=f"{width} x {height}",
                fill='#00FF00', font=('Arial', 14, 'bold'),
                tags='size_text'
            )
            
    def on_mouse_up(self, event):
        """Mouse released - finish selection"""
        if not self.is_dragging:
            return
            
        self.is_dragging = False
        
        # Calculate rectangle bounds (in canvas coordinates)
        canvas_x1 = min(self.start_x, event.x)
        canvas_y1 = min(self.start_y, event.y)
        canvas_x2 = max(self.start_x, event.x)
        canvas_y2 = max(self.start_y, event.y)
        
        width = canvas_x2 - canvas_x1
        height = canvas_y2 - canvas_y1
        
        # Minimum size check
        if width < 50 or height < 50:
            self.canvas.delete('size_text')
            self.canvas.create_text(
                (canvas_x1 + canvas_x2) // 2, (canvas_y1 + canvas_y2) // 2,
                text=f"Too small! ({width}x{height})\nDrag a larger area",
                fill='red', font=('Arial', 14, 'bold'),
                tags='size_text'
            )
            if self.current_rect:
                self.canvas.itemconfig(self.current_rect, outline='red')
            return
        
        # Convert canvas coordinates to SCREEN coordinates
        # Canvas (0,0) = screen (offset_x, offset_y)
        screen_x1 = canvas_x1 + self.offset_x
        screen_y1 = canvas_y1 + self.offset_y
        screen_x2 = canvas_x2 + self.offset_x
        screen_y2 = canvas_y2 + self.offset_y
        
        # Valid selection
        self.root.destroy()
        self.callback((screen_x1, screen_y1, screen_x2, screen_y2))
        
    def on_cancel(self, event):
        """ESC pressed - cancel selection"""
        self.root.destroy()
        self.callback(None)


class StatusCalibrator:
    """Tool to click on status icon and sample its position relative to region"""
    
    def __init__(self, region, callback):
        self.region = region  # (x1, y1, x2, y2)
        self.callback = callback
        
    def calibrate(self):
        """Show overlay on region and let user click on status icon"""
        x1, y1, x2, y2 = self.region
        width = x2 - x1
        height = y2 - y1
        
        self.root = tk.Toplevel()
        self.root.overrideredirect(True)
        self.root.attributes('-topmost', True)
        self.root.attributes('-alpha', 0.6)
        self.root.geometry(f"{width}x{height}+{x1}+{y1}")
        self.root.configure(bg='blue')
        self.root.config(cursor="crosshair")
        
        # Canvas
        canvas = tk.Canvas(self.root, bg='blue', highlightthickness=0, cursor="crosshair")
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # Instructions box at top
        canvas.create_rectangle(0, 0, width, 60, fill='black', outline='white')
        canvas.create_text(
            width // 2, 30,
            text="👆 Click on the STATUS DOT (colored circle)\nPress ESC to cancel",
            fill='white', font=('Arial', 12, 'bold'),
            justify='center'
        )
        
        # Bind events
        canvas.bind('<Button-1>', self.on_click)
        self.root.bind('<Escape>', self.on_cancel)
        
        self.root.focus_force()
        self.root.grab_set()
        
    def on_click(self, event):
        """User clicked - record position relative to region"""
        # Position relative to the region (not screen)
        rel_x = event.x
        rel_y = event.y
        
        # Ignore clicks in the instruction area
        if rel_y < 65:
            return
        
        self.root.destroy()
        self.callback((rel_x, rel_y))
        
    def on_cancel(self, event):
        self.root.destroy()
        self.callback(None)


class ChatStatusMonitorV2:
    """Main application with Region Selection and System Tray"""
    
    CONFIG_FILE = "monitor_config_v2.json"
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Chat Status Monitor v2")
        self.root.geometry("550x700")
        self.root.resizable(True, True)
        self.root.minsize(450, 500)
        
        # State
        self.running = False
        self.monitor_thread = None
        self.tray_icon = None
        self.region = None  # (x1, y1, x2, y2)
        self.status_position = None  # (rel_x, rel_y) relative to name
        self.last_status = None
        self.last_email_time = None
        self.last_notified_status = None
        
        # Setup
        self.setup_scrollable_frame()
        self.setup_ui()
        self.load_config()
        
        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        
    def setup_scrollable_frame(self):
        """Create scrollable container"""
        self.canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        
        self.scrollable_frame = ttk.Frame(self.canvas)
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.itemconfig(self.canvas_frame, width=e.width))
        
        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # Mousewheel scrolling
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
    def setup_ui(self):
        """Setup the user interface"""
        main = ttk.Frame(self.scrollable_frame, padding="10")
        main.pack(fill=tk.BOTH, expand=True)
        
        # Title
        ttk.Label(main, text="🔍 Chat Status Monitor v2", font=("Arial", 16, "bold")).pack(pady=(0, 10))
        ttk.Label(main, text="Region-based monitoring with System Tray", font=("Arial", 10)).pack(pady=(0, 15))
        
        # === REGION SELECTION ===
        region_frame = ttk.LabelFrame(main, text="1. Select Chat List Region", padding="10")
        region_frame.pack(fill=tk.X, pady=5)
        
        self.region_label = ttk.Label(region_frame, text="No region selected", foreground="red")
        self.region_label.pack(anchor=tk.W)
        
        btn_frame = ttk.Frame(region_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="📐 Select Region", command=self.select_region).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="👁️ Preview Region", command=self.preview_region).pack(side=tk.LEFT, padx=2)
        
        # === TARGET PERSON ===
        person_frame = ttk.LabelFrame(main, text="2. Target Person", padding="10")
        person_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(person_frame, text="Name to monitor:").pack(anchor=tk.W)
        self.person_entry = ttk.Entry(person_frame, width=40)
        self.person_entry.pack(fill=tk.X, pady=2)
        self.person_entry.insert(0, "Arne Kaulfuß")
        
        # === CALIBRATION ===
        calib_frame = ttk.LabelFrame(main, text="3. Calibrate Status Position", padding="10")
        calib_frame.pack(fill=tk.X, pady=5)
        
        self.calib_label = ttk.Label(calib_frame, text="Not calibrated", foreground="red")
        self.calib_label.pack(anchor=tk.W)
        
        ttk.Button(calib_frame, text="🎯 Calibrate Status Position", command=self.calibrate_status).pack(anchor=tk.W, pady=5)
        
        ttk.Label(calib_frame, text="Click on the status dot (colored circle) next to the person's name", 
                  font=("Arial", 9), foreground="gray").pack(anchor=tk.W)
        
        # === TESSERACT ===
        tess_frame = ttk.LabelFrame(main, text="4. Tesseract OCR", padding="10")
        tess_frame.pack(fill=tk.X, pady=5)
        
        path_frame = ttk.Frame(tess_frame)
        path_frame.pack(fill=tk.X)
        
        ttk.Label(path_frame, text="Path:").pack(side=tk.LEFT)
        self.tesseract_path = ttk.Entry(path_frame, width=35)
        self.tesseract_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.tesseract_path.insert(0, r"C:\Program Files\Tesseract-OCR\tesseract.exe")
        ttk.Button(path_frame, text="Browse", command=self.browse_tesseract).pack(side=tk.LEFT)
        
        # === EMAIL SETTINGS ===
        email_frame = ttk.LabelFrame(main, text="5. Email Notifications (Optional)", padding="10")
        email_frame.pack(fill=tk.X, pady=5)
        
        self.email_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(email_frame, text="Enable email notifications", variable=self.email_enabled).pack(anchor=tk.W)
        
        # SMTP settings in a sub-frame
        smtp_frame = ttk.Frame(email_frame)
        smtp_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(smtp_frame, text="SMTP:").grid(row=0, column=0, sticky=tk.W)
        self.smtp_server = ttk.Entry(smtp_frame, width=25)
        self.smtp_server.grid(row=0, column=1, padx=2)
        self.smtp_server.insert(0, "smtp.gmail.com")
        
        ttk.Label(smtp_frame, text="Port:").grid(row=0, column=2, padx=(10,0))
        self.smtp_port = ttk.Entry(smtp_frame, width=6)
        self.smtp_port.grid(row=0, column=3, padx=2)
        self.smtp_port.insert(0, "587")
        
        ttk.Label(smtp_frame, text="From:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.sender_email = ttk.Entry(smtp_frame, width=35)
        self.sender_email.grid(row=1, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        ttk.Label(smtp_frame, text="Password:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.sender_password = ttk.Entry(smtp_frame, width=35, show="*")
        self.sender_password.grid(row=2, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        ttk.Label(smtp_frame, text="To:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.recipient_email = ttk.Entry(smtp_frame, width=35)
        self.recipient_email.grid(row=3, column=1, columnspan=3, sticky=tk.W, pady=2)
        
        # Email schedule
        schedule_frame = ttk.Frame(email_frame)
        schedule_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(schedule_frame, text="Only after:").pack(side=tk.LEFT)
        self.email_start_hour = tk.StringVar(value="9")
        ttk.Spinbox(schedule_frame, from_=0, to=23, width=3, textvariable=self.email_start_hour).pack(side=tk.LEFT, padx=2)
        ttk.Label(schedule_frame, text=":00 Berlin  |  Min interval:").pack(side=tk.LEFT)
        self.email_rate_limit = tk.StringVar(value="60")
        ttk.Spinbox(schedule_frame, from_=1, to=1440, width=5, textvariable=self.email_rate_limit).pack(side=tk.LEFT, padx=2)
        ttk.Label(schedule_frame, text="min").pack(side=tk.LEFT)
        
        # Notify options
        self.notify_green = tk.BooleanVar(value=True)
        self.notify_red = tk.BooleanVar(value=True)
        ttk.Checkbutton(email_frame, text="Notify on GREEN", variable=self.notify_green).pack(anchor=tk.W)
        ttk.Checkbutton(email_frame, text="Notify on RED", variable=self.notify_red).pack(anchor=tk.W)
        
        # === SETTINGS ===
        settings_frame = ttk.LabelFrame(main, text="6. Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=5)
        
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.pack(fill=tk.X)
        ttk.Label(interval_frame, text="Check every:").pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="3")
        ttk.Spinbox(interval_frame, from_=1, to=60, width=5, textvariable=self.interval_var).pack(side=tk.LEFT, padx=5)
        ttk.Label(interval_frame, text="seconds").pack(side=tk.LEFT)
        
        # === STATUS ===
        status_frame = ttk.LabelFrame(main, text="Status", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(status_frame, text="⏸️ Not running", font=("Arial", 12))
        self.status_label.pack()
        
        self.detection_label = ttk.Label(status_frame, text="", font=("Arial", 10))
        self.detection_label.pack()
        
        # === CONTROL BUTTONS ===
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.start_btn = ttk.Button(btn_frame, text="▶️ Start Monitoring", command=self.toggle_monitoring)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame, text="🧪 Test", command=self.test_detection).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="💾 Save", command=self.save_config).pack(side=tk.LEFT, padx=2)
        
        if TRAY_AVAILABLE:
            ttk.Button(btn_frame, text="📥 Minimize to Tray", command=self.minimize_to_tray).pack(side=tk.LEFT, padx=2)
        
    def select_region(self):
        """Let user select the chat list region on any monitor"""
        # Minimize main window to avoid interference
        self.root.iconify()
        time.sleep(0.3)
        
        def on_region_selected(region):
            self.root.deiconify()
            if region:
                self.region = region
                x1, y1, x2, y2 = region
                self.region_label.config(
                    text=f"✓ Region: ({x1}, {y1}) to ({x2}, {y2}) - {x2-x1}x{y2-y1} px",
                    foreground="green"
                )
                self.save_config_silent()
                
                # Show quick preview
                messagebox.showinfo("Region Selected", 
                    f"Region selected successfully!\n\n"
                    f"Position: ({x1}, {y1}) to ({x2}, {y2})\n"
                    f"Size: {x2-x1} x {y2-y1} pixels\n\n"
                    f"Click 'Preview Region' to see what will be captured."
                )
            else:
                self.region_label.config(text="Selection cancelled", foreground="orange")
        
        selector = RegionSelector(on_region_selected)
        selector.select()
        
    def preview_region(self):
        """Show the selected region with a border"""
        if not self.region:
            messagebox.showwarning("No Region", "Please select a region first")
            return
        
        x1, y1, x2, y2 = self.region
        
        # Capture just that region (works across monitors)
        with mss() as sct:
            monitor = {
                "left": x1,
                "top": y1, 
                "width": x2 - x1,
                "height": y2 - y1
            }
            try:
                screenshot = sct.grab(monitor)
                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
            except Exception as e:
                messagebox.showerror("Error", f"Could not capture region: {e}")
                return
        
        # Show in a new window
        preview = tk.Toplevel(self.root)
        preview.title("Region Preview")
        preview.attributes('-topmost', True)
        
        photo = ImageTk.PhotoImage(img)
        label = ttk.Label(preview, image=photo)
        label.image = photo  # Keep reference
        label.pack()
        
        ttk.Label(preview, text=f"This is what will be monitored\nRegion: ({x1}, {y1}) to ({x2}, {y2})").pack(pady=5)
        ttk.Button(preview, text="Close", command=preview.destroy).pack(pady=5)
        
    def calibrate_status(self):
        """Calibrate the status icon position"""
        if not self.region:
            messagebox.showwarning("No Region", "Please select a region first")
            return
        
        # Minimize main window
        self.root.iconify()
        time.sleep(0.2)
        
        def on_calibrated(pos):
            self.root.deiconify()
            if pos:
                self.status_position = pos
                self.calib_label.config(
                    text=f"✓ Status position: ({pos[0]}, {pos[1]}) relative to region",
                    foreground="green"
                )
                self.save_config_silent()
                messagebox.showinfo("Calibrated", 
                    f"Status position calibrated!\n\n"
                    f"Position: ({pos[0]}, {pos[1]}) within the region\n\n"
                    f"Click 'Test' to verify detection works."
                )
            else:
                self.calib_label.config(text="Calibration cancelled", foreground="orange")
        
        calibrator = StatusCalibrator(self.region, on_calibrated)
        calibrator.calibrate()
        
    def browse_tesseract(self):
        """Browse for Tesseract executable"""
        path = filedialog.askopenfilename(
            title="Select Tesseract",
            filetypes=[("Executable", "*.exe"), ("All", "*.*")]
        )
        if path:
            self.tesseract_path.delete(0, tk.END)
            self.tesseract_path.insert(0, path)
            
    def capture_region(self) -> Optional[np.ndarray]:
        """Capture just the selected region (works across monitors)"""
        if not self.region:
            return None
        
        x1, y1, x2, y2 = self.region
        
        with mss() as sct:
            # Use absolute coordinates - mss handles multi-monitor
            monitor = {
                "left": x1,
                "top": y1,
                "width": x2 - x1,
                "height": y2 - y1
            }
            try:
                screenshot = sct.grab(monitor)
                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                return cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            except Exception as e:
                print(f"Capture error: {e}")
                return None
        
    def find_name_in_region(self, image: np.ndarray, target: str) -> Optional[Tuple[int, int, int, int]]:
        """Find the target name in the region image"""
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        # Set tesseract path
        tess_path = self.tesseract_path.get()
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path
        
        try:
            data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
        except Exception as e:
            print(f"OCR Error: {e}")
            return None
        
        target_words = target.lower().split()
        first_word = target_words[0]
        
        # Word variants for special characters
        def get_variants(word):
            variants = [word]
            if 'ß' in word:
                variants.extend([word.replace('ß', 'ss'), word.replace('ß', 'b')])
            return variants
        
        n_boxes = len(data['text'])
        
        for i in range(n_boxes):
            text = data['text'][i].strip()
            if not text or len(text) < 2:
                continue
            
            text_lower = text.lower().rstrip(':.,;')
            
            # Check first word match
            if first_word in text_lower or text_lower.startswith(first_word[:3]):
                x, y, w, h = data['left'][i], data['top'][i], data['width'][i], data['height'][i]
                
                # If multi-word name, look for second word nearby
                if len(target_words) >= 2:
                    second_variants = get_variants(target_words[1])
                    
                    for j in range(max(0, i-2), min(n_boxes, i+5)):
                        if j == i:
                            continue
                        nearby = data['text'][j].strip().lower()
                        if abs(data['top'][j] - y) < 20:  # Same line
                            for var in second_variants:
                                if var[:4] in nearby or nearby[:4] in var:
                                    # Combine boxes
                                    x2 = max(x + w, data['left'][j] + data['width'][j])
                                    w = x2 - x
                                    return (x, y, w, h)
                
                return (x, y, w, h)
        
        return None
    
    def detect_status_color(self, image: np.ndarray, name_box: Tuple[int, int, int, int]) -> str:
        """Detect status color using calibrated position"""
        if not self.status_position:
            # Fallback: look left of name
            x, y, w, h = name_box
            search_x = max(0, x - 50)
            search_y = max(0, y)
        else:
            # Use calibrated position (relative to region)
            search_x = max(0, self.status_position[0] - 15)
            search_y = max(0, self.status_position[1] - 15)
        
        search_size = 30
        region = image[search_y:search_y+search_size, search_x:search_x+search_size]
        
        if region.size == 0:
            return "unknown"
        
        hsv = cv2.cvtColor(region, cv2.COLOR_BGR2HSV)
        
        # Color detection
        green_mask = cv2.inRange(hsv, np.array([35, 80, 80]), np.array([85, 255, 255]))
        red_mask1 = cv2.inRange(hsv, np.array([0, 80, 80]), np.array([10, 255, 255]))
        red_mask2 = cv2.inRange(hsv, np.array([160, 80, 80]), np.array([180, 255, 255]))
        yellow_mask = cv2.inRange(hsv, np.array([15, 80, 80]), np.array([35, 255, 255]))
        
        green_px = cv2.countNonZero(green_mask)
        red_px = cv2.countNonZero(cv2.bitwise_or(red_mask1, red_mask2))
        yellow_px = cv2.countNonZero(yellow_mask)
        
        print(f"Colors - G:{green_px} R:{red_px} Y:{yellow_px}")
        
        if green_px > 5 and green_px >= red_px and green_px >= yellow_px:
            return "green"
        elif red_px > 5 and red_px > green_px:
            return "red"
        elif yellow_px > 5:
            return "yellow"
        
        return "unknown"
    
    def test_detection(self):
        """Test the detection with current settings"""
        if not self.region:
            messagebox.showwarning("Setup Required", "Please select a region first")
            return
        
        self.status_label.config(text="🔄 Testing...")
        self.root.update()
        
        # Capture region
        image = self.capture_region()
        if image is None:
            self.status_label.config(text="❌ Capture failed")
            return
        
        # Find name
        target = self.person_entry.get()
        name_box = self.find_name_in_region(image, target)
        
        if name_box:
            x, y, w, h = name_box
            status = self.detect_status_color(image, name_box)
            
            self.status_label.config(text=f"✅ Found: {target}")
            self.detection_label.config(text=f"Position: ({x}, {y}) - Status: {status.upper()}")
            
            # Show visual result
            self.show_test_result(image, name_box, status)
        else:
            self.status_label.config(text=f"❌ Not found: {target}")
            self.detection_label.config(text="Check if name is visible in the selected region")
            
            # Show what was captured
            self.show_captured_region(image)
    
    def show_test_result(self, image: np.ndarray, name_box: Tuple[int, int, int, int], status: str):
        """Show test result in a window"""
        x, y, w, h = name_box
        
        # Draw rectangles on image
        img_copy = image.copy()
        
        # Name box (green)
        cv2.rectangle(img_copy, (x, y), (x+w, y+h), (0, 255, 0), 2)
        
        # Status search area (blue)
        if self.status_position:
            sx, sy = self.status_position
            cv2.rectangle(img_copy, (sx-15, sy-15), (sx+15, sy+15), (255, 0, 0), 2)
        
        # Convert to PIL
        img_rgb = cv2.cvtColor(img_copy, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)
        
        # Show window
        preview = tk.Toplevel(self.root)
        preview.title(f"Test Result - Status: {status.upper()}")
        preview.attributes('-topmost', True)
        
        photo = ImageTk.PhotoImage(pil_img)
        label = ttk.Label(preview, image=photo)
        label.image = photo
        label.pack()
        
        info = f"Name found at: ({x}, {y})\nStatus: {status.upper()}\n\nGreen box = Name\nBlue box = Status search area"
        ttk.Label(preview, text=info, justify=tk.LEFT).pack(pady=5)
        ttk.Button(preview, text="Close", command=preview.destroy).pack(pady=5)
        
    def show_captured_region(self, image: np.ndarray):
        """Show what was captured (for debugging)"""
        img_rgb = cv2.cvtColor(image, cv2.COLOR_BGR2RGB)
        pil_img = Image.fromarray(img_rgb)
        
        preview = tk.Toplevel(self.root)
        preview.title("Captured Region (Name not found)")
        preview.attributes('-topmost', True)
        
        photo = ImageTk.PhotoImage(pil_img)
        label = ttk.Label(preview, image=photo)
        label.image = photo
        label.pack()
        
        ttk.Label(preview, text="This is what was captured.\nMake sure the person's name is visible here.").pack(pady=5)
        ttk.Button(preview, text="Close", command=preview.destroy).pack(pady=5)
    
    def toggle_monitoring(self):
        """Start or stop monitoring"""
        if self.running:
            self.stop_monitoring()
        else:
            self.start_monitoring()
            
    def start_monitoring(self):
        """Start the monitoring loop"""
        if not self.region:
            messagebox.showwarning("Setup Required", "Please select a region first")
            return
        
        self.running = True
        self.start_btn.config(text="⏹️ Stop Monitoring")
        self.status_label.config(text="▶️ Monitoring...")
        
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
        
    def stop_monitoring(self):
        """Stop monitoring"""
        self.running = False
        self.start_btn.config(text="▶️ Start Monitoring")
        self.status_label.config(text="⏸️ Stopped")
        
    def monitor_loop(self):
        """Background monitoring loop"""
        while self.running:
            try:
                image = self.capture_region()
                if image is None:
                    continue
                
                target = self.person_entry.get()
                name_box = self.find_name_in_region(image, target)
                
                if name_box:
                    status = self.detect_status_color(image, name_box)
                    
                    # Update UI
                    self.root.after(0, lambda s=status: self.status_label.config(text=f"✅ {target}: {s.upper()}"))
                    self.root.after(0, lambda s=status: self.detection_label.config(text=f"Last check: {time.strftime('%H:%M:%S')}"))
                    
                    # Check for status change
                    if status != self.last_status and status in ['green', 'red']:
                        should_notify = (
                            (status == 'green' and self.notify_green.get()) or
                            (status == 'red' and self.notify_red.get())
                        )
                        if should_notify and self.email_enabled.get():
                            self.send_notification(target, status)
                        self.last_status = status
                else:
                    self.root.after(0, lambda: self.status_label.config(text=f"🔍 Searching for {target}..."))
                    
            except Exception as e:
                print(f"Monitor error: {e}")
                
            time.sleep(int(self.interval_var.get()))
    
    def can_send_email(self) -> Tuple[bool, str]:
        """Check email constraints"""
        try:
            berlin_tz = ZoneInfo("Europe/Berlin")
            now_berlin = datetime.now(berlin_tz)
        except:
            now_berlin = datetime.now()
        
        if now_berlin.hour < int(self.email_start_hour.get()):
            return False, f"Before {self.email_start_hour.get()}:00"
        
        if self.last_email_time:
            minutes_since = (datetime.now() - self.last_email_time).total_seconds() / 60
            rate_limit = int(self.email_rate_limit.get())
            if minutes_since < rate_limit:
                return False, f"Rate limited ({rate_limit - minutes_since:.0f} min left)"
        
        return True, "OK"
    
    def send_notification(self, name: str, status: str):
        """Send email notification"""
        can_send, reason = self.can_send_email()
        if not can_send:
            print(f"Email blocked: {reason}")
            return
        
        if self.last_notified_status == status:
            print(f"Already notified about {status}")
            return
        
        try:
            msg = MIMEMultipart()
            msg['From'] = self.sender_email.get()
            msg['To'] = self.recipient_email.get()
            msg['Subject'] = f"Status Alert: {name} is now {status.upper()}"
            
            try:
                berlin_time = datetime.now(ZoneInfo("Europe/Berlin")).strftime('%Y-%m-%d %H:%M:%S %Z')
            except:
                berlin_time = time.strftime('%Y-%m-%d %H:%M:%S')
            
            body = f"Person: {name}\nStatus: {status.upper()}\nTime: {berlin_time}"
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get())) as server:
                server.starttls()
                server.login(self.sender_email.get(), self.sender_password.get())
                server.send_message(msg)
            
            self.last_email_time = datetime.now()
            self.last_notified_status = status
            print(f"✉️ Email sent: {name} is {status}")
            
        except Exception as e:
            print(f"Email error: {e}")
    
    # === SYSTEM TRAY ===
    
    def minimize_to_tray(self):
        """Minimize to system tray"""
        if not TRAY_AVAILABLE:
            messagebox.showwarning("Not Available", "Install pystray: pip install pystray")
            return
        
        self.root.withdraw()
        
        # Create tray icon
        icon_image = Image.new('RGB', (64, 64), color='green')
        draw = ImageDraw.Draw(icon_image)
        draw.ellipse([16, 16, 48, 48], fill='lime')
        
        menu = (
            item('Show Window', self.show_from_tray),
            item('Start Monitoring', self.start_monitoring),
            item('Stop Monitoring', self.stop_monitoring),
            item('Exit', self.exit_app)
        )
        
        self.tray_icon = pystray.Icon("ChatMonitor", icon_image, "Chat Status Monitor", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()
        
    def show_from_tray(self):
        """Show window from tray"""
        if self.tray_icon:
            self.tray_icon.stop()
            self.tray_icon = None
        self.root.deiconify()
        self.root.lift()
        
    def exit_app(self):
        """Exit application"""
        self.running = False
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()
        
    def on_close(self):
        """Handle window close"""
        if self.running and TRAY_AVAILABLE:
            if messagebox.askyesno("Minimize?", "Minimize to system tray instead of closing?\n\n(Monitoring will continue in background)"):
                self.minimize_to_tray()
                return
        self.exit_app()
    
    # === CONFIG ===
    
    def save_config(self):
        """Save config with feedback"""
        if self.save_config_silent():
            messagebox.showinfo("Saved", "Configuration saved!")
            
    def save_config_silent(self):
        """Save config without message"""
        config = {
            "region": self.region,
            "status_position": self.status_position,
            "target_person": self.person_entry.get(),
            "tesseract_path": self.tesseract_path.get(),
            "interval": self.interval_var.get(),
            "email_enabled": self.email_enabled.get(),
            "smtp_server": self.smtp_server.get(),
            "smtp_port": self.smtp_port.get(),
            "sender_email": self.sender_email.get(),
            "recipient_email": self.recipient_email.get(),
            "notify_green": self.notify_green.get(),
            "notify_red": self.notify_red.get(),
            "email_start_hour": self.email_start_hour.get(),
            "email_rate_limit": self.email_rate_limit.get(),
        }
        
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
            return True
        except Exception as e:
            print(f"Save error: {e}")
            return False
            
    def load_config(self):
        """Load saved config"""
        if not os.path.exists(self.CONFIG_FILE):
            return
        
        try:
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)
            
            if config.get("region"):
                self.region = tuple(config["region"])
                x1, y1, x2, y2 = self.region
                self.region_label.config(
                    text=f"✓ Region: ({x1}, {y1}) to ({x2}, {y2})",
                    foreground="green"
                )
            
            if config.get("status_position"):
                self.status_position = tuple(config["status_position"])
                self.calib_label.config(
                    text=f"✓ Status position: {self.status_position}",
                    foreground="green"
                )
            
            self.person_entry.delete(0, tk.END)
            self.person_entry.insert(0, config.get("target_person", ""))
            
            self.tesseract_path.delete(0, tk.END)
            self.tesseract_path.insert(0, config.get("tesseract_path", ""))
            
            self.interval_var.set(config.get("interval", "3"))
            self.email_enabled.set(config.get("email_enabled", False))
            
            self.smtp_server.delete(0, tk.END)
            self.smtp_server.insert(0, config.get("smtp_server", "smtp.gmail.com"))
            
            self.smtp_port.delete(0, tk.END)
            self.smtp_port.insert(0, config.get("smtp_port", "587"))
            
            self.sender_email.delete(0, tk.END)
            self.sender_email.insert(0, config.get("sender_email", ""))
            
            self.recipient_email.delete(0, tk.END)
            self.recipient_email.insert(0, config.get("recipient_email", ""))
            
            self.notify_green.set(config.get("notify_green", True))
            self.notify_red.set(config.get("notify_red", True))
            self.email_start_hour.set(config.get("email_start_hour", "9"))
            self.email_rate_limit.set(config.get("email_rate_limit", "60"))
            
        except Exception as e:
            print(f"Load error: {e}")
    
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    app = ChatStatusMonitorV2()
    app.run()


if __name__ == "__main__":
    main()