"""
Chat Status Monitor - GUI Version
==================================
A user-friendly version with a configuration window.

Requirements:
    pip install pyautogui opencv-python numpy pytesseract mss Pillow

Also install Tesseract OCR:
    Download from: https://github.com/UB-Mannheim/tesseract/wiki
"""

import cv2
import numpy as np
import pytesseract
from mss import mss
from PIL import Image, ImageTk
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
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except ImportError:
    from backports.zoneinfo import ZoneInfo  # For older Python


class ConfigWindow:
    """Configuration and control window"""
    
    CONFIG_FILE = "monitor_config.json"
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Chat Status Monitor")
        self.root.geometry("500x850")
        self.root.resizable(True, True)
        self.root.minsize(500, 850)
        
        self.monitor = None
        self.overlay = None
        self.running = False
        self.monitor_thread = None
        self.last_email_time = None  # Track when last email was sent
        self.last_notified_status = None  # Track what status we last notified about
        
        self.setup_ui()
        self.load_config()
        
    def setup_ui(self):
        """Setup the configuration UI"""
        # Main frame with padding
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="🔍 Chat Status Monitor",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=(0, 15))
        
        # === Target Person Section ===
        person_frame = ttk.LabelFrame(main_frame, text="Target Person", padding="10")
        person_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(person_frame, text="Name to monitor:").pack(anchor=tk.W)
        self.person_entry = ttk.Entry(person_frame, width=50)
        self.person_entry.pack(fill=tk.X, pady=2)
        self.person_entry.insert(0, "Fabian Thomas")
        
        # === Email Settings Section ===
        email_frame = ttk.LabelFrame(main_frame, text="Email Notifications", padding="10")
        email_frame.pack(fill=tk.X, pady=5)
        
        # Enable email checkbox
        self.email_enabled = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            email_frame, 
            text="Enable email notifications",
            variable=self.email_enabled
        ).pack(anchor=tk.W)
        
        # SMTP Server
        ttk.Label(email_frame, text="SMTP Server:").pack(anchor=tk.W, pady=(5, 0))
        self.smtp_server = ttk.Entry(email_frame, width=50)
        self.smtp_server.pack(fill=tk.X, pady=2)
        self.smtp_server.insert(0, "smtp.gmail.com")
        
        # SMTP Port
        ttk.Label(email_frame, text="SMTP Port:").pack(anchor=tk.W)
        self.smtp_port = ttk.Entry(email_frame, width=50)
        self.smtp_port.pack(fill=tk.X, pady=2)
        self.smtp_port.insert(0, "587")
        
        # Sender email
        ttk.Label(email_frame, text="Sender Email:").pack(anchor=tk.W)
        self.sender_email = ttk.Entry(email_frame, width=50)
        self.sender_email.pack(fill=tk.X, pady=2)
        
        # Password
        ttk.Label(email_frame, text="App Password:").pack(anchor=tk.W)
        self.sender_password = ttk.Entry(email_frame, width=50, show="*")
        self.sender_password.pack(fill=tk.X, pady=2)
        
        # Recipient email
        ttk.Label(email_frame, text="Recipient Email:").pack(anchor=tk.W)
        self.recipient_email = ttk.Entry(email_frame, width=50)
        self.recipient_email.pack(fill=tk.X, pady=2)
        
        # === Tesseract Section ===
        tesseract_frame = ttk.LabelFrame(main_frame, text="Tesseract OCR", padding="10")
        tesseract_frame.pack(fill=tk.X, pady=5)
        
        path_frame = ttk.Frame(tesseract_frame)
        path_frame.pack(fill=tk.X)
        
        ttk.Label(path_frame, text="Path:").pack(side=tk.LEFT)
        self.tesseract_path = ttk.Entry(path_frame, width=40)
        self.tesseract_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.tesseract_path.insert(0, r"C:\Program Files\Tesseract-OCR\tesseract.exe")
        
        browse_btn = ttk.Button(path_frame, text="Browse", command=self.browse_tesseract)
        browse_btn.pack(side=tk.LEFT)
        
        # === Monitor Settings ===
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.pack(fill=tk.X, pady=5)
        
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.pack(fill=tk.X)
        
        ttk.Label(interval_frame, text="Check interval (seconds):").pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value="3")
        interval_spin = ttk.Spinbox(
            interval_frame, 
            from_=1, 
            to=60, 
            width=5,
            textvariable=self.interval_var
        )
        interval_spin.pack(side=tk.LEFT, padx=5)
        
        # Notify on statuses
        self.notify_green = tk.BooleanVar(value=True)
        self.notify_red = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(
            settings_frame, 
            text="Notify when status is GREEN (online)",
            variable=self.notify_green
        ).pack(anchor=tk.W)
        
        ttk.Checkbutton(
            settings_frame, 
            text="Notify when status is RED (do not disturb)",
            variable=self.notify_red
        ).pack(anchor=tk.W)
        
        # Email schedule settings
        ttk.Separator(settings_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        ttk.Label(settings_frame, text="Email Schedule (Berlin Time):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        
        hour_frame = ttk.Frame(settings_frame)
        hour_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(hour_frame, text="Only send emails after:").pack(side=tk.LEFT)
        self.email_start_hour = tk.StringVar(value="9")
        hour_spin = ttk.Spinbox(
            hour_frame, 
            from_=0, 
            to=23, 
            width=3,
            textvariable=self.email_start_hour
        )
        hour_spin.pack(side=tk.LEFT, padx=5)
        ttk.Label(hour_frame, text=":00 (24h format)").pack(side=tk.LEFT)
        
        rate_frame = ttk.Frame(settings_frame)
        rate_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(rate_frame, text="Minimum time between emails:").pack(side=tk.LEFT)
        self.email_rate_limit = tk.StringVar(value="60")
        rate_spin = ttk.Spinbox(
            rate_frame, 
            from_=1, 
            to=1440, 
            width=5,
            textvariable=self.email_rate_limit
        )
        rate_spin.pack(side=tk.LEFT, padx=5)
        ttk.Label(rate_frame, text="minutes").pack(side=tk.LEFT)
        
        # === Status Display ===
        status_frame = ttk.LabelFrame(main_frame, text="Status", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.status_label = ttk.Label(
            status_frame, 
            text="⏸️ Not running",
            font=("Arial", 12)
        )
        self.status_label.pack()
        
        self.detection_label = ttk.Label(
            status_frame,
            text="",
            font=("Arial", 10)
        )
        self.detection_label.pack()
        
        # === Control Buttons ===
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        self.start_btn = ttk.Button(
            button_frame, 
            text="▶️ Start Monitoring",
            command=self.toggle_monitoring
        )
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        test_btn = ttk.Button(
            button_frame,
            text="🧪 Test Detection",
            command=self.test_detection
        )
        test_btn.pack(side=tk.LEFT, padx=5)
        
        calibrate_btn = ttk.Button(
            button_frame,
            text="🎯 Calibrate",
            command=self.start_calibration
        )
        calibrate_btn.pack(side=tk.LEFT, padx=5)
        
        save_btn = ttk.Button(
            button_frame,
            text="💾 Save Config",
            command=self.save_config
        )
        save_btn.pack(side=tk.LEFT, padx=5)
        
        # Calibration offset storage
        self.status_offset_x = -40  # Default: 40px left of name
        self.status_offset_y = 5    # Default: 5px down from name top
        
    def browse_tesseract(self):
        """Open file dialog to select Tesseract executable"""
        path = filedialog.askopenfilename(
            title="Select Tesseract Executable",
            filetypes=[("Executable", "*.exe"), ("All files", "*.*")]
        )
        if path:
            self.tesseract_path.delete(0, tk.END)
            self.tesseract_path.insert(0, path)
            
    def save_config(self):
        """Save configuration to file"""
        if self.save_config_silent():
            messagebox.showinfo("Saved", "Configuration saved successfully!")
    
    def save_config_silent(self):
        """Save configuration to file without showing message"""
        config = {
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
            "status_offset_x": getattr(self, 'status_offset_x', -50),
            "status_offset_y": getattr(self, 'status_offset_y', 10),
        }
        
        try:
            with open(self.CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save config: {e}")
            return False
            
    def load_config(self):
        """Load configuration from file"""
        if not os.path.exists(self.CONFIG_FILE):
            return
            
        try:
            with open(self.CONFIG_FILE, 'r') as f:
                config = json.load(f)
                
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
            
            # Load email schedule settings
            self.email_start_hour.set(config.get("email_start_hour", "9"))
            self.email_rate_limit.set(config.get("email_rate_limit", "60"))
            
            # Load calibration offsets
            self.status_offset_x = config.get("status_offset_x", -40)
            self.status_offset_y = config.get("status_offset_y", 5)
            print(f"Loaded calibration offset: ({self.status_offset_x}, {self.status_offset_y})")
            
        except Exception as e:
            print(f"Failed to load config: {e}")
            
    def test_detection(self):
        """Test detection without continuous monitoring"""
        self.update_status("🔄 Testing detection...")
        
        # Set tesseract path
        tess_path = self.tesseract_path.get()
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path
        
        try:
            # Capture screen
            with mss() as sct:
                monitor = sct.monitors[1]
                screenshot = sct.grab(monitor)
                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                screen = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            
            # Find target name
            target = self.person_entry.get()
            result = self.find_name(screen, target)
            
            if result:
                x, y, w, h = result
                status = self.detect_color(screen, result, save_debug=True)  # Save debug images
                
                self.update_status(f"✅ Found: {target}")
                self.detection_label.config(
                    text=f"Location: ({x}, {y}) - Status: {status.upper()}"
                )
                
                # Show overlay with debugging info
                self.show_test_overlay(x, y, w, h, status, screen)
                
                # Note: debug images are saved - user can check them if needed
                print(f"Debug images saved to current folder (debug_region.png, debug_context.png)")
                
            else:
                self.update_status(f"❌ Not found: {target}")
                self.detection_label.config(text="Try adjusting the name or check if visible on screen")
                
        except Exception as e:
            self.update_status(f"❌ Error: {str(e)[:30]}...")
            messagebox.showerror("Detection Error", str(e))
            
    def show_test_overlay(self, x, y, w, h, status, screen_image=None):
        """Show overlay for test detection - stays until clicked"""
        # Main rectangle around the name
        overlay = tk.Toplevel(self.root)
        overlay.attributes('-alpha', 0.8)
        overlay.attributes('-topmost', True)
        overlay.overrideredirect(True)
        
        padding = 5
        border_width = 3
        total_w = w + padding * 2 + border_width * 2
        total_h = h + padding * 2 + border_width * 2
        
        overlay.geometry(f"{total_w}x{total_h}+{x - padding - border_width}+{y - padding - border_width}")
        
        color = "lime" if status == "green" else "red" if status == "red" else "yellow" if status == "yellow" else "purple"
        
        # Create frame border effect
        canvas = tk.Canvas(overlay, width=total_w, height=total_h, highlightthickness=0)
        canvas.pack()
        
        # Draw colored border rectangle
        canvas.create_rectangle(
            border_width, border_width,
            total_w - border_width, total_h - border_width,
            outline=color, width=border_width
        )
        
        # Make center transparent/clickable by filling with a neutral color we'll make transparent
        canvas.configure(bg='gray1')
        overlay.wm_attributes('-transparentcolor', 'gray1')
        
        # Show the status icon search region using CLAMPED offset (same as detect_color)
        status_overlay = tk.Toplevel(self.root)
        status_overlay.attributes('-alpha', 0.5)
        status_overlay.attributes('-topmost', True)
        status_overlay.overrideredirect(True)
        
        # Use clamped offset (same logic as detect_color)
        offset_x = getattr(self, 'status_offset_x', -50)
        offset_y = getattr(self, 'status_offset_y', 10)
        offset_x = max(-100, min(-20, offset_x))
        offset_y = max(-20, min(30, offset_y))
        
        region_size = 20
        icon_x = max(0, x + offset_x - region_size)
        icon_y = max(0, y + offset_y - region_size)
        icon_w = region_size * 2
        icon_h = region_size * 2
        
        status_overlay.geometry(f"{icon_w}x{icon_h}+{icon_x}+{icon_y}")
        
        status_canvas = tk.Canvas(status_overlay, width=icon_w, height=icon_h, bg='blue', highlightthickness=2, highlightbackground='cyan')
        status_canvas.pack()
        
        # Info window
        info = tk.Toplevel(self.root)
        info.attributes('-topmost', True)
        info.title("Detection Result")
        info.geometry(f"350x200+{x + w + 20}+{y}")
        
        info_text = f"""
Name found at: ({x}, {y})
Size: {w} x {h}
Status detected: {status.upper()}

Offset used: ({offset_x}, {offset_y})
(Clamped to reasonable range)

Green rectangle = Name location
Blue box = Status search area

Click "Close All" to dismiss
        """
        
        label = tk.Label(info, text=info_text, justify=tk.LEFT, padx=10, pady=10, font=("Arial", 9))
        label.pack()
        
        # Close all overlays on click
        def close_all(event=None):
            try:
                overlay.destroy()
            except:
                pass
            try:
                status_overlay.destroy()
            except:
                pass
            try:
                info.destroy()
            except:
                pass
        
        overlay.bind('<Button-1>', close_all)
        status_overlay.bind('<Button-1>', close_all)
        info.bind('<Button-1>', close_all)
        
        close_btn = tk.Button(info, text="Close All", command=close_all)
        close_btn.pack(pady=5)
        
        # Auto-close after 30 seconds as fallback
        overlay.after(30000, lambda: overlay.destroy() if overlay.winfo_exists() else None)
        status_overlay.after(30000, lambda: status_overlay.destroy() if status_overlay.winfo_exists() else None)
        info.after(30000, lambda: info.destroy() if info.winfo_exists() else None)
        
    def find_name(self, image: np.ndarray, target: str) -> Optional[Tuple[int, int, int, int]]:
        """Find name in image using OCR - looks for name in chat LIST (left side), not header"""
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        
        try:
            data = pytesseract.image_to_data(gray, output_type=pytesseract.Output.DICT)
        except Exception as e:
            print(f"OCR Error: {e}")
            return None
        
        target_words = target.lower().split()
        
        # Create alternative versions of each word (for special characters)
        def get_word_variants(word):
            """Get possible OCR variations of a word"""
            variants = [word]
            if 'ß' in word:
                variants.append(word.replace('ß', 'ss'))
                variants.append(word.replace('ß', 'b'))
                variants.append(word.replace('ß', 'fs'))
                variants.append(word.replace('ß', '8'))
            variants.append(word.replace('ü', 'u').replace('ä', 'a').replace('ö', 'o'))
            return variants
        
        n_boxes = len(data['text'])
        
        # Collect all candidate matches
        all_matches = []
        first_word = target_words[0]
        
        for i in range(n_boxes):
            text = data['text'][i].strip()
            if not text or len(text) < 2:
                continue
            
            text_lower = text.lower()
            
            # Check if this matches the first word (ignore trailing punctuation like "Arne:")
            clean_text = text_lower.rstrip(':.,;!?')
            if first_word in clean_text or clean_text.startswith(first_word[:3]):
                loc = {
                    'index': i,
                    'text': text,
                    'x': data['left'][i],
                    'y': data['top'][i],
                    'w': data['width'][i],
                    'h': data['height'][i],
                }
                
                # Look for second word nearby
                if len(target_words) >= 2:
                    second_word = target_words[1]
                    second_variants = get_word_variants(second_word)
                    
                    for j in range(max(0, i-2), min(n_boxes, i+5)):
                        if j == i:
                            continue
                        nearby_text = data['text'][j].strip()
                        if not nearby_text:
                            continue
                        
                        nearby_lower = nearby_text.lower()
                        j_y = data['top'][j]
                        j_x = data['left'][j]
                        
                        # Must be on same line and to the right
                        if abs(j_y - loc['y']) < 25 and j_x > loc['x']:
                            for variant in second_variants:
                                if (variant in nearby_lower or 
                                    nearby_lower.startswith(variant[:4]) or
                                    variant.startswith(nearby_lower[:4])):
                                    
                                    # Found full name match!
                                    combined_x = loc['x']
                                    combined_y = min(loc['y'], j_y)
                                    combined_w = (j_x + data['width'][j]) - loc['x']
                                    combined_h = max(loc['h'], data['height'][j])
                                    
                                    all_matches.append({
                                        'x': combined_x,
                                        'y': combined_y,
                                        'w': combined_w,
                                        'h': combined_h,
                                        'full_match': True,
                                        'text': f"{text} {nearby_text}"
                                    })
                                    break
                
                # Also add partial match
                all_matches.append({
                    'x': loc['x'],
                    'y': loc['y'],
                    'w': loc['w'],
                    'h': loc['h'],
                    'full_match': False,
                    'text': text
                })
        
        if not all_matches:
            print(f"No matches found for '{target}'")
            return None
        
        # Prefer full matches
        full_matches = [m for m in all_matches if m['full_match']]
        candidates = full_matches if full_matches else all_matches
        
        # IMPORTANT: Prefer matches on the LEFT side of screen (sidebar/chat list)
        # Chat list is typically in the left 400 pixels
        sidebar_matches = [m for m in candidates if m['x'] < 400]
        
        if sidebar_matches:
            # Take the topmost match in the sidebar (most likely the correct chat entry)
            best = min(sidebar_matches, key=lambda m: m['y'])
            print(f"Found in sidebar: '{best['text']}' at ({best['x']}, {best['y']})")
        else:
            # Fallback to leftmost match
            best = min(candidates, key=lambda m: m['x'])
            print(f"Found (not in sidebar): '{best['text']}' at ({best['x']}, {best['y']})")
        
        return (best['x'], best['y'], best['w'], best['h'])
        
    def detect_color(self, image: np.ndarray, box: Tuple[int, int, int, int], save_debug: bool = False) -> str:
        """Detect status color - looks to the LEFT of the name for the avatar's status dot"""
        x, y, w, h = box
        
        # In Teams/Discord, the status dot is on the avatar, which is LEFT of the name
        # Typically 30-80 pixels to the left, and vertically centered with the name
        # We'll search a region to the left of the name
        
        # Use calibrated offset if reasonable, otherwise use defaults
        offset_x = getattr(self, 'status_offset_x', -50)
        offset_y = getattr(self, 'status_offset_y', 10)
        
        # CLAMP offsets to reasonable values (status dot should be close to name)
        offset_x = max(-100, min(-20, offset_x))  # Must be 20-100 pixels to the LEFT
        offset_y = max(-20, min(30, offset_y))     # Within 20 pixels vertically
        
        # Search region: to the left of the name
        region_size = 20
        search_x = max(0, x + offset_x - region_size)
        search_y = max(0, y + offset_y - region_size)
        search_w = region_size * 2
        search_h = region_size * 2
        
        # Ensure we don't go out of bounds
        img_h, img_w = image.shape[:2]
        search_x = min(search_x, img_w - search_w)
        search_y = min(search_y, img_h - search_h)
        
        region = image[search_y:search_y + search_h, search_x:search_x + search_w]
        
        if region.size == 0:
            print(f"ERROR: Empty region at ({search_x}, {search_y})")
            return "unknown"
        
        # Save debug image if requested
        if save_debug:
            debug_path = os.path.join(os.path.dirname(__file__) or ".", "debug_region.png")
            cv2.imwrite(debug_path, region)
            print(f"DEBUG: Saved search region to {debug_path}")
            
            # Also save a larger context image showing where we're looking
            ctx_x = max(0, x - 120)
            ctx_y = max(0, y - 20)
            ctx_w = w + 140
            ctx_h = h + 40
            ctx_region = image[ctx_y:ctx_y + ctx_h, ctx_x:ctx_x + ctx_w].copy()
            
            # Draw rectangle on context showing search area (YELLOW)
            cv2.rectangle(
                ctx_region,
                (search_x - ctx_x, search_y - ctx_y),
                (search_x - ctx_x + search_w, search_y - ctx_y + search_h),
                (0, 255, 255),  # Yellow
                2
            )
            # Draw rectangle showing name location (GREEN)
            cv2.rectangle(
                ctx_region,
                (x - ctx_x, y - ctx_y),
                (x - ctx_x + w, y - ctx_y + h),
                (0, 255, 0),  # Green
                2
            )
            
            context_path = os.path.join(os.path.dirname(__file__) or ".", "debug_context.png")
            cv2.imwrite(context_path, ctx_region)
            print(f"DEBUG: Saved context image to {context_path}")
        
        hsv = cv2.cvtColor(region, cv2.COLOR_BGR2HSV)
        
        # Calculate average values for debugging
        avg_h = np.mean(hsv[:, :, 0])
        avg_s = np.mean(hsv[:, :, 1])
        avg_v = np.mean(hsv[:, :, 2])
        
        # Green - online status
        green_mask = cv2.inRange(hsv, np.array([35, 80, 80]), np.array([85, 255, 255]))
        green_px = cv2.countNonZero(green_mask)
        
        # Red - DND/busy status
        red1 = cv2.inRange(hsv, np.array([0, 80, 80]), np.array([10, 255, 255]))
        red2 = cv2.inRange(hsv, np.array([160, 80, 80]), np.array([180, 255, 255]))
        red_mask = cv2.bitwise_or(red1, red2)
        red_px = cv2.countNonZero(red_mask)
        
        # Yellow/Orange - idle/away
        yellow_mask = cv2.inRange(hsv, np.array([15, 80, 80]), np.array([35, 255, 255]))
        yellow_px = cv2.countNonZero(yellow_mask)
        
        # Gray - offline
        gray_mask = cv2.inRange(hsv, np.array([0, 0, 60]), np.array([180, 30, 200]))
        gray_px = cv2.countNonZero(gray_mask)
        
        min_pixels = 5
        
        print(f"Color detection - Green: {green_px}, Red: {red_px}, Yellow: {yellow_px}, Gray: {gray_px}")
        print(f"  Name at: ({x}, {y}), Search at: ({search_x}, {search_y}) size {search_w}x{search_h}")
        print(f"  Average HSV: H={avg_h:.1f}, S={avg_s:.1f}, V={avg_v:.1f}")
        
        if green_px > min_pixels and green_px >= red_px and green_px >= yellow_px:
            return "green"
        elif red_px > min_pixels and red_px > green_px and red_px > yellow_px:
            return "red"
        elif yellow_px > min_pixels and yellow_px > green_px and yellow_px > red_px:
            return "yellow"
        elif gray_px > 20:
            return "offline"
        
        return "unknown"
        
    def toggle_monitoring(self):
        """Start or stop monitoring"""
        if self.running:
            self.stop_monitoring()
        else:
            self.start_monitoring()
    
    def start_calibration(self):
        """Start calibration mode - click on the status icon"""
        messagebox.showinfo(
            "Calibration", 
            "After clicking OK:\n\n"
            "1. Make sure the chat window is visible\n"
            "2. Click directly on the STATUS ICON (colored dot) of the person you want to monitor\n\n"
            "This helps the app know where to look for status colors."
        )
        
        # Hide main window temporarily
        self.root.withdraw()
        time.sleep(0.3)
        
        # Create fullscreen overlay to capture click
        cal_window = tk.Toplevel()
        cal_window.attributes('-fullscreen', True)
        cal_window.attributes('-alpha', 0.3)
        cal_window.attributes('-topmost', True)
        cal_window.configure(bg='blue')
        
        label = tk.Label(
            cal_window, 
            text="Click on the STATUS ICON (colored dot) next to the person's name\n\nPress ESC to cancel",
            font=("Arial", 20, "bold"),
            bg='blue',
            fg='white'
        )
        label.place(relx=0.5, rely=0.1, anchor='center')
        
        def on_click(event):
            click_x, click_y = event.x_root, event.y_root
            cal_window.destroy()
            self.root.deiconify()
            self.process_calibration_click(click_x, click_y)
        
        def on_escape(event):
            cal_window.destroy()
            self.root.deiconify()
        
        cal_window.bind('<Button-1>', on_click)
        cal_window.bind('<Escape>', on_escape)
        cal_window.focus_force()
    
    def process_calibration_click(self, click_x, click_y):
        """Process the calibration click to determine offset from name"""
        # Set tesseract path
        tess_path = self.tesseract_path.get()
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path
        
        try:
            # Capture screen
            with mss() as sct:
                monitor = sct.monitors[1]
                screenshot = sct.grab(monitor)
                img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                screen = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            
            # Find the name
            target = self.person_entry.get()
            result = self.find_name(screen, target)
            
            if result:
                name_x, name_y, name_w, name_h = result
                
                # Calculate offset from name to where user clicked
                raw_offset_x = click_x - name_x
                raw_offset_y = click_y - name_y
                
                # Validate: status dot should be LEFT of name (negative X) and close vertically
                if raw_offset_x > 0:
                    messagebox.showwarning(
                        "Calibration Warning",
                        f"You clicked to the RIGHT of the name.\n"
                        f"The status dot should be to the LEFT of the name.\n\n"
                        f"Offset: ({raw_offset_x}, {raw_offset_y})\n\n"
                        f"Using default offset instead."
                    )
                    self.status_offset_x = -50
                    self.status_offset_y = 10
                elif abs(raw_offset_x) > 150 or abs(raw_offset_y) > 50:
                    messagebox.showwarning(
                        "Calibration Warning",
                        f"Offset seems too large.\n"
                        f"Raw offset: ({raw_offset_x}, {raw_offset_y})\n\n"
                        f"The status dot should be 30-100 pixels left of the name.\n"
                        f"Using clamped values."
                    )
                    self.status_offset_x = max(-100, min(-20, raw_offset_x))
                    self.status_offset_y = max(-20, min(30, raw_offset_y))
                else:
                    self.status_offset_x = raw_offset_x
                    self.status_offset_y = raw_offset_y
                
                # Sample color at click location
                sample_region = screen[
                    max(0, click_y - 5):click_y + 5,
                    max(0, click_x - 5):click_x + 5
                ]
                
                if sample_region.size > 0:
                    hsv = cv2.cvtColor(sample_region, cv2.COLOR_BGR2HSV)
                    avg_h = np.mean(hsv[:, :, 0])
                    avg_s = np.mean(hsv[:, :, 1])
                    avg_v = np.mean(hsv[:, :, 2])
                    
                    color_name = "Unknown"
                    if avg_s > 50:  # Saturated color
                        if avg_h < 15 or avg_h > 160:
                            color_name = "RED"
                        elif 35 < avg_h < 85:
                            color_name = "GREEN"
                        elif 15 < avg_h < 35:
                            color_name = "YELLOW"
                    
                    messagebox.showinfo(
                        "Calibration Complete",
                        f"Calibration successful!\n\n"
                        f"Name found at: ({name_x}, {name_y})\n"
                        f"You clicked at: ({click_x}, {click_y})\n"
                        f"Stored offset: ({self.status_offset_x}, {self.status_offset_y})\n\n"
                        f"Color at click: {color_name}\n"
                        f"HSV: ({avg_h:.0f}, {avg_s:.0f}, {avg_v:.0f})\n\n"
                        f"Configuration auto-saved!"
                    )
                    
                    # Auto-save after calibration
                    self.save_config_silent()
                else:
                    messagebox.showwarning("Calibration", "Could not sample color at click location.")
            else:
                messagebox.showwarning(
                    "Calibration Failed",
                    f"Could not find '{target}' on screen.\n"
                    "Make sure the name is visible in the chat LIST (left sidebar) and try again."
                )
                
        except Exception as e:
            messagebox.showerror("Calibration Error", str(e))
            
    def start_monitoring(self):
        """Start the monitoring loop"""
        # Set tesseract path
        tess_path = self.tesseract_path.get()
        if os.path.exists(tess_path):
            pytesseract.pytesseract.tesseract_cmd = tess_path
        
        self.running = True
        self.start_btn.config(text="⏹️ Stop Monitoring")
        self.update_status("▶️ Monitoring...")
        
        self.last_status = None
        
        # Start monitoring thread
        self.monitor_thread = threading.Thread(target=self.monitor_loop, daemon=True)
        self.monitor_thread.start()
        
    def stop_monitoring(self):
        """Stop the monitoring loop"""
        self.running = False
        self.start_btn.config(text="▶️ Start Monitoring")
        self.update_status("⏸️ Stopped")
        
    def monitor_loop(self):
        """Background monitoring loop"""
        while self.running:
            try:
                # Capture screen
                with mss() as sct:
                    monitor = sct.monitors[1]
                    screenshot = sct.grab(monitor)
                    img = Image.frombytes("RGB", screenshot.size, screenshot.bgra, "raw", "BGRX")
                    screen = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
                
                target = self.person_entry.get()
                result = self.find_name(screen, target)
                
                if result:
                    x, y, w, h = result
                    status = self.detect_color(screen, result)
                    
                    # Update UI from main thread
                    self.root.after(0, lambda: self.update_status(f"✅ Found: {target}"))
                    self.root.after(0, lambda s=status: self.detection_label.config(
                        text=f"Status: {s.upper()}"
                    ))
                    
                    # Check for status change and send email
                    if status != self.last_status:
                        should_notify = (
                            (status == "green" and self.notify_green.get()) or
                            (status == "red" and self.notify_red.get())
                        )
                        
                        if should_notify and self.email_enabled.get():
                            self.send_notification(target, status)
                        
                        self.last_status = status
                else:
                    self.root.after(0, lambda: self.update_status(f"🔍 Searching for: {target}"))
                    self.root.after(0, lambda: self.detection_label.config(text="Not visible on screen"))
                    
            except Exception as e:
                self.root.after(0, lambda e=e: self.update_status(f"⚠️ Error: {str(e)[:20]}"))
            
            time.sleep(int(self.interval_var.get()))
            
    def can_send_email(self) -> Tuple[bool, str]:
        """Check if we can send an email based on time and rate limits"""
        # Get current time in Berlin
        try:
            berlin_tz = ZoneInfo("Europe/Berlin")
            now_berlin = datetime.now(berlin_tz)
        except Exception:
            # Fallback: just use local time
            now_berlin = datetime.now()
            print("Warning: Could not get Berlin timezone, using local time")
        
        current_hour = now_berlin.hour
        start_hour = int(self.email_start_hour.get())
        
        # Check if it's after the start hour
        if current_hour < start_hour:
            return False, f"Before {start_hour}:00 Berlin time (current: {current_hour}:00)"
        
        # Check rate limit
        rate_limit_minutes = int(self.email_rate_limit.get())
        if self.last_email_time:
            time_since_last = datetime.now() - self.last_email_time
            minutes_since_last = time_since_last.total_seconds() / 60
            
            if minutes_since_last < rate_limit_minutes:
                remaining = rate_limit_minutes - minutes_since_last
                return False, f"Rate limited: {remaining:.0f} min until next email allowed"
        
        return True, "OK"
    
    def send_notification(self, name: str, status: str):
        """Send email notification with time and rate limit checks"""
        # Check if we can send
        can_send, reason = self.can_send_email()
        
        if not can_send:
            print(f"Email NOT sent: {reason}")
            return
        
        # Check if this is the same status we already notified about
        if self.last_notified_status == status:
            print(f"Email NOT sent: Already notified about {status} status")
            return
        
        try:
            # Get Berlin time for the email
            try:
                berlin_tz = ZoneInfo("Europe/Berlin")
                berlin_time = datetime.now(berlin_tz).strftime('%Y-%m-%d %H:%M:%S %Z')
            except Exception:
                berlin_time = time.strftime('%Y-%m-%d %H:%M:%S') + " (local)"
            
            msg = MIMEMultipart()
            msg['From'] = self.sender_email.get()
            msg['To'] = self.recipient_email.get()
            msg['Subject'] = f"Status Alert: {name} is now {status.upper()}"
            
            body = f"""
Chat Status Alert
=================

Person: {name}
Status: {status.upper()}
Time: {berlin_time}

This is an automated notification.
Next email allowed after: {int(self.email_rate_limit.get())} minutes
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP(self.smtp_server.get(), int(self.smtp_port.get())) as server:
                server.starttls()
                server.login(self.sender_email.get(), self.sender_password.get())
                server.send_message(msg)
            
            # Update tracking
            self.last_email_time = datetime.now()
            self.last_notified_status = status
            
            print(f"✉️ Email SENT: {name} is {status} (Berlin time: {berlin_time})")
            
        except Exception as e:
            print(f"Email error: {e}")
            
    def update_status(self, text: str):
        """Update status label"""
        self.status_label.config(text=text)
        
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    app = ConfigWindow()
    app.run()


if __name__ == "__main__":
    main()