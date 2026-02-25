# 🔍 Chat Status Monitor

A Windows desktop application that monitors chat applications (Microsoft Teams, Discord, Slack, etc.) for a specific person's online status and sends email notifications when their status changes.

![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-green.svg)

---

## ✨ Features

- **🔍 OCR-based Name Detection** - Finds any person by name in your chat list, even if their position changes
- **🎨 Status Color Detection** - Detects green (online), red (busy/DND), yellow (away), and gray (offline)
- **📧 Email Notifications** - Sends alerts when the person's status changes
- **⏰ Smart Email Scheduling** - Only sends emails after a specified hour (e.g., 9 AM Berlin time)
- **🚫 Rate Limiting** - Prevents spam with configurable minimum time between emails (e.g., 1 hour)
- **🖼️ Visual Feedback** - Shows a rectangle around the found person's name
- **🎯 Calibration Tool** - Click on the status dot to teach the app where to look
- **💾 Persistent Settings** - Saves your configuration between sessions
- **🌍 Timezone Support** - Uses Berlin timezone for email scheduling

---

## 📋 Table of Contents

- [Requirements](#-requirements)
- [Installation](#-installation)
  - [Step 1: Install Tesseract OCR](#step-1-install-tesseract-ocr)
  - [Step 2: Install Python Dependencies](#step-2-install-python-dependencies)
- [Usage](#-usage)
- [Building Standalone Executable](#-building-standalone-executable)
- [Configuration](#-configuration)
- [Troubleshooting](#-troubleshooting)

---

## 📦 Requirements

- **Windows 10/11**
- **Python 3.8+** (for running from source)
- **Tesseract OCR** (required for text recognition)

---

## 🚀 Installation

### Step 1: Install Tesseract OCR

Tesseract OCR is required for the app to read text from your screen. Follow these steps:

#### Download Tesseract

1. Go to the **Tesseract at UB Mannheim** page:
   
   👉 https://github.com/UB-Mannheim/tesseract/wiki

2. Download the latest **Windows installer** (64-bit recommended):
   - Look for: `tesseract-ocr-w64-setup-v5.x.x.exe`

#### Install Tesseract

3. Run the downloaded installer

4. **Important Installation Options:**
   - ✅ Use the default installation path: `C:\Program Files\Tesseract-OCR\`
   - ✅ Select "Add to PATH" if the option is available
   - ✅ You can leave the default language packs (English is sufficient)

5. Click through the installer to complete installation

#### Verify Installation

6. Open **Command Prompt** and type:
   ```cmd
   tesseract --version
   ```
   
   You should see output like:
   ```
   tesseract v5.3.1
   ...
   ```

> **Note:** If the command is not found, you may need to add Tesseract to your PATH manually, or specify the full path in the app settings.

---

### Step 2: Install Python Dependencies

#### Option A: Using pip (Recommended)

1. Open **Command Prompt** or **PowerShell**

2. Navigate to the app folder:
   ```cmd
   cd C:\path\to\ChatStatusMonitor
   ```

3. Install all dependencies:
   ```cmd
   pip install -r requirements.txt
   ```

   Or install manually:
   ```cmd
   pip install pyautogui opencv-python numpy pytesseract mss Pillow tzdata
   ```

#### Option B: Using a Virtual Environment

```cmd
# Create virtual environment
python -m venv .venv

# Activate it
.venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## 🎮 Usage

### Running the Application

```cmd
python chat_monitor_gui.py
```

### First-Time Setup

1. **Enter the person's name** you want to monitor (e.g., "Arne Kaulfuß")

2. **Set Tesseract path** (if not auto-detected):
   - Default: `C:\Program Files\Tesseract-OCR\tesseract.exe`
   - Use the "Browse" button if installed elsewhere

3. **Select chat list and Calibrate the status detection:**
   - Put your chat app on any monitor
   - Click "📐 Select Region"
   - You'll see all your monitors in the overlay
   - Drag a rectangle around the chat list (on whichever monitor it's on)
   - Click **"🎯 Calibrate"**
   - When the pointer changed to **+**, **click directly on the status dot** (the colored circle) next to the person's name in your chat sidebar
   - This teaches the app where to look for the status color

4. **Test the detection:**
   - Click **"🧪 Test Detection"**
   - You should see a Preview image with:
     - A **green rectangle** around the person's name
     - A **blue box** on their status dot
     - Status detected: GREEN, RED, YELLOW, or UNKNOWN

5. **Configure email settings** (optional):
   - Enable email notifications
   - Enter SMTP settings (see [Email Configuration](#email-configuration))

6. **Click "▶️ Start Monitoring"** to begin!

### Email Schedule Settings

| Setting | Description | Default |
|---------|-------------|---------|
| Only send emails after | Hour in 24h format (Berlin time) | 9 (9:00 AM) |
| Minimum time between emails | Rate limit in minutes | 60 (1 hour) |

This prevents email spam - you'll get maximum 1 email per hour, and only during work hours.

---

## 📦 Building Standalone Executable

You can create a standalone `.exe` file that runs without Python installed.

### Automatic Build (Recommended)

1. **Double-click `BUILD.bat`**

2. Wait for the build to complete (may take 1-2 minutes)

3. Find your executable at:
   ```
   dist\ChatStatusMonitor.exe
   ```

### Manual Build

If the batch file doesn't work, run these commands manually:

```cmd
# Navigate to the app folder
cd C:\path\to\ChatStatusMonitor

# Install PyInstaller
pip install pyinstaller

# Install all dependencies
pip install pyautogui opencv-python numpy pytesseract mss Pillow tzdata

# Build the executable
pyinstaller --name=ChatStatusMonitor --onefile --windowed --noconfirm --clean chat_monitor_gui.py
```

The executable will be created at `dist\ChatStatusMonitor.exe`

### Using the Standalone Executable

After building, you can:

| Action | Supported |
|--------|-----------|
| Copy to USB drive | ✅ |
| Run on another Windows PC | ✅ |
| No Python needed on target PC | ✅ |
| Share with coworkers | ✅ |

**Requirements on target computer:**
- Windows 10/11
- Tesseract OCR installed at `C:\Program Files\Tesseract-OCR\`

**Not required on target computer:**
- ❌ Python
- ❌ Any Python packages
- ❌ Source code files

---

## ⚙️ Configuration

### Email Configuration

To receive email notifications, configure the following:

| Setting | Description | Example |
|---------|-------------|---------|
| SMTP Server | Your email provider's SMTP server | `smtp.gmail.com` |
| SMTP Port | Usually 587 for TLS | `587` |
| Sender Email | Your email address | `your.email@gmail.com` |
| App Password | See below for Gmail | `xxxx xxxx xxxx xxxx` |
| Recipient Email | Where to send alerts | `alerts@example.com` |

#### Gmail Setup

Gmail requires an **App Password** (not your regular password):

1. Go to your Google Account → **Security**
2. Enable **2-Step Verification** (if not already enabled)
3. Go to **App Passwords** (search for it in account settings)
4. Generate a new app password for "Mail"
5. Use this 16-character password in the app

#### Other Email Providers

| Provider | SMTP Server | Port |
|----------|-------------|------|
| Gmail | smtp.gmail.com | 587 |
| Outlook/Hotmail | smtp.office365.com | 587 |
| Yahoo | smtp.mail.yahoo.com | 587 |

### Configuration File

Settings are saved to `monitor_config.json` in the same folder as the app:

```json
{
  "target_person": "Arne Kaulfuß",
  "tesseract_path": "C:\\Program Files\\Tesseract-OCR\\tesseract.exe",
  "interval": "3",
  "email_enabled": true,
  "smtp_server": "smtp.gmail.com",
  "smtp_port": "587",
  "sender_email": "your.email@gmail.com",
  "recipient_email": "alerts@example.com",
  "notify_green": true,
  "notify_red": true,
  "email_start_hour": "9",
  "email_rate_limit": "60",
  "status_offset_x": -50,
  "status_offset_y": 10
}
```

---

## 🔧 Troubleshooting

### Tesseract Issues

**"Tesseract not found"**
- Make sure Tesseract is installed
- Check the path in the app: `C:\Program Files\Tesseract-OCR\tesseract.exe`
- Use the "Browse" button to locate `tesseract.exe`

**OCR not recognizing text**
- Make sure the chat window is visible and not minimized
- Try increasing screen brightness/contrast
- Special characters (like ß, ü, ä) may be read differently by OCR

### Name Detection Issues

**"Name not found"**
- Make sure the person's name is visible in the chat **sidebar** (left panel)
- The app looks for names on the left side of the screen (X < 400 pixels)
- Try using just the first name if full name isn't found
- Check console output to see what OCR is detecting

**Wrong person detected**
- Make sure you're using the full name to avoid matching partial names
- The app prioritizes matches in the sidebar over the chat header

### Status Detection Issues

**"Status: UNKNOWN"**
- Click **"🎯 Calibrate"** and click on the actual status dot
- The status dot must be visible (not covered by other windows)
- Check `debug_context.png` to see where the app is looking

**Status detected incorrectly**
- Recalibrate by clicking on the status dot
- Make sure you're clicking on the colored circle, not the avatar

### Email Issues

**"Email NOT sent: Before 9:00 Berlin time"**
- This is expected! Emails only send after the configured hour
- Change the start hour in settings if needed

**"Email NOT sent: Rate limited"**
- This is expected! Only 1 email per hour (configurable)
- Wait for the rate limit to expire or adjust the setting

**"Email error: Authentication failed"**
- For Gmail: Use an App Password, not your regular password
- Check that 2-Step Verification is enabled on your Google account

### Build Issues

**"Windows protected your PC" when running .exe**
- Click "More info" → "Run anyway"
- This happens because the .exe is not digitally signed

**Antivirus blocks the file**
- PyInstaller executables sometimes trigger false positives
- Add an exception for the file in your antivirus software

**Build fails with "Module not found"**
- Install all dependencies first:
  ```cmd
  pip install pyautogui opencv-python numpy pytesseract mss Pillow tzdata pyinstaller
  ```

---

## 📁 File Structure

```
ChatStatusMonitor/
├── chat_monitor_gui.py    # Main application (GUI version)
├── chat_monitor.py        # Command-line version
├── requirements.txt       # Python dependencies
├── BUILD.bat              # Windows build script
├── build_exe.py           # Python build script
├── README.md              # This file
├── STANDALONE_GUIDE.md    # Detailed build instructions
├── monitor_config.json    # Your saved settings (created on first run)
├── debug_region.png       # Debug: status search area (created during testing)
└── debug_context.png      # Debug: context around name (created during testing)
```

---

## 🔒 Privacy & Security

- The app captures screenshots of your screen to detect text and colors
- No data is sent anywhere except your configured email
- Settings are stored locally in `monitor_config.json`
- Email passwords are stored in plain text in the config file - keep it secure!

---

## 📝 License

This project is provided as-is for personal use.

---

## 🤝 Support

If you encounter issues:

1. Check the [Troubleshooting](#-troubleshooting) section
2. Look at the console output for error messages
3. Check `debug_region.png` and `debug_context.png` for visual debugging
4. Make sure Tesseract OCR is properly installed

---

Made with ❤️ for monitoring chat status