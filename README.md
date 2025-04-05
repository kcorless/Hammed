

# Hammed Email Filtering Tool - User Guide
 
## 📘 Overview
 
**Hammed** is a custom-built Outlook automation tool that helps manage and automatically respond to unsolicited emails. It works by maintaining a **whitelist of approved domains**, and optionally checking for a **secret word** in the subject line to bypass the filter. Emails not meeting those criteria are automatically moved to a special folder (`Hammed`), marked as read, and optionally replied to with a standard response.

**NOTE:** Hammed will only work on Microsoft Windows Outlook environment.  Sorry Mac users.
 
---
 
## ⚙️ Installation & Setup
 
### 1. Enable Developer Tools in Outlook
 
If you haven’t already:
 
- Open Outlook
- Go to **File > Options > Customize Ribbon**
- Check the **Developer** box in the right column and click OK
 
### 2. Open the VBA Editor
 
- Press `Alt + F11` to open the **Visual Basic for Applications (VBA)** editor
 
### 3. Import the Required Modules
 
In the VBA editor:
 
1. Go to **File > Import File...**
2. Import each of the following `.bas` or `.cls` files:
   - `ThisOutlookSession.cls`
   - `ConfigModule.bas`
   - `EmailHandlerModule.bas`
   - `DomainWhitelistModule.bas`
   - `LoggingModule.bas`
   - `EmailUtilsModule.bas`
   - `AddWLModule.bas`
 
After importing, ensure the module names (in the Properties window) match the filenames listed above.
 
### 4. Trust Access to the VBA Project
 
- Go to **File > Options > Trust Center > Trust Center Settings > Macro Settings**
- Check **"Enable all macros (not recommended...)"**
- Click OK
 
### 5. Restart Outlook
 
After setup is complete, **restart Outlook** to ensure `Application_Startup()` runs and initializes the tool.
 
---
 
## 🔧 Configuration
 
### Required Files
 
The Hammed tool expects the following files to be present in a user-specific folder:
 
```
D:\Users\<your-username>\AppData\Local\Hammed\
```
 
These include:
 
- `whitelist.txt` — List of approved domains (one per line). Note that the string in whitelist.txt is search for in your domain.  So if you have "goo" in the whitelist, it will whitelist google.com and googly.com and abcgoo.net
- `reply.txt` — Standard reply body for auto-responses
- `secret.txt` — Optional keyword to allow a message through
- `Hammed.ini` — Configuration file for the tool
 
### `Hammed.ini` Sample Configuration
 
```
NO_AUTO_SEND_REPLY_MODE=True
DEBUG_MODE=True
MSGBOX_MODE=False
```

 ### `whitelist.txt` Sample Configuration
 
```
gmail.com
united.com
apple
```
 
---
 
## 📬 What Hammed Does
 
### Workflow:
 
1. A new mail arrives in your Inbox.
2. If it's from a **whitelisted domain** or contains the **secret word** in the subject:
   - It is **left alone**
3. If not:
   - It is **moved to the `Hammed` folder**
   - Marked as **read**
   - A reply is optionally sent (or saved to Drafts, depending on config)
   - A full audit log is written to `HamDebugLog.txt`
   - If it **should** be in your whitelist but it is not, see the AddWL macro below
 
### Logging
 
Logs are stored in:
```
D:\Users\<your-username>\AppData\Local\Hammed\HamDebugLog.txt
```
 
If this file cannot be written to, a fallback `HamDebugLog2.txt` is used.
 
---
 
## ✅ Using the AddWL Macro
 
### What It Does
 
The `AddWL` macro adds the domain of the currently selected email's sender to the whitelist file (`whitelist.txt`).
 
### How to Use It
 
1. Select an email in your inbox.
2. Run the macro `AddWLModule.AddWL`
   - The domain of the sender will be added to `whitelist.txt`
   - You will receive a confirmation via message box.
 
### Add `AddWL` to Quick Access Toolbar
 
1. Go to **File > Options > Quick Access Toolbar**
2. From “Choose commands from,” select `Macros`
3. Choose `AddWLModule.AddWL`
4. Click **Add >>**
5. (Optional) Click **Modify...** to choose an icon and friendly name
6. Click OK
 
Now you can add senders to the whitelist with **one click** after the desired email is selected.
 
---
 
## 🧪 Testing
 
For testing, you can trigger the macro manually or simulate inbox arrivals via test emails.
 
If something goes wrong, check:
 
- **`HamDebugLog.txt`** for log output
- **`Hammed.ini`** for configuration issues
- That all required files exist in the expected directory
 
---
 
## 🙋‍♂️ Support
 
For updates or assistance, try debugging with your favorite LLM.  Or go to the project GitHub and request help.

If you like using this tool, there is no cost but you can buy me a cup of coffee: https://buymeacupofcoffee/goldenchimp

