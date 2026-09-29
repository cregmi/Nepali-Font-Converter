# Preeti ↔ Unicode Converter for Microsoft Word & Excel

A VBA-based tool to seamlessly convert text between **Preeti** (Nepali legacy font) and **Unicode** in Microsoft Word and Microsoft Excel.

This repository provides both ready-to-use add-in/template files for quick installation, as well as raw `.bas` source files for developers who prefer to build or inspect the code manually.

---

## 📂 Repository Contents

| File | Description |
| :--- | :--- |
| `unicode-preeti-conversion-macro.dotm` | Ready-to-use Word Macro-Enabled Template. |
| `unicode-preeti-conversion-add-in.xlam` | Ready-to-use Excel Add-In file. |
| `PreetiUnicode_Word.bas` | Raw VBA Module source code for Word. |
| `PreetiUnicode_Excel.bas` | Raw VBA Module source code for Excel. |
| `start-up-folders.txt` | Reference file containing startup directory paths. |

---

## ⚡ Method 1: Quick Setup (Recommended)

Use the pre-compiled `.dotm` and `.xlam` files for quick installation.

### Step 1: Download & Unblock Files
Due to Microsoft Office security policies, downloaded macro files from the internet are blocked by default. You **must** unblock them before installation:

1. Download `unicode-preeti-conversion-macro.dotm` and `unicode-preeti-conversion-add-in.xlam`.
2. Right-click the downloaded file and select **Properties**.
3. Under the **General** tab, look at the bottom for the **Security** section.
4. Check the box labeled **Unblock**.
5. Click **Apply**, then **OK**.

> ⚠️ *Repeat this step for both files.*

---

### Step 2: Copy Files to Startup Folders

Copy the unblocked files into their respective Microsoft Office startup folders:

#### 📝 For Word:
* **Destination:** `%APPDATA%\Microsoft\Word\STARTUP`
* **File to paste:** `unicode-preeti-conversion-macro.dotm`

#### 📊 For Excel:
* **Destination:** `%APPDATA%\Microsoft\Excel\XLSTART`
* **File to paste:** `unicode-preeti-conversion-add-in.xlam`

*(Tip: Press `Win + R`, paste the folder path above into the Run dialog, and press Enter to open the folder directly).*

---

### Step 3: Add Functions to Ribbon/Quick Access Toolbar

To easily run the macro/functions directly inside Word or Excel:

1. Open **Word** or **Excel**.
2. Go to **File** > **Options** > **Customize Ribbon** (or **Quick Access Toolbar**).
3. Under *Choose commands from*, select **Macros** (for Word) or **Add-Ins / Macros** (for Excel).
4. Create a new Group on your Ribbon tab or select an existing one.
5. Select the conversion functions and click **Add >>**.
6. Click **OK** to save.

---

## 🛠️ Method 2: Manual Setup via `.bas` Import (Developer Setup)

If you prefer to build the macro templates yourself using the raw source code:

1. Open Microsoft Word or Excel.
2. Press `Alt + F11` to open the **VBA Editor**.
3. Go to **File** > **Import File...** (`Ctrl + M`).
4. Select `PreetiUnicode_Word.bas` (for Word) or `PreetiUnicode_Excel.bas` (for Excel).
5. Save the file:
   * **In Word:** Save as **Word Macro-Enabled Template (`.dotm`)** inside `%APPDATA%\Microsoft\Word\STARTUP`.
   * **In Excel:** Save as **Excel Add-In (`.xlam`)** inside `%APPDATA%\Microsoft\Excel\XLSTART`.

---

## ❓ Troubleshooting & Security Notes

* **Security Warning / Red Bar:** If Office displays a security warning stating macros are blocked, ensure you performed **Step 1 (Unblocking files)** in file properties.
* **Trust Center Settings:** Go to **File** > **Options** > **Trust Center** > **Trust Center Settings** > **Macro Settings**, and ensure macros from startup locations are allowed to run.

---

## 🤖 Development & Acknowledgments

* **AI Assistance:** VBA code modules were optimized, debugged, and tested with the assistance of the Claude Sonnet 5 AI model, free tier service.

---

## 📄 License

This project is open-source and free to use or modify.
