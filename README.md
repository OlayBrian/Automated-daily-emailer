# Daily Report Tool

A Windows desktop utility that automates the process of pulling a formatted table from an Excel workbook and composing a daily Outlook email with that table embedded — preserving all cell formatting.

<img width="747" height="748" alt="Screenshot 2025-07-28 091924" src="https://github.com/user-attachments/assets/8fb9dc84-25b8-4b20-9d2e-beab6dbbdd3d" />


---

## Requirements

- **OS:** Windows only
- **Python:** 3.7+
- **Microsoft Office:** Excel and Outlook must be installed and configured
- **Python packages:**
  ```
  pip install pywin32
  ```
  > `tkinter` is included with standard Python installations.

---

## Setup

1. Clone or download the script to your machine.
2. Install the required dependency:
   ```bash
   pip install pywin32
   ```
3. Run the script:
   ```bash
   python "UI Automated Report.py"
   ```

---

## How to Use

### 1. Configure the Excel Source
- **Excel File Path** — Click **Browse** to select your `.xlsx` or `.xlsm` workbook.
- **Excel Table Range** — Enter the cell range to copy (e.g., `B8:H12`). This range is pulled from the sheet named `Transmission SLO Report`.

> The workbook will refresh all data connections automatically before the range is copied.

### 2. Configure the Email
| Field | Description |
|---|---|
| **To** | Primary recipient email address(es) |
| **CC** | CC recipient email address(es) |
| **Subject** | Defaults to today's date (e.g., `Daily Report – March 27, 2025`) |
| **Body (Before Table)** | Text inserted above the pasted table (default: `Good morning,`) |
| **Body (After Table)** | Text inserted below the pasted table (default: `Thank you.`) |

### 3. Run
Click **Run Report**. The tool will:
1. Open the workbook in Excel and refresh all data connections.
2. Copy the specified range (with formatting intact).
3. Open a new Outlook email draft and paste the table using Word Editor.
4. Leave the email open for your review before sending.

---

## Presets

You can save and reload your full configuration (file path, range, recipients, subject, body text) as a `.json` preset file.

- **Save Preset** — Saves current settings to a `.json` file in `~/Desktop/presets/`.
- **Load Preset** — Loads a previously saved `.json` preset, restoring all fields.

Preset files are stored at:
```
C:\Users\<YourName>\Desktop\presets\
```

---

## Notes

- The email is opened as a **draft** — it is never sent automatically. Always review before sending.
- Excel and Outlook will briefly become visible during execution; this is expected behavior.
- Multiple email addresses in **To** or **CC** fields should be separated by semicolons (`;`).
- The workbook must contain a sheet named exactly **`Transmission SLO Report`** for the script to work correctly.

---

## Troubleshooting

| Issue | Solution |
|---|---|
| `Failed to open workbook` | Verify the file path is correct and the file is not already open/locked. |
| Table not pasting with formatting | Ensure Outlook is fully open and not minimized before running. |
| Data not refreshing | Check that Excel data connections are valid and credentials are available. |
| `win32com` import error | Run `pip install pywin32` and restart your terminal. |
