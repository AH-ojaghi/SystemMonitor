# Advanced Terminal Activity Tracker for Windows

This Python script, created out of curiosity, tracks system activity on Windows, providing insights into running processes, system uptime, and resource usage in a clean, color-coded terminal interface. It aggregates processes with the same name (e.g., multiple Chrome instances) and allows exporting reports to Excel for further analysis.

## Features
- **System Information**: Displays system boot time, current time, and uptime in a human-readable format (`YYYY-MM-DD HH:MM:SS`).
- **Process Tracking**: Lists programs active for more than 1 minute, showing their name, active time, CPU usage (%), and memory usage (MB).
- **Color-Coded Output**: Highlights CPU and memory usage in the terminal (green for low, yellow for medium, red for high).
- **Process Aggregation**: Combines multiple instances of the same program (e.g., all `chrome.exe` processes) into a single entry with summed CPU and memory usage.
- **Excel Export**: Type `ED` to export the report to an Excel file. Users can specify a custom file path or use the default (`activity_report_YYYY-MM-DD_HH-MM-SS.xlsx`) in the current directory.
- **Filtering**: Excludes the "System Idle Process" and focuses on activities since system boot.

## Requirements
- **Python Version**: Python 3.6 or higher
- **Dependencies**:
  - `psutil`: For system and process information
  - `tabulate`: For formatting terminal tables
  - `colorama`: For colored terminal output
  - `openpyxl`: For Excel file generation
- Install dependencies using pip:
  ```bash
  pip install psutil tabulate colorama openpyxl
  ```

## Installation
1. Clone or download this repository:
   ```bash
   git clone https://github.com/AH-ojaghi/SystemMonitor.git
   ```
2. Navigate to the project directory:
   ```bash
   cd SystemMonitor
   ```
3. Install the required Python packages:
   ```bash
   pip install psutil tabulate colorama openpyxl
   ```

## Usage
1. Run the script:
   ```bash
   python SystemMonitor.py
   ```
2. The script will display:
   - System boot time, current time, and uptime.
   - A table of active programs (running > 1 minute) with their active time, CPU %, and memory usage (MB).
3. To export to Excel:
   - Type `ED` and press Enter.
   - Enter a custom file path or press Enter to use the default (`activity_report_YYYY-MM-DD_HH-MM-SS.xlsx`).
4. Check the output Excel file for the report.

## Example Output
```
System booted at: 2025-08-30 08:00:00
Current time: 2025-08-30 20:48:00
Uptime: 12:48:00

Active Programs (more than 1 minute):
Program Name    Active Time    CPU %    Memory MB
--------------  -------------  -------  ---------
chrome.exe      12:00:00       15.2%    850.5 MB
code.exe        10:30:00       5.0%     320.1 MB
explorer.exe    12:45:00       0.5%     50.2 MB

Type ED and press Enter to export report to Excel (or any other key to exit):
```

## Notes
- **Windows-Specific**: The script is designed for Windows due to its reliance on `psutil` for process handling.
- **Active Time Calculation**: Uses the earliest start time of aggregated processes for accuracy.
- **Error Handling**: If you encounter a "Permission denied" error when saving the Excel file, ensure you have write access to the specified directory or run the terminal as an administrator.
- **Contributing**: Feel free to fork the repository, suggest improvements, or submit pull requests!

## License
MIT License - Free to use, modify, and distribute.
