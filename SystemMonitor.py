#!/usr/bin/env python3
"""
Advanced Terminal Activity Tracker for Windows

Features:
- Display system boot time, current time, and uptime
- Show programs active for more than 1 minute
- Display CPU and Memory usage for each program
- Color-coded table for CPU/Memory usage
- Export to Excel file when user types 'ED' + Enter
- Optional: user can specify Excel file path; default is execution directory
- Timestamped Excel file names by default
- Aggregate multiple processes with the same name

Requirements:
    pip install psutil tabulate colorama openpyxl
"""

import psutil
from datetime import datetime, timedelta
from tabulate import tabulate
from colorama import init, Fore, Style
import openpyxl
import os

# Initialize colorama
init(autoreset=True)

# Minimum active time filter in seconds
MIN_ACTIVE_TIME = 60  # 1 minute

# 1. System boot time and uptime
boot_time = datetime.fromtimestamp(psutil.boot_time())
now = datetime.now()
uptime = now - boot_time
print(f"System booted at: {boot_time.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"Current time: {now.strftime('%Y-%m-%d %H:%M:%S')}")
print(f"Uptime: {str(uptime).split('.')[0]}")  # Remove milliseconds

# Always from system boot
report_start_time = boot_time

# 2. Collect processes and aggregate by name
process_groups = {}
for p in psutil.process_iter(['name', 'cpu_percent', 'memory_info', 'create_time']):
    try:
        name = p.info['name']
        if name.lower() == 'system idle process':
            continue  # Skip System Idle Process
        process_start = datetime.fromtimestamp(p.info['create_time'])
        if process_start < report_start_time:
            continue  # Should not happen, but keep
        elapsed_seconds = (now - process_start).total_seconds()
        if elapsed_seconds < MIN_ACTIVE_TIME:
            continue
        cpu_percent = p.info['cpu_percent'] or 0
        memory_mb = p.info['memory_info'].rss / 1024 / 1024  # Convert bytes to MB
        if name not in process_groups:
            process_groups[name] = {
                'min_start': process_start,
                'cpu': 0,
                'mem': 0,
                'count': 0
            }
        group = process_groups[name]
        group['min_start'] = min(group['min_start'], process_start)
        group['cpu'] += cpu_percent
        group['mem'] += memory_mb
        group['count'] += 1
    except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
        continue

# Prepare processes list
processes = []
for name, group in process_groups.items():
    active_time = now - group['min_start']
    processes.append({
        'name': name,
        'active': str(active_time).split('.')[0],
        'cpu': group['cpu'],
        'mem': group['mem']
    })

# 3. Sort by Active Time descending
for p in processes:
    p['active_seconds'] = (now - process_groups[p['name']]['min_start']).total_seconds()
processes.sort(key=lambda x: x['active_seconds'], reverse=True)

# 4. Display table with colors
def colorize_cpu(cpu):
    if cpu > 20:
        return Fore.RED + f"{cpu:.1f}%" + Style.RESET_ALL
    elif cpu >= 5:
        return Fore.YELLOW + f"{cpu:.1f}%" + Style.RESET_ALL
    else:
        return Fore.GREEN + f"{cpu:.1f}%" + Style.RESET_ALL

def colorize_memory(mem):
    if mem > 500:
        return Fore.RED + f"{mem:.1f} MB" + Style.RESET_ALL
    elif mem >= 100:
        return Fore.YELLOW + f"{mem:.1f} MB" + Style.RESET_ALL
    else:
        return Fore.GREEN + f"{mem:.1f} MB" + Style.RESET_ALL

if processes:
    table = [(p['name'], p['active'], colorize_cpu(p['cpu']), colorize_memory(p['mem'])) for p in processes]
    print('\nActive Programs (more than 1 minute):')
    print(tabulate(table, headers=['Program Name', 'Active Time', 'CPU %', 'Memory MB']))
else:
    print('\nNo programs active for more than 1 minute.')

# 5. Wait for user input to export
user_input = input('\nType ED and press Enter to export report to Excel (or any other key to exit): ').strip().upper()
if user_input == 'ED':
    timestamp = now.strftime('%Y-%m-%d_%H-%M-%S')
    default_path = os.path.join(os.getcwd(), f'activity_report_{timestamp}.xlsx')
    excel_path_input = input(f'Enter path to save Excel report (press Enter for default: {default_path}): ').strip()
    excel_file = excel_path_input if excel_path_input else default_path

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Activity Report'
    ws.append(['Program Name', 'Active Time', 'CPU %', 'Memory MB'])
    for p in processes:
        ws.append([p['name'], p['active'], p['cpu'], p['mem']])

    try:
        wb.save(excel_file)
        print(f'Report saved to {excel_file}')
    except Exception as e:
        print(f'Error saving Excel file: {e}')