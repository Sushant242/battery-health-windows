import subprocess
import re
import os
import sys

def get_cycle_count():
    try:
        result = subprocess.check_output(
            ["powershell",
             "-Command",
             "Get-WmiObject Win32_Battery | Select -ExpandProperty CycleCount"],
            stderr=subprocess.DEVNULL,
            text=True
        ).strip()
        return result if result else "Not supported"
    except:
        return "Not supported"

try:
    report_path = os.path.join(os.path.expanduser("~"), "battery_report.html")

    subprocess.run(
        ["powercfg", "/batteryreport", f"/output={report_path}"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        shell=True
    )

    with open(report_path, "r", encoding="utf-8") as f:
        data = f.read()

    design = re.search(r"Design Capacity</span>\s*<span.*?>([\d,]+)", data)
    full = re.search(r"Full Charge Capacity</span>\s*<span.*?>([\d,]+)", data)

    cycle_count = get_cycle_count()

    print("\n BATTERY HEALTH REPORT")
    print("=" * 32)

    if design and full:
        design = int(design.group(1).replace(",", ""))
        full = int(full.group(1).replace(",", ""))
        health = round((full / design) * 100, 2)

        print(f"Battery Health       : {health}%")
        print(f"Design Capacity      : {design} mWh")
        print(f"Full Charge Capacity : {full} mWh")
    else:
        print("Battery capacity data not found")

    print(f"Battery Cycle Count  : {cycle_count}")

except Exception as e:
    print("Error:", e)

input("\nPress Enter to exit...")
