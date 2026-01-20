import csv
import datetime as dt
import os
import sys
import time
import serial

# ---- User settings ----
PORT = os.environ.get("PORT", "COM32")         # Windows example: "COM5"
# PORT = os.environ.get("PORT", "/dev/ttyACM0")  # Linux example
# PORT = os.environ.get("PORT", "/dev/tty.usbmodemXXXX")  # macOS example
BAUD = int(os.environ.get("BAUD", "115200"))
CSV_PATH = os.environ.get("CSV", "rssi_log.csv")
NOTES = os.environ.get("NOTES", "")           # e.g. set NOTES="2E lab corner"
# -----------------------

FIELDNAMES = [
    "pc_time_iso",
    "arduino_ms",
    "dur_ms",
    "mean_rssi_dbm",
    "std_rssi_db",
    "n_samples",
    "min_rssi_dbm",
    "max_rssi_dbm",
    "notes",
]

def parse_avg_line(line: str) -> dict | None:
    """
    Expected example:
    AVG,ms=123456,dur_ms=10001,mean=-72.40,std=3.10,n=86,min=-90,max=-60
    """
    if not line.startswith("AVG,"):
        return None

    parts = line.strip().split(",")
    kv = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            kv[k.strip()] = v.strip()

    # Helper to safely parse numbers
    def to_int(x, default=None):
        try:
            return int(float(x))
        except Exception:
            return default

    def to_float(x, default=None):
        try:
            return float(x)
        except Exception:
            return default

    row = {
        "pc_time_iso": dt.datetime.now().isoformat(timespec="seconds"),
        "arduino_ms": to_int(kv.get("ms")),
        "dur_ms": to_int(kv.get("dur_ms")),
        "mean_rssi_dbm": to_float(kv.get("mean")),
        "std_rssi_db": to_float(kv.get("std")),
        "n_samples": to_int(kv.get("n")),
        "min_rssi_dbm": to_int(kv.get("min")),
        "max_rssi_dbm": to_int(kv.get("max")),
        "notes": NOTES,
    }
    return row

def ensure_csv_header(path: str):
    need_header = (not os.path.exists(path)) or (os.path.getsize(path) == 0)
    if need_header:
        with open(path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
            writer.writeheader()

def append_row(path: str, row: dict):
    with open(path, "a", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writerow(row)

def main():
    ensure_csv_header(CSV_PATH)

    print(f"Opening serial port {PORT} @ {BAUD} ...")
    try:
        ser = serial.Serial(PORT, BAUD, timeout=1)
    except Exception as e:
        print(f"Failed to open {PORT}: {e}")
        print("Tip: close Arduino Serial Monitor (only one program can use the port).")
        sys.exit(1)

    # On some boards, opening serial resets the MCU; give it a moment
    time.sleep(1.5)
    ser.reset_input_buffer()

    print(f"Logging to: {CSV_PATH}")
    if NOTES:
        print(f"Notes: {NOTES}")
    print("Waiting for AVG lines... (Ctrl+C to stop)\n")

    try:
        while True:
            raw = ser.readline()
            if not raw:
                continue
            try:
                line = raw.decode("utf-8", errors="replace").strip()
            except Exception:
                continue

            if not line:
                continues

            # Print live serial output
            print(line)

            row = parse_avg_line(line)
            if row is not None:
                append_row(CSV_PATH, row)
                print("-> CSV appended\n")

    except KeyboardInterrupt:
        print("\nStopping...")
    finally:
        try:
            ser.close()
        except Exception:
            pass

if __name__ == "__main__":
    main()

# python logger.py