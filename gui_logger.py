import csv
import datetime as dt
import queue
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import serial
from serial.tools import list_ports


BAUD = 115200


def parse_avg_line(line: str):
    """
    Expected line format from Arduino:
    AVG,ms=123456,dur_ms=10001,mean=-72.40,std=3.10,n=86,min=-90,max=-60
    Returns dict or None.
    """
    line = line.strip()
    if not line.startswith("AVG,"):
        return None

    parts = line.split(",")[1:]
    kv = {}
    for p in parts:
        if "=" in p:
            k, v = p.split("=", 1)
            kv[k.strip()] = v.strip()

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

    return {
        "arduino_ms": to_int(kv.get("ms")),
        "dur_ms": to_int(kv.get("dur_ms")),
        "mean": to_float(kv.get("mean")),
        "std": to_float(kv.get("std")),
        "n": to_int(kv.get("n")),
        "min": to_int(kv.get("min")),
        "max": to_int(kv.get("max")),
    }


def available_ports():
    ports = []
    for p in list_ports.comports():
        # p.device like "COM32"
        # p.description like "USB Serial Device"
        ports.append((p.device, p.description))
    return ports


class SerialReader(threading.Thread):
    def __init__(self, port, baud, out_queue, stop_event):
        super().__init__(daemon=True)
        self.port = port
        self.baud = baud
        self.out_queue = out_queue
        self.stop_event = stop_event
        self.ser = None

    def run(self):
        try:
            self.ser = serial.Serial(self.port, self.baud, timeout=1)
            # Many boards reset on open; give it a moment and clear buffer
            time.sleep(1.5)
            try:
                self.ser.reset_input_buffer()
            except Exception:
                pass
            self.out_queue.put(("status", f"Connected to {self.port} @ {self.baud}"))
        except Exception as e:
            self.out_queue.put(("error", f"Failed to open {self.port}: {e}"))
            return

        try:
            while not self.stop_event.is_set():
                raw = self.ser.readline()
                if not raw:
                    continue
                line = raw.decode("utf-8", errors="replace").strip()
                if not line:
                    continue
                self.out_queue.put(("line", line))
        finally:
            try:
                if self.ser:
                    self.ser.close()
            except Exception:
                pass
            self.out_queue.put(("status", "Disconnected"))


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("RSSI Logger (RFM69)")

        self.q = queue.Queue()
        self.stop_event = threading.Event()
        self.reader = None

        self.rows = []  # in-memory list of dicts

        self._build_ui()
        self._refresh_ports()
        self.after(100, self._poll_queue)

    def _build_ui(self):
        # Top controls frame
        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(7, weight=1)

        ttk.Label(top, text="Port:").grid(row=0, column=0, sticky="w")
        self.port_var = tk.StringVar(value="COM32")
        self.port_combo = ttk.Combobox(top, textvariable=self.port_var, width=12, state="readonly")
        self.port_combo.grid(row=0, column=1, padx=(5, 10), sticky="w")

        ttk.Button(top, text="Refresh", command=self._refresh_ports).grid(row=0, column=2, padx=(0, 10))

        self.connect_btn = ttk.Button(top, text="Connect", command=self._connect)
        self.connect_btn.grid(row=0, column=3, padx=(0, 10))

        self.disconnect_btn = ttk.Button(top, text="Disconnect", command=self._disconnect, state="disabled")
        self.disconnect_btn.grid(row=0, column=4, padx=(0, 10))

        ttk.Label(top, text="Location:").grid(row=0, column=5, sticky="e")
        self.location_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self.location_var, width=22).grid(row=0, column=6, padx=(5, 10), sticky="w")

        ttk.Label(top, text="Notes:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        self.notes_var = tk.StringVar(value="")
        ttk.Entry(top, textvariable=self.notes_var, width=60).grid(row=1, column=1, columnspan=6, pady=(8, 0), sticky="ew")

        # Status + last line
        status_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        status_frame.grid(row=1, column=0, sticky="ew")
        status_frame.columnconfigure(1, weight=1)

        ttk.Label(status_frame, text="Status:").grid(row=0, column=0, sticky="w")
        self.status_var = tk.StringVar(value="Not connected")
        ttk.Label(status_frame, textvariable=self.status_var).grid(row=0, column=1, sticky="w")

        ttk.Label(status_frame, text="Last serial:").grid(row=1, column=0, sticky="w", pady=(5, 0))
        self.lastline_var = tk.StringVar(value="")
        ttk.Label(status_frame, textvariable=self.lastline_var).grid(row=1, column=1, sticky="w", pady=(5, 0))

        # Table
        table_frame = ttk.Frame(self, padding=10)
        table_frame.grid(row=2, column=0, sticky="nsew")
        self.rowconfigure(2, weight=1)
        self.columnconfigure(0, weight=1)

        columns = (
            "pc_time",
            "location",
            "notes",
            "arduino_ms",
            "dur_ms",
            "mean",
            "std",
            "n",
            "min",
            "max",
        )

        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=14)
        for c in columns:
            self.tree.heading(c, text=c)
            # sensible widths
            w = 110
            if c in ("pc_time",):
                w = 150
            if c in ("location",):
                w = 140
            if c in ("notes",):
                w = 220
            if c in ("mean", "std"):
                w = 80
            if c in ("n", "min", "max"):
                w = 70
            self.tree.column(c, width=w, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.grid(row=0, column=1, sticky="ns")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Bottom buttons
        bottom = ttk.Frame(self, padding=10)
        bottom.grid(row=3, column=0, sticky="ew")
        bottom.columnconfigure(5, weight=1)

        ttk.Button(bottom, text="Delete selected", command=self._delete_selected).grid(row=0, column=0, padx=(0, 10))
        ttk.Button(bottom, text="Clear all", command=self._clear_all).grid(row=0, column=1, padx=(0, 10))
        ttk.Button(bottom, text="Save CSV", command=self._save_csv).grid(row=0, column=2, padx=(0, 10))

        ttk.Label(bottom, text="Tip: close Arduino Serial Monitor while using this.").grid(row=0, column=3, sticky="w")

    def _refresh_ports(self):
        ports = available_ports()
        if not ports:
            self.port_combo["values"] = []
            return

        values = [p[0] for p in ports]  # "COM32"
        self.port_combo["values"] = values

        # keep current if still available, else select first
        cur = self.port_var.get()
        if cur not in values:
            self.port_var.set(values[0])

    def _connect(self):
        if self.reader and self.reader.is_alive():
            return

        port = self.port_var.get().strip()
        if not port:
            messagebox.showerror("No port", "Select a COM port first.")
            return

        self.stop_event.clear()
        self.reader = SerialReader(port, BAUD, self.q, self.stop_event)
        self.reader.start()

        self.connect_btn.config(state="disabled")
        self.disconnect_btn.config(state="normal")
        self.status_var.set(f"Connecting to {port}...")

    def _disconnect(self):
        self.stop_event.set()
        self.connect_btn.config(state="normal")
        self.disconnect_btn.config(state="disabled")

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "status":
                    self.status_var.set(payload)
                    if payload == "Disconnected":
                        self.connect_btn.config(state="normal")
                        self.disconnect_btn.config(state="disabled")
                elif kind == "error":
                    self.status_var.set(payload)
                    messagebox.showerror("Serial error", payload)
                    self.connect_btn.config(state="normal")
                    self.disconnect_btn.config(state="disabled")
                elif kind == "line":
                    self.lastline_var.set(payload)
                    parsed = parse_avg_line(payload)
                    if parsed:
                        self._add_row(parsed)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    def _add_row(self, parsed):
        row = {
            "pc_time": dt.datetime.now().isoformat(timespec="seconds"),
            "location": self.location_var.get().strip(),
            "notes": self.notes_var.get().strip(),
            "arduino_ms": parsed.get("arduino_ms"),
            "dur_ms": parsed.get("dur_ms"),
            "mean": parsed.get("mean"),
            "std": parsed.get("std"),
            "n": parsed.get("n"),
            "min": parsed.get("min"),
            "max": parsed.get("max"),
        }
        self.rows.append(row)

        # format floats nicely
        mean_str = "" if row["mean"] is None else f"{row['mean']:.2f}"
        std_str = "" if row["std"] is None else f"{row['std']:.2f}"

        values = (
            row["pc_time"],
            row["location"],
            row["notes"],
            row["arduino_ms"] if row["arduino_ms"] is not None else "",
            row["dur_ms"] if row["dur_ms"] is not None else "",
            mean_str,
            std_str,
            row["n"] if row["n"] is not None else "",
            row["min"] if row["min"] is not None else "",
            row["max"] if row["max"] is not None else "",
        )
        # iid is index in rows list
        iid = str(len(self.rows) - 1)
        self.tree.insert("", "end", iid=iid, values=values)

    def _delete_selected(self):
        sel = self.tree.selection()
        if not sel:
            return

        # Delete from tree first
        for iid in sel:
            self.tree.delete(iid)

        # Rebuild rows from what's left in tree (simpler than index surgery)
        self._rebuild_rows_from_tree()

    def _clear_all(self):
        if not self.rows:
            return
        if not messagebox.askyesno("Clear all", "Delete all logged rows?"):
            return
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.rows.clear()

    def _rebuild_rows_from_tree(self):
        new_rows = []
        items = self.tree.get_children()

        # Extract values back into dict form
        for item in items:
            vals = self.tree.item(item, "values")
            # vals order matches columns
            row = {
                "pc_time": vals[0],
                "location": vals[1],
                "notes": vals[2],
                "arduino_ms": vals[3],
                "dur_ms": vals[4],
                "mean": vals[5],
                "std": vals[6],
                "n": vals[7],
                "min": vals[8],
                "max": vals[9],
            }
            # convert numeric fields where possible
            def as_int(x):
                try:
                    return int(x)
                except Exception:
                    return None

            def as_float(x):
                try:
                    return float(x)
                except Exception:
                    return None

            row["arduino_ms"] = as_int(row["arduino_ms"])
            row["dur_ms"] = as_int(row["dur_ms"])
            row["n"] = as_int(row["n"])
            row["min"] = as_int(row["min"])
            row["max"] = as_int(row["max"])
            row["mean"] = as_float(row["mean"])
            row["std"] = as_float(row["std"])

            new_rows.append(row)

        self.rows = new_rows

        # Reassign iids to be 0..N-1
        for item in items:
            self.tree.delete(item)
        for i, row in enumerate(self.rows):
            mean_str = "" if row["mean"] is None else f"{row['mean']:.2f}"
            std_str = "" if row["std"] is None else f"{row['std']:.2f}"
            values = (
                row["pc_time"],
                row["location"],
                row["notes"],
                row["arduino_ms"] if row["arduino_ms"] is not None else "",
                row["dur_ms"] if row["dur_ms"] is not None else "",
                mean_str,
                std_str,
                row["n"] if row["n"] is not None else "",
                row["min"] if row["min"] is not None else "",
                row["max"] if row["max"] is not None else "",
            )
            self.tree.insert("", "end", iid=str(i), values=values)

    def _save_csv(self):
        if not self.rows:
            messagebox.showinfo("Nothing to save", "No rows logged yet.")
            return

        default_name = f"rssi_log_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return

        fieldnames = [
            "pc_time_iso",
            "location",
            "notes",
            "arduino_ms",
            "dur_ms",
            "mean_rssi_dbm",
            "std_rssi_db",
            "n_samples",
            "min_rssi_dbm",
            "max_rssi_dbm",
        ]

        try:
            with open(path, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=fieldnames)
                w.writeheader()
                for r in self.rows:
                    w.writerow({
                        "pc_time_iso": r["pc_time"],
                        "location": r["location"],
                        "notes": r["notes"],
                        "arduino_ms": r["arduino_ms"],
                        "dur_ms": r["dur_ms"],
                        "mean_rssi_dbm": r["mean"],
                        "std_rssi_db": r["std"],
                        "n_samples": r["n"],
                        "min_rssi_dbm": r["min"],
                        "max_rssi_dbm": r["max"],
                    })
            messagebox.showinfo("Saved", f"Saved {len(self.rows)} rows to:\n{path}")
        except Exception as e:
            messagebox.showerror("Save failed", f"Could not save CSV:\n{e}")

    def on_close(self):
        self.stop_event.set()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()

