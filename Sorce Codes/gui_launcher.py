import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import socket
import threading
import os
import sys
import pyperclip

from saleae.grpc import saleae_pb2
from decode_spi_core import decode_spi, SPI_SETTINGS
from decode_i2c_core import decode_i2c, I2C_SETTINGS


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


ICON_PATH = resource_path("chip_icon.ico")


class ProtocolDecoderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Saleae Protocol Decoder v1.1.0")

        try:
            self.root.iconbitmap(ICON_PATH)
        except Exception as e:
            print(f"Warning: Could not load icon: {e}")

        self.sal_file_path = ""
        self.protocol_var = tk.StringVar(value="SPI")
        self.settings_entries = {}
        self.export_excel_var = tk.BooleanVar(value=False)

        main = ttk.Frame(root, padding=10)
        main.pack(fill="both", expand=True)

        # --- Protocol Selection ---
        protocol_frame = ttk.Frame(main)
        protocol_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(protocol_frame, text="Select Protocol:").pack(side="left")
        ttk.OptionMenu(
            protocol_frame,
            self.protocol_var,
            "SPI",
            "SPI",
            "I2C",
            command=self.on_protocol_change,
        ).pack(side="left", padx=(5, 0))

        # --- File Selector ---
        file_frame = ttk.Frame(main)
        file_frame.pack(fill="x", pady=5)
        ttk.Button(file_frame, text="Browse .sal File", command=self.select_file).pack(
            side="left"
        )
        self.file_label = ttk.Label(file_frame, text="No file selected", anchor="w")
        self.file_label.pack(side="left", fill="x", expand=True, padx=10)

        # --- Settings Frame ---
        self.settings_frame = ttk.LabelFrame(main, text="Analyzer Settings", padding=10)
        self.settings_frame.pack(fill="x", pady=5)

        # --- Optional Excel Export ---
        excel_frame = ttk.Frame(main)
        excel_frame.pack(fill="x", pady=5)
        ttk.Checkbutton(
            excel_frame,
            text="Export also to Excel (.xlsx)",
            variable=self.export_excel_var,
        ).pack(side="left")

        # --- Start Button ---
        ttk.Button(
            main, text="Start Decode", command=self.start_decode_threaded
        ).pack(pady=5)

        # --- Progress with label and percentage ---
        progress_frame = ttk.Frame(main)
        progress_frame.pack(fill="x", pady=(0, 5), padx=5)
        ttk.Label(progress_frame, text="Progress:").pack(side="left")
        self.progress = ttk.Progressbar(progress_frame, length=300)
        self.progress.pack(side="left", fill="x", expand=True, padx=5)
        self.percent_label = ttk.Label(progress_frame, text="0%")
        self.percent_label.pack(side="right")

        # --- Log Output ---
        log_frame = ttk.LabelFrame(main, text="Log Output", padding=5)
        log_frame.pack(fill="both", expand=True)
        self.log_output = tk.Text(log_frame, height=12, wrap="word")
        self.log_output.pack(fill="both", expand=True)

        # --- Copy Button ---
        self.copy_button = ttk.Button(
            main,
            text="Copy File Path to Clipboard",
            command=self.copy_output_path,
            state="disabled",
        )
        self.copy_button.pack(pady=(5, 0))

        self.output_path_txt = ""
        self.output_path_xlsx = ""
        self.status = tk.StringVar()
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x")

        # Initialize with SPI settings
        self.update_settings_fields("SPI")

    def on_protocol_change(self, protocol):
        """Clear progress and log when switching protocol"""
        self.update_settings_fields(protocol)
        self.clear_progress()
        self.clear_log()

    def update_settings_fields(self, protocol):
        for widget in self.settings_frame.winfo_children():
            widget.destroy()
        self.settings_entries.clear()

        settings = SPI_SETTINGS if protocol == "SPI" else I2C_SETTINGS

        for key, value in settings.items():
            val_type = value.WhichOneof("value")
            val_data = getattr(value, val_type)

            row = ttk.Frame(self.settings_frame)
            row.pack(fill="x", pady=2)
            ttk.Label(row, text=key, width=25).pack(side="left")

            # Special handling for dropdown options in SPI
            if protocol == "SPI" and key in [
                "Significant Bit",
                "Clock State",
                "Clock Phase",
                "Enable Line",
            ]:
                if key == "Significant Bit":
                    options = [
                        "Most Significant Bit First (Standard)",
                        "Least Significant Bit First",
                    ]
                elif key == "Clock State":
                    options = [
                        "Clock is Low when inactive (CPOL = 0)",
                        "Clock is High when inactive (CPOL = 1)",
                    ]
                elif key == "Clock Phase":
                    options = [
                        "Data is Valid on Clock Leading Edge (CPHA = 0)",
                        "Data is Valid on Clock Trailing Edge (CPHA = 1)",
                    ]
                elif key == "Enable Line":
                    options = [
                        "Enable line is Active Low (Standard)",
                        "Enable line is Active High",
                    ]
                var = tk.StringVar(value=val_data)
                dropdown = ttk.OptionMenu(row, var, val_data, *options)
                dropdown.pack(side="right", fill="x", expand=True)
                self.settings_entries[key] = (val_type, var)
            else:
                entry = ttk.Entry(row)
                entry.insert(0, str(val_data))
                entry.pack(side="right", fill="x", expand=True)
                self.settings_entries[key] = (val_type, entry)

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Saleae Files", "*.sal")])
        if path:
            self.sal_file_path = path
            self.file_label.config(text=os.path.basename(path))

    def build_settings_dict(self):
        settings = {}
        for key, (val_type, widget) in self.settings_entries.items():
            if isinstance(widget, tk.StringVar):
                val = widget.get()
                settings[key] = saleae_pb2.AnalyzerSettingValue(string_value=val)
            else:
                val = widget.get()
                if val_type == "int64_value":
                    settings[key] = saleae_pb2.AnalyzerSettingValue(
                        int64_value=int(val)
                    )
                elif val_type == "string_value":
                    settings[key] = saleae_pb2.AnalyzerSettingValue(string_value=val)
        return settings

    def is_saleae_running(self):
        try:
            with socket.create_connection(("localhost", 10430), timeout=2):
                return True
        except OSError:
            return False

    def log(self, msg):
        self.log_output.insert(tk.END, msg + "\n")
        self.log_output.see(tk.END)

    def clear_log(self):
        self.log_output.delete("1.0", tk.END)

    def clear_progress(self):
        self.progress["value"] = 0
        self.percent_label.config(text="0%")

    def update_progress(self, percent):
        self.progress["value"] = percent
        self.percent_label.config(text=f"{percent}%")
        self.root.update_idletasks()

    def copy_output_path(self):
        if self.output_path_txt:
            pyperclip.copy(self.output_path_txt)
            messagebox.showinfo("Copied", "Output file path copied to clipboard.")

    def start_decode_threaded(self):
        thread = threading.Thread(target=self.start_decode)
        thread.start()

    def start_decode(self):
        self.clear_log()
        self.clear_progress()
        self.status.set("Starting decoding...")
        self.copy_button.config(state="disabled")
        self.output_path_txt = ""
        self.output_path_xlsx = ""

        if not self.sal_file_path:
            messagebox.showerror("Error", "Please select a .sal file.")
            self.status.set("Error: No file selected")
            return

        if not self.is_saleae_running():
            messagebox.showerror(
                "Error", "Saleae Logic 2 is not running on port 10430."
            )
            self.status.set("Error: Saleae not connected")
            return

        protocol = self.protocol_var.get()
        settings = self.build_settings_dict()

        try:
            if protocol == "SPI":
                self.output_path_txt, self.output_path_xlsx = decode_spi(
                    self.sal_file_path,
                    settings_override=settings,
                    progress_callback=self.update_progress,
                    export_excel=self.export_excel_var.get(),
                )
            elif protocol == "I2C":
                self.output_path_txt, self.output_path_xlsx = decode_i2c(
                    self.sal_file_path,
                    settings_override=settings,
                    progress_callback=self.update_progress,
                    export_excel=self.export_excel_var.get(),
                )
            else:
                raise ValueError("Unsupported protocol")

            self.log(f"Decoded using {protocol}")
            self.log(f"Text Output saved to: {self.output_path_txt}")
            if self.output_path_xlsx:
                self.log(f"Excel Output saved to: {self.output_path_xlsx}")

            self.copy_button.config(state="normal")
            self.status.set("Decoding completed")
        except Exception as e:
            self.log(f"Error: {e}")
            self.status.set("Error during decoding")


if __name__ == "__main__":
    root = tk.Tk()
    app = ProtocolDecoderGUI(root)
    root.mainloop()
