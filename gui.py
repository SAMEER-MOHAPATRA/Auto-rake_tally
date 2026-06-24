from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext

from wagon_converter import SCHEMA, run_conversion


class AutoRakeTallyGUI:
    """
    GUI for the Auto Rake Tally Converter using Tkinter.
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Auto Rake Tally Converter")
        self.root.geometry("700x500")

        # Variables
        self.input_file = tk.StringVar()
        self.output_file = tk.StringVar(value="output_cleaned.csv")

        self.create_widgets()

    def create_widgets(self) -> None:
        """Creates the GUI layout and widgets."""
        # Input File Selection
        tk.Label(self.root, text="Input Excel File:").grid(
            row=0, column=0, padx=10, pady=10, sticky="e"
        )
        tk.Entry(self.root, textvariable=self.input_file, width=50).grid(
            row=0, column=1, padx=10, pady=10
        )
        tk.Button(self.root, text="Browse...", command=self.browse_input).grid(
            row=0, column=2, padx=10, pady=10
        )

        # Output File Selection
        tk.Label(self.root, text="Output CSV File:").grid(
            row=1, column=0, padx=10, pady=10, sticky="e"
        )
        tk.Entry(self.root, textvariable=self.output_file, width=50).grid(
            row=1, column=1, padx=10, pady=10
        )
        tk.Button(self.root, text="Browse...", command=self.browse_output).grid(
            row=1, column=2, padx=10, pady=10
        )

        # Run Button
        self.run_btn = tk.Button(
            self.root,
            text="Start Conversion",
            command=self.start_conversion,
            bg="green",
            fg="white",
            font=("Arial", 10, "bold"),
            height=2,
        )
        self.run_btn.grid(row=2, column=0, columnspan=3, pady=20)

        # Log Window
        tk.Label(self.root, text="Process Logs:").grid(
            row=3, column=0, padx=10, sticky="w"
        )
        self.log_area = scrolledtext.ScrolledText(
            self.root, width=80, height=15, state="disabled"
        )
        self.log_area.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

    def browse_input(self) -> None:
        """Opens a file dialog to select the input Excel file."""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.input_file.set(filename)

    def browse_output(self) -> None:
        """Opens a file dialog to select the output CSV file."""
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
        )
        if filename:
            self.output_file.set(filename)

    def log(self, message: str) -> None:
        """Appends a message to the scrolled text log area."""
        self.log_area.config(state="normal")
        self.log_area.insert(tk.END, str(message) + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state="disabled")

    def start_conversion(self) -> None:
        """Initiates the conversion process in a background thread."""
        input_path = self.input_file.get()
        output_path = self.output_file.get()

        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid input Excel file.")
            return

        self.run_btn.config(state="disabled", text="Processing...")
        self.log_area.config(state="normal")
        self.log_area.delete("1.0", tk.END)
        self.log_area.config(state="disabled")

        # Run in a separate thread to keep UI responsive
        thread = threading.Thread(
            target=self.run_process, args=(input_path, output_path)
        )
        thread.daemon = True
        thread.start()

    def run_process(self, input_path: str, output_path: str) -> None:
        """Executes the conversion logic and updates the UI with results."""
        try:
            stats = run_conversion(input_path, SCHEMA, output_path, logger=self.log)

            if "error" in stats:
                self.log(f"CRITICAL ERROR: {stats['error']}")
                messagebox.showerror(
                    "Process Failed", f"An error occurred: {stats['error']}"
                )
            else:
                self.log("\n" + "=" * 30)
                self.log("CONVERSION SUMMARY")
                self.log("=" * 30)
                self.log(f"Primary Sheet: {stats['primary_sheet']}")
                self.log(f"Total Rows:    {stats['total']}")
                self.log(f"Matched:       {stats['matched']}")
                self.log(f"Fuzzy/Prefix:  {stats['fuzzy']}")
                self.log(f"Failed:        {stats['failed']}")
                self.log(f"Output:        {stats['output_path']}")
                self.log("=" * 30)

                messagebox.showinfo(
                    "Success",
                    f"Conversion completed successfully!\nMatched: {stats['matched']}/{stats['total']}",
                )
        except Exception as e:
            self.log(f"Unexpected Error: {e}")
            messagebox.showerror("Error", f"An unexpected error occurred: {e}")
        finally:
            self.root.after(
                0, lambda: self.run_btn.config(state="normal", text="Start Conversion")
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = AutoRakeTallyGUI(root)
    root.mainloop()
