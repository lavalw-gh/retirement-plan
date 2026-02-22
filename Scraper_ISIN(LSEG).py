import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import re
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException


# ---------------------------------------------------------
# Selenium-based ISIN extraction
# ---------------------------------------------------------
ISIN_REGEX = re.compile(r"\b[A-Z]{2}[A-Z0-9]{9}[0-9]\b")


def create_driver():
    options = Options()
    # comment this out if you want to see the browser
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--window-size=1200,800")

    # If chromedriver is not on PATH, set executable_path here
    try:
        driver = webdriver.Chrome(options=options)
    except WebDriverException as e:
        raise RuntimeError(f"Could not start ChromeDriver: {e}")
    return driver


def get_isin_for_ticker(ticker, driver):
    url = f"https://www.londonstockexchange.com/stock/{ticker}/details"

    try:
        driver.get(url)
    except WebDriverException:
        return "Navigation Error"

    # Give the page some time to render JS
    time.sleep(3)

    html = driver.page_source

    # First, try to find an ISIN-looking string anywhere in the HTML
    match = ISIN_REGEX.search(html)
    if match:
        return match.group(0)

    return "ISIN Not Found"


# ---------------------------------------------------------
# GUI Application
# ---------------------------------------------------------
class ISINScraperGUI:
    def __init__(self, root):
        self.root = root
        root.title("LSE Ticker → ISIN Lookup Tool")
        root.geometry("750x550")

        # Create Selenium driver once and reuse
        try:
            self.driver = create_driver()
        except RuntimeError as e:
            messagebox.showerror("Driver Error", str(e))
            self.driver = None

        # Input Label
        tk.Label(root, text="Enter LSE Tickers (comma or newline separated):").pack(
            pady=5)

        # Text Input Box
        self.text_input = tk.Text(root, height=6, width=70)
        self.text_input.pack()

        # Buttons Frame
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)

        tk.Button(btn_frame, text="Lookup ISINs", command=self.lookup,
                  width=15).grid(row=0, column=0, padx=10)
        tk.Button(btn_frame, text="Export CSV", command=self.export_csv,
                  width=15).grid(row=0, column=1, padx=10)

        # Results Table
        self.tree = ttk.Treeview(root, columns=(
            "Ticker", "ISIN"), show="headings", height=15)
        self.tree.heading("Ticker", text="Ticker")
        self.tree.heading("ISIN", text="ISIN")
        self.tree.column("Ticker", width=120)
        self.tree.column("ISIN", width=250)
        self.tree.pack(pady=10)

        # Status Label
        self.status = tk.Label(root, text="", fg="blue")
        self.status.pack()

        # Clean up driver on close
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    # ---------------------------------------------------------
    # Lookup Function
    # ---------------------------------------------------------
    def lookup(self):
        if self.driver is None:
            messagebox.showerror(
                "Driver Error", "WebDriver is not available. Check your ChromeDriver setup.")
            return

        self.tree.delete(*self.tree.get_children())
        self.status.config(text="Working...")

        raw_text = self.text_input.get("1.0", tk.END).strip()
        if not raw_text:
            messagebox.showwarning(
                "Input Error", "Please enter at least one ticker.")
            return

        tickers = [t.strip().upper()
                   for t in raw_text.replace(",", "\n").split("\n") if t.strip()]

        for ticker in tickers:
            self.status.config(text=f"Fetching ISIN for {ticker}...")
            self.root.update_idletasks()

            isin = get_isin_for_ticker(ticker, self.driver)
            self.tree.insert("", tk.END, values=(ticker, isin))

        self.status.config(text="Lookup complete.")

    # ---------------------------------------------------------
    # Export CSV
    # ---------------------------------------------------------
    def export_csv(self):
        rows = self.tree.get_children()
        if not rows:
            messagebox.showwarning("No Data", "No results to export.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")]
        )

        if not file_path:
            return

        with open(file_path, "w", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["Ticker", "ISIN"])
            for row in rows:
                writer.writerow(self.tree.item(row)["values"])

        messagebox.showinfo("Export Complete", "CSV file saved successfully.")

    # ---------------------------------------------------------
    # Cleanup
    # ---------------------------------------------------------
    def on_close(self):
        if self.driver is not None:
            try:
                self.driver.quit()
            except Exception:
                pass
        self.root.destroy()


# ---------------------------------------------------------
# Run the App
# ---------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = ISINScraperGUI(root)
    root.mainloop()
