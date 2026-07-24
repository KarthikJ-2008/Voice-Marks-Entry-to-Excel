import queue, re
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import speech_recognition as sr

import xlwings as xw

try:
    from word2number import w2n
    HAS_W2N = True
except ImportError:
    HAS_W2N = False


# -------------------------------------------------------
# Helpers
# -------------------------------------------------------
NULL_WORDS = {
    "null", "none", "empty", "blank", "skip", "no value",
    "empty cell", "nil", "not applicable", "n a", "na",
}
_ONES = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
    "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14,
    "fifteen": 15, "sixteen": 16, "seventeen": 17, "eighteen": 18,
    "nineteen": 19,
}
_TENS = {
    "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
    "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
}

def _fallback_word_to_number(text):
    words = text.replace("-", " ").split()
    total = 0
    current = 0
    found = False
    for word in words:
        if word == "hundred":
            current = (current or 1) * 100
            found = True
        elif word in _TENS:
            current += _TENS[word]
            found = True
        elif word in _ONES:
            current += _ONES[word]
            found = True
        elif word in ("and",):
            continue
        else:
            return None
    total += current
    return total if found else None


def parse_spoken_value(raw_text):
    """
    Convert whatever the recognizer heard into either:
      - a number (int or float), or
      - "NULL" (meaning: leave the cell blank), or
      - None (meaning: could not understand, ask again)
    """
    if raw_text is None:
        return None

    text = raw_text.strip().lower()
    text = re.sub(r"[.,]", "", text)

    if not text:
        return None

    if text in NULL_WORDS or any(text == w for w in NULL_WORDS):
        return "NULL"

    for filler in ("marks", "mark", "points", "point", "out of ten",
                   "out of hundred", "out of 100"):
        text = text.replace(filler, "").strip()

    if not text:
        return None
    if text in NULL_WORDS:
        return "NULL"

    try:
        if "." in text:
            return float(text)
        return int(text)
    except ValueError:
        pass

    if HAS_W2N:
        try:
            return w2n.word_to_num(text)
        except ValueError:
            pass
    else:
        val = _fallback_word_to_number(text)
        if val is not None:
            return val

    return None


def split_cell_ref(cell_ref):
    """'B2' -> ('B', 2). Raises ValueError if invalid."""
    match = re.fullmatch(r"([A-Za-z]+)([0-9]+)", cell_ref.strip())
    if not match:
        raise ValueError("Cell reference must look like B2, C10, AA5, etc.")
    return match.group(1).upper(), int(match.group(2))


def column_letter_to_index(letter):
    """'A' -> 1, 'B' -> 2, 'AA' -> 27, etc."""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result


def column_index_to_letter(index):
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

# ---------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------
class VoiceMarksApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Voice Marks Entry to Excel")
        self.root.geometry("560x640")
        self.root.minsize(580, 640)
        self.root.resizable(False, False)

        self.app = None
        self.workbook = None
        self.worksheet = None
        self.file_path = None

        self.listening = False
        self.listen_thread = None
        self.gui_queue = queue.Queue()

        self.current_col_letter = "A"
        self.current_row = 1
        self.direction = tk.StringVar(value="down")

        self.start_col_letter = "A"
        self.header_row = 1

        self._build_widgets()
        self.root.after(150, self._poll_queue)

    # ---------------- UI ----------------

    def _build_widgets(self):
        pad = {"padx": 10, "pady": 6}
        ttk.Label(self.root, text="Thank you from Karthik",
                  foreground="gray", font=("Segoe UI", 9, "italic")).pack(
                      side="bottom", pady=(0, 8))

        frame_save = ttk.Frame(self.root)
        frame_save.pack(side="bottom", fill="x", padx=10, pady=8)
        ttk.Button(frame_save, text="Save", command=self.save_file).pack(side="left")
        ttk.Label(frame_save,
                  text="( Writes into the live Excel File - make sure to save your file ! )",
                  foreground="gray").pack(side="left", padx=8)

        frame_file = ttk.LabelFrame(self.root, text="1. Excel File")
        frame_file.pack(fill="x", **pad)

        ttk.Button(frame_file, text="Open Excel File",
                   command=self.select_file).pack(side="left", padx=8, pady=8)
        self.lbl_file = ttk.Label(frame_file, text="No file selected", width=40)
        self.lbl_file.pack(side="left", padx=8)

        frame_sheet = ttk.LabelFrame(self.root, text="2. Sheet")
        frame_sheet.pack(fill="x", **pad)

        self.sheet_combo = ttk.Combobox(frame_sheet, state="readonly", width=30)
        self.sheet_combo.pack(side="left", padx=8, pady=8)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        frame_start = ttk.LabelFrame(self.root, text="3. Starting Cell & Direction")
        frame_start.pack(fill="x", **pad)

        ttk.Label(frame_start, text="Start cell (e.g. B2):").grid(
            row=0, column=0, padx=8, pady=8, sticky="w")
        self.entry_start_cell = ttk.Entry(frame_start, width=10)
        self.entry_start_cell.insert(0, "A1")
        self.entry_start_cell.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        ttk.Radiobutton(frame_start, text="Go DOWN the column",
                        variable=self.direction, value="down").grid(
            row=1, column=0, columnspan=2, padx=8, sticky="w")
        ttk.Radiobutton(frame_start, text="Go ACROSS the row (right)",
                        variable=self.direction, value="right").grid(
            row=2, column=0, columnspan=2, padx=8, sticky="w")

        ttk.Button(frame_start, text="Set Starting Point",
                   command=self.set_start_point).grid(
            row=0, column=2, rowspan=3, padx=12)

        frame_listen = ttk.LabelFrame(self.root, text="4. Voice Entry")
        frame_listen.pack(fill="x", **pad)

        self.btn_start = ttk.Button(frame_listen, text="Start Listening",
                                     command=self.start_listening)
        self.btn_start.pack(side="left", padx=8, pady=8)
        self.btn_stop = ttk.Button(frame_listen, text="Stop Listening",
                                    command=self.stop_listening, state="disabled")
        self.btn_stop.pack(side="left", padx=8, pady=8)

        self.status_label = ttk.Label(self.root, text="Status: idle",
                                       foreground="gray")
        self.status_label.pack(fill="x", padx=14)

        self.next_cell_label = ttk.Label(
            self.root, text="Next cell to fill: (set a starting point first)",
            font=("Segoe UI", 11, "bold"))
        self.next_cell_label.pack(fill="x", padx=14, pady=(6, 2))

        self.heard_label = ttk.Label(self.root, text="Last heard: -")
        self.heard_label.pack(fill="x", padx=14)

        frame_log = ttk.LabelFrame(self.root, text="Activity Log")
        frame_log.pack(fill="both", expand=True, **pad)
        self.log_text = tk.Text(frame_log, height=10, state="disabled", wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------- File / sheet handling ----------------

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if not path:
            return

        try:
            self.workbook = xw.Book(path)  # reuses the file if it's already open in Excel
            self.app = self.workbook.app  # no duplicate copy, no file lock conflict.
            self.app.visible = True
        except Exception as exc:
            messagebox.showerror("Could not open/attach file", str(exc))
            return

        self.file_path = path
        self.lbl_file.config(text=path.split("/")[-1].split("\\")[-1])

        sheet_names = [s.name for s in self.workbook.sheets]
        self.sheet_combo["values"] = sheet_names
        self.sheet_combo.current(0)
        self.worksheet = self.workbook.sheets[0]
        self._log(f"Attached to workbook: {path}")
        if self.workbook.app.pid:
            self._log("Connected to the live Excel instance - "
                       "you can keep the file open while dictating.")

    def on_sheet_selected(self, event=None):
        if self.workbook is None:
            return
        name = self.sheet_combo.get()
        self.worksheet = self.workbook.sheets[name]
        self._log(f"Active sheet set to '{name}'")

    def set_start_point(self):
        if self.worksheet is None:
            messagebox.showwarning("No file", "Please select an Excel file first.")
            return
        try:
            col, row = split_cell_ref(self.entry_start_cell.get())
        except ValueError as exc:
            messagebox.showerror("Invalid cell", str(exc))
            return
        self.current_col_letter = col
        self.current_row = row
        self.start_col_letter = col
        self._update_next_cell_label()
        self._log(f"Starting point set to {col}{row}, "
                  f"direction = {self.direction.get()}")

    # ---------------- Listening ----------------

    def start_listening(self):
        if self.worksheet is None:
            messagebox.showwarning("No file", "Please select an Excel file first.")
            return
        if self.current_row is None:
            messagebox.showwarning("No starting point",
                                    "Please set a starting cell first.")
            return
        if self.listening:
            return

        self.listening = True
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_label.config(text="Status: listening... speak a number",
                                  foreground="green")

        self.listen_thread = threading.Thread(target=self._listen_loop, daemon=True)
        self.listen_thread.start()

    def stop_listening(self):
        self.listening = False
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.status_label.config(text="Status: idle", foreground="gray")

    def _listen_loop(self):

        recognizer = sr.Recognizer()
        try:
            mic = sr.Microphone()
        except Exception as exc:
            self.gui_queue.put(("error", f"Microphone error: {exc}"))
            return

        try:
            with mic as source:
                recognizer.adjust_for_ambient_noise(source, duration=1)
        except Exception as exc:
            self.gui_queue.put(("error", f"Could not access microphone: {exc}"))
            return

        while self.listening:
            try:
                with mic as source:
                    audio = recognizer.listen(source, timeout=5, phrase_time_limit=6)
                text = recognizer.recognize_google(audio)
                self.gui_queue.put(("heard", text))
            except sr.WaitTimeoutError:
                continue
            except sr.UnknownValueError:
                self.gui_queue.put(("unclear", ""))
            except sr.RequestError as exc:
                self.gui_queue.put(("error", f"Speech service error: {exc}"))
                break
            except Exception as exc:
                self.gui_queue.put(("error", str(exc)))
                break

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.gui_queue.get_nowait()
                if kind == "heard":
                    self._handle_heard_text(payload)
                elif kind == "unclear":
                    self._log("(did not catch that, please repeat)")
                elif kind == "error":
                    self._log(f"ERROR: {payload}")
                    self.stop_listening()
        except queue.Empty:
            pass
        self.root.after(150, self._poll_queue)

    # ---------------- Core writing logic ----------------

    def _handle_heard_text(self, raw_text):
        self.heard_label.config(text=f"Last heard: \"{raw_text}\"")
        value = parse_spoken_value(raw_text)

        cell_ref = f"{self.current_col_letter}{self.current_row}"

        if value is None:
            self._log(f"Could not understand \"{raw_text}\" - please say it again.")
            return

        if value == "NULL":
            self._log(f"{cell_ref}: left blank (heard '{raw_text}')")
        else:
            try:
                self.worksheet.range(cell_ref).value = value
                self._log(f"{cell_ref}: wrote {value} (heard '{raw_text}')")
            except Exception as exc:
                self._log(f"{cell_ref}: could not write value - {exc}")
                return

        self._advance_cell()

    def _advance_cell(self):
        if self.direction.get() == "down":
            self.current_row += 1
        else:
            col_index = column_letter_to_index(self.current_col_letter)
            next_col_letter = column_index_to_letter(col_index + 1)
            header_value = self.worksheet.range(
                f"{next_col_letter}{self.header_row}").value

            if header_value is None or str(header_value).strip() == "":
                self.current_row += 1
                self.current_col_letter = self.start_col_letter
                self._log(f"Reached end of named columns after "
                          f"'{next_col_letter}{self.header_row}' is blank - "
                          f"wrapping to row {self.current_row}, "
                          f"column {self.start_col_letter}")
            else:
                self.current_col_letter = next_col_letter
        self._update_next_cell_label()

    def _update_next_cell_label(self):
        self.next_cell_label.config(
            text=f"Next cell to fill: {self.current_col_letter}{self.current_row}")

    # ---------------- Saving ----------------

    def save_file(self):
        if not self.workbook:
            messagebox.showwarning("No file", "Nothing to save yet.")
            return
        try:
            self.workbook.save()
            self._log(f"Saved to {self.file_path}")
            messagebox.showinfo("Saved", "Workbook saved successfully.")
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))

    # ---------------- Logging ----------------

    def _log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

def main():
    root = tk.Tk()
    app = VoiceMarksApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()