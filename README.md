# 🎙️ Voice Marks Entry to Excel

**Speak the marks. Watch them land in Excel.**

A lightweight desktop tool that lets teachers and evaluators dictate student marks out loud and have them typed directly into a live Excel workbook — no typing, no switching windows, no manual entry errors.

---

## ⬇️ Download

Grab the latest ready-to-run Windows app — no Python installation needed:

**[⬇️ Download VoiceMarksToExcel.exe](https://github.com/KarthikJ-2008/Voice-Marks-Entry-to-Excel/releases/download/v1.0.0/VoiceMarksToExcel.exe)**

⚠️ Since the `.exe` is unsigned, Windows SmartScreen may show a warning on first launch — click **"More info" → "Run anyway"** to proceed.

---

## ✨ Features

- 🗣️ **Voice-to-cell entry** — say a number, it lands in the next cell automatically
- 📊 **Works on a live Excel file** — attaches to an already-open workbook, no need to close it
- ↕️ **Flexible fill direction** — go down a column or across a row
- 🔢 **Understands numbers and number words** — "seventeen", "17", and "17.5" all work
- ⛔ **Skip cells on command** — say "null", "blank", or "skip" to leave a cell empty
- 🔄 **Auto-wraps to the next row** — across-row mode automatically detects the end of your headers and wraps around
- 📝 **Live activity log** — every heard phrase and every cell written is logged on screen
- 💾 **One-click Save** — saves straight back into the open workbook

---

## 🖥️ Requirements

- Windows with **Microsoft Excel** installed
- A working **microphone**
- Internet connection (speech recognition uses Google's online API)

---

## 🚀 Getting Started

1. **Download and run** `VoiceMarksToExcel.exe` (see above)
2. **Open your Excel file** using the _"Open Excel File"_ button
3. **Pick the sheet** you want to fill in
4. **Set your starting cell** (e.g. `B2`) and choose a direction:
   - ⬇️ Down the column
   - ➡️ Across the row
5. Click **Start Listening** and start speaking marks out loud
6. Click **Save** when you're done

---

## ⚠️ Notes

- Speech recognition uses Google's free web API — accuracy depends on your internet connection and microphone quality.
- Keep the Excel file open while dictating for the smoothest experience; the app attaches to the live instance rather than a hidden background copy.
- Always click **Save** — changes are written to the open workbook in memory until saved to disk.

---

## 🧩 Built With

- [SpeechRecognition](https://pypi.org/project/SpeechRecognition/) — converts spoken audio to text
- [xlwings](https://www.xlwings.org/) — connects Python to a live Excel workbook
- [word2number](https://pypi.org/project/word2number/) — converts spoken number words to digits
- [PyAudio](https://pypi.org/project/PyAudio/) — captures microphone input
- **Tkinter** — the desktop GUI (built into Python)

---

## 🙏 Credits

Built with ❤️ by **Karthik**
