## Usage

- Install with `pip install -r requirements.txt`

1. **Run the application:**
   ```bash
   python main.py
   ```
2. **Select a PDF file** using the "Browse" button.
3. **Enter circle codes** (comma-separated, e.g., `P1,P2,T3`) to filter results.
4. Click:
   - **Extract Material Codes**: Extracts and displays only materials with the specified circle codes.
   - **Extract ALL Circle Codes**: Shows all circle codes and their materials found in the PDF.
   - **Extract All Materials**: Shows all material codes, even those without a circle code.
5. **Save Results**: Click "Save Results to Excel" to export the displayed results.

## Dependencies

Your system must have:

- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [Poppler](https://poppler.freedesktop.org/)

## Linux

bash
sudo apt update
sudo apt install tesseract-ocr poppler-utils

## MacOS

brew install tesseract poppler

## How It Works

- The app first tries to extract text natively from the PDF using [PyMuPDF](https://pymupdf.readthedocs.io/).
- If no results are found, it uses [pdf2image](https://github.com/Belval/pdf2image) and [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) to perform OCR on each page.
- Material codes and circle codes are extracted using regular expressions. The code is robust to minor OCR errors and formatting inconsistencies.
- Results are displayed in the GUI and can be exported to Excel.

### Example Material Code Patterns Supported

- `39Rfi12/15cm,L=2.5m`
- `12Rfi8/10cm,L=1m`
- `10Rfi10/10/10cm,L=3.5m`

### Example Circle Code Patterns Supported

- `P1`, `T3`, `UT7a`, `B1`, `14`, `R14`, etc.

---

## Samples

- Included in the main directory a sample folder is added with supported formats.

## Troubleshooting

- **Poppler or Tesseract not found?**

  - Make sure the `poppler/bin` and `tesseract/tesseract.exe` paths exist.
  - If you use your own installation, update the `resource_path` logic in `main.py`.

- **Missing dependencies?**

  - Install with `pip install -r requirements.txt`.

- **OCR results are poor?**

  - Try increasing the DPI in the `convert_from_path` call in `main.py`.
  - Ensure your PDFs are high quality and not too noisy.

- **App crashes or freezes?**
  - Check the console for error messages.
  - Large PDFs may take time to process, especially with OCR.

---

## Customization

- **Add new material/circle code patterns:**
  Edit the regular expressions in `main.py` to support additional formats.
- **Change default paths:**
  Modify the `resource_path` function or set `TESSERACT_PATH` and `POPPLER_PATH` directly.
- **Change output format:**
  Edit the `save_to_excel` function to customize the Excel output.

---

## Developer Notes

- The code is designed to work both as a standalone script and as a PyInstaller executable.
- The `resource_path` function ensures compatibility with PyInstaller's `_MEIPASS` temp directory.
- The GUI is built with Tkinter for maximum portability.
- Logging is enabled for debugging; check the console for info and error messages.

---

## License

This project is licensed under the MIT License.
You are free to use, modify, and distribute this code for personal or commercial use — just give proper credit.
See included licenses for Poppler and Tesseract.
See the [LICENSE](./LICENSE) file for full details.

---

**Enjoy extracting your circle codes! If you have suggestions or improvements, feel free to contribute.**
```
