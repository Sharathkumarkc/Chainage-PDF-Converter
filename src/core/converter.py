from pathlib import Path
import traceback


from src.core.excel_to_pdf_chainagewise_v2 import (
    split_workbook_by_chainage,
)


class Converter:
    """
    Main converter class.
    Handles:
        - Logging
        - Progress updates
        - Calling the LibreOffice conversion engine
    """

    def __init__(self, log_callback=None, progress_callback=None):
        self.log_callback = log_callback
        self.progress_callback = progress_callback

    # --------------------------------------------------

    def log(self, message):
        print(message)

        if self.log_callback:
            self.log_callback(message)

    # --------------------------------------------------

    def progress(self, value):
        if self.progress_callback:
            self.progress_callback(value)

    # --------------------------------------------------

    def convert(self, excel_file, output_folder):

        excel_file = Path(excel_file)
        output_folder = Path(output_folder)

        if not excel_file.exists():
            raise FileNotFoundError(
                f"Excel file not found:\n{excel_file}"
            )

        output_folder.mkdir(parents=True, exist_ok=True)

        try:

            self.progress(0)

            self.log("========================================")
            self.log("Chainage PDF Converter")
            self.log("========================================")
            self.log(f"Input : {excel_file}")
            self.log(f"Output: {output_folder}")
            self.log("")

            # ------------------------------------------
            # Check LibreOffice
            # ------------------------------------------

            self.log("Checking LibreOffice...")
            self.progress(10)

            # LibreOffice will be checked automatically when conversion starts.
            self.log("Starting conversion...")

            

            # ------------------------------------------
            # Start Conversion
            # ------------------------------------------

            self.log("")
            self.log("Starting conversion...")
            self.progress(20)

            results = split_workbook_by_chainage(
                str(excel_file),
                str(output_folder),
            )

            self.progress(100)

            self.log("")
            self.log("========================================")
            self.log("Conversion Completed Successfully")
            self.log("========================================")

            if results:
                self.log(f"Generated {len(results)} PDF file(s).")

        except Exception:

            self.progress(0)

            error = traceback.format_exc()

            self.log(error)

            raise Exception(error)