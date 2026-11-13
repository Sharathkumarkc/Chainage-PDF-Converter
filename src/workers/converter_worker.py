from PySide6.QtCore import QObject, Signal, Slot

from src.core.converter import Converter


class ConverterWorker(QObject):
    """
    Background worker running in a QThread.
    """

    finished = Signal()
    error = Signal(str)
    log = Signal(str)
    progress = Signal(int)

    def __init__(self, excel_file, output_folder):
        super().__init__()

        self.excel_file = excel_file
        self.output_folder = output_folder

    @Slot()
    def run(self):

        try:
            converter = Converter(
                log_callback=self.log.emit,
                progress_callback=self.progress.emit,
            )

            converter.convert(
                self.excel_file,
                self.output_folder,
            )

            self.progress.emit(100)

            self.finished.emit()

        except Exception as e:
            self.error.emit(str(e))