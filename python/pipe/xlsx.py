from openpyxl import load_workbook

class EXCEL(object):
    def __init__(self, filename):
        self._wb = load_workbook(filename, read_only=True)
        self._active_ws = self._wb.active()

    def active(self, sheet):
        pass