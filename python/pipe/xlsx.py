from typing import List
from openpyxl import load_workbook

class EXCEL(object):
    def __init__(self, filename: str, read_only: bool = True) -> None:
        """Initiate EXCEL class
        
        Keyword arguments:
        filename -- The path of excel file.
        read_only -- Read/Write mode. Default read mode to protect raw data.
        Return: None
        """
        self._xl = openpyxl.open(filename, read_only)
        self._curr_ws = None

    def _check_ws(self):
        if self._curr_ws is None:
            raise RuntimeError('Please active worksheet first')

    def get_sheets_name(self) -> List[str]:
        _check_ws()
        return self._curr_ws.sheetnames

    def active_ws_by_name(self, ws_name: str) -> None:
        self._curr_ws = self._xl[ws_name]

    def elements_slice(self, start: str, end: str) -> List[int, float, str]:
        """Get elements of cell, action like built-in range function in python.
        Notion: [start:end] means double closed region.
        
        Keyword arguments:
        start -- Beginning position of the slice.
        end -- End position of the slice.
        Return: List which contain int/float/str elements in the cells.
        """
        elements = []
        ws_range = "{}:{}".format(start, end)
        for row in self._curr_ws[ws_range]:
            for cell in row:
                elements.append(cell.value)
        return elements
    
    @property
    def rows(self) -> int:
        _check_ws()
        return len(self._curr_ws['A'])
