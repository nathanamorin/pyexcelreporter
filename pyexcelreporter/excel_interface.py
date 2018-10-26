import pandas
from openpyxl import load_workbook


class ExcelInterface:
    """
    Interfaces with excel to read/write sheets
    """

    def __init__(self, excel_workbook):
        self.__workbook_loc = excel_workbook

    def __enter__(self):
        self.__workbook = load_workbook(self.__workbook_loc)

        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def write_sheet(self, sheet_name, df, workbook_out=None):
        """
        Update given sheet with a pandas dataframe
        :param workbook_out: Output workbook to write. Default is input sheet
        :param sheet_name: name of sheet to update
        :param df: pandas dataframe
        :return:
        """

        if workbook_out is None:
            workbook_out = self.__workbook_loc

        writer = pandas.ExcelWriter(workbook_out, engine='openpyxl')
        writer.book = self.__workbook
        writer.sheets = dict((ws.title, ws) for ws in self.__workbook.worksheets)

        df.to_excel(writer, sheet_name)

        writer.save()
