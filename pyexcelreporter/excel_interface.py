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

    def write_sheet(self, sheet_name, df):
        """
        Update given sheet with a pandas dataframe
        :param sheet_name: name of sheet to update
        :param df: pandas dataframe
        :return:
        """

        writer = pandas.ExcelWriter(self.__workbook_loc, engine='openpyxl')
        writer.book = self.__workbook
        writer.sheets = dict((ws.title, ws) for ws in self.__workbook.worksheets)

        df.to_excel(writer, sheet_name, cols=list(df.columns))

        writer.save()
