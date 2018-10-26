from pyexcelreporter.excel_interface import ExcelInterface
import pandas as pd


with ExcelInterface(excel_workbook="example_workbook_template.xlsx") as excel_interface:

    iris_df = pd.read_csv("iris.csv")
    excel_interface.write_sheet("iris", df=iris_df, workbook_out="iris.xlsx")
