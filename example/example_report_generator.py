import os

import pandas as pd
from pyexcelreporter.report_generator import ReportGenerator

# Suggest using https://github.com/mailslurper/mailslurper for local email testing

with ReportGenerator(email_host="localhost", email_port=2500, secure=False,
                     username=None, password=None) as rg:

    iris_df = pd.read_csv("iris.csv")
    rg.add_file(excel_workbook_template="example_workbook_template.xlsx",
                output_filename="iris.xlsx", sheets={"iris": iris_df})

    rg.send_report(subject="Iris Report", message="New report attached",
                   from_email="fisher@iris.org",
                   to_emails=os.environ["EXAMPLE_EMAIL_TO"] if "EXAMPLE_EMAIL_TO" in os.environ else input("Send Email To: "))
