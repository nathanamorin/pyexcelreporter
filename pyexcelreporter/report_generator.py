import os
import tempfile

from pyexcelreporter.email_utils import EmailConnection, Email
from pyexcelreporter.excel_interface import ExcelInterface

class ReportGenerator:

    __files = []

    def __init__(self, email_host, email_port, username, password, secure):
        """

        :param email_host:
        :param email_port:
        :param username:
        :param password:
        """
        self.__email_host = email_host
        self.__email_port = email_port
        self.__username = username
        self.__password = password
        self.__secure = secure

    def __enter__(self):
        self.__tmp_dir = tempfile.TemporaryDirectory()
        self.__server = EmailConnection(server=self.__email_host, port=self.__email_port, username=self.__username,
                                        password=self.__password, secure=self.__secure)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.__tmp_dir.cleanup()
        self.__server.close()

    def add_file(self, excel_workbook_template, output_filename, sheets):
        """

        :param excel_workbook_template: Path to excel workbook template
        :param sheets: dictionary {"sheetname" : dataframe_to_populate }
        :return:
        """

        out_file = os.path.join(self.__tmp_dir.name, output_filename)

        with ExcelInterface(excel_workbook_template) as excel_interface:
            for sheet, df in dict(sheets).items():
                excel_interface.write_sheet(sheet, df, workbook_out=out_file)

        print(out_file)
        self.__files.append(out_file)

    def send_report(self, subject, message, to_emails, from_email):
        """

        :param subject:
        :param message:
        :param to_emails:
        :param from_email:
        :return:
        """

        email = Email(from_email=from_email,
                      to_emails=to_emails,
                      subject=subject, message=message, attachments=self.__files)
        self.__server.send(email)


