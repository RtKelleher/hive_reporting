#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pylint: disable=too-many-instance-attributes
"""
Hive-Reporting provides easy to read case metrics supporting team contirubtions
and frequency without the need to access or create custom report in
The Hive Dashboard

"""
from __future__ import print_function
from __future__ import unicode_literals
import datetime
import smtplib
import time
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from thehive4py.api import TheHiveApi

import pandas as pd

API = TheHiveApi("", "")
SMTP_SERVER = ""
SENT_TO = "comma,seperated,as,needed"


class SIRPPipeline(object):
    """Security Incident Response Platform prosessing pipeline.

    Attributes:
        TIME_FMT (str): Time format.
        data_frame_INDEX (list): List of desired parsed values from dictionary.
    """

    TIME_FMT = "%m/%d/%Y %H:%M:%S"
    data_frame_INDEX = ["Created", "Severity", "Owner", "Name", "Closed", "Resolution"]
    counts_frame_INDEX = [
        "totals",
        "Team.Member",
        "Team.Member1",
        "Team.Member2",
        "Team.Member3",
        "Team.Member4",
        "Team.Member5",
        "Team.Member6",
        "Duplicated",
        "TruePositive",
        "FalsePositive",
        int("1"),
        int("2"),
        int("3"),
    ]

    def __init__(
        self, api,
    ):
        """
        Security Incident Response Platform prosessing pipeline.
        Accepts API object on initialization phase.
        """
        self._api = api
        self._api_response = self._api.find_cases(range="all", sort=[])
        self._all30_dict = {}
        self._all60_dict = {}
        self._all90_dict = {}

        self._data_frame_30days = None
        self._data_frame_60days = None
        self._data_frame_90days = None
        self._data_frame_counts = None
        self._dataset = None

    def _load_data(self):
        """Finds all cases on SIRP endpoint

        Returns:
            (obj): api_response
        """
        if self._api_response.status_code == 200:
            self._dataset = self._api_response.json()
            self._fill_day_dicts()

    @staticmethod
    def _add_record(days_dict, record, key):
        """creates objects for dictionary
            (obj): Name
            (obj): Owner
            (obj): Severity
            (obj): Created
            (obj): Closed
            (obj): Resolution
        """
        days_dict[key] = {
            "Name": record["title"],
            "Owner": record["owner"],
            "Severity": record["severity"],
            "Created": (time.strftime(SIRPPipeline.TIME_FMT, time.gmtime(record["createdAt"] / 1000.0))),
        }
        if "endDate" in record:
            days_dict[key].update(
                {
                    "Closed": (time.strftime(SIRPPipeline.TIME_FMT, time.gmtime(record["endDate"] / 1000.0),)),
                    "Resolution": record["resolutionStatus"],
                }
            )

    def _fill_day_dicts(self):
        """Set keys for dictionary based on comparitive EPOCH
            (obj): self._all30_dict
            (obj): self._all60_dict
            (obj): self._all90_dict

        Returns:
            Date corrected (obj)
        """
        today = datetime.date.today()
        for i, record in enumerate(self._dataset):
            if (record["createdAt"] / 1000) > time.mktime((today - datetime.timedelta(days=30)).timetuple()):
                self._add_record(self._all30_dict, record, key=i)

            elif (record["createdAt"] / 1000) > time.mktime((today - datetime.timedelta(days=60)).timetuple()):
                self._add_record(self._all60_dict, record, key=i)

            else:
                self._add_record(self._all90_dict, record, key=i)

    def make_dataframes(self):
        """Creates (4) pandas dataframes
            (obj): data_frame_30day
            (obj): data_frame_60days
            (obj): data_frame_90days
            (obj): data_frame_counts
        """
        self._data_frame_30days = pd.DataFrame(self._all30_dict, index=SIRPPipeline.data_frame_INDEX).transpose()
        self._data_frame_60days = pd.DataFrame(self._all60_dict, index=SIRPPipeline.data_frame_INDEX).transpose()
        self._data_frame_90days = pd.DataFrame(self._all90_dict, index=SIRPPipeline.data_frame_INDEX).transpose()
        self._data_frame_counts = pd.DataFrame(
            {
                "Created": {"totals": self._data_frame_30days.count()["Created"]},
                "Closed": {"totals": self._data_frame_30days.count()["Closed"]},
                "Owner": (self._data_frame_30days["Owner"].value_counts().to_dict()),
                "Resolution": (self._data_frame_30days["Resolution"].value_counts().to_dict()),
                "Severity": (self._data_frame_30days["Severity"].value_counts().to_dict()),
            },
            index=self.counts_frame_INDEX,
        )
        self._data_frame_counts.fillna(0, inplace=True)

    @staticmethod
    def _set_workbook_layout(workbook, worksheet, data_frame):
        """Workbook settings

        Args:
            workbook (obj): excel object
            worksheet (obj): worksheet object
            data_frame (obj): pandas dataframe
        """
        worksheet.set_column("A:A", 3.5, workbook.add_format())
        worksheet.set_column("B:B", 17.25, workbook.add_format())
        worksheet.set_column("C:C", 10, workbook.add_format())
        worksheet.set_column("D:D", 10, workbook.add_format())
        worksheet.set_column("E:E", 100, workbook.add_format())
        worksheet.set_column("F:F", 17.25, workbook.add_format())
        worksheet.set_column("G:G", 11.25, workbook.add_format())
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter("A1:G100")

        header_format = workbook.add_format(
            {"bold": True, "text_wrap": True, "valign": "top", "fg_color": "#CCCCCC", "border": 1}
        )

        for col_num, value in enumerate(data_frame.columns.values, 1):
            worksheet.write(0, col_num, value, header_format)

    def make_charts(self):
        """Create charts for workbook

        Args:
            wbook (obj): excel object
            wsheet (obj): worksheet object
            title (str): name of chart
            cell_pos (str): chart position on sheet
            series (func): pandas series
        """

        def _insert_pie_chart(wbook, wsheet, title, cell_pos, series):
            piechart = wbook.add_chart({"type": "pie"})
            piechart.set_title({"name": title})
            piechart.set_style(10)
            piechart.add_series(series)
            wsheet.insert_chart(cell_pos, piechart, {"x_offset": 25, "y_offset": 10})

        def _data_frame_days_to_excel(writer, sheet_name, data_frame_days):
            data_frame_days.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False)
            self._set_workbook_layout(writer.book, (writer.sheets[sheet_name]), data_frame_days)

        with pd.ExcelWriter("Hive Metrics.xlsx", engine="xlsxwriter", options={"strings_to_urls": False}) as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Summary Charts")
            worksheet.hide_gridlines(2)

            _insert_pie_chart(
                workbook,
                worksheet,
                title="New vs. Closed Cases",
                cell_pos="D2",
                series={
                    "name": "Open vs. Closed Cases Last 30",
                    "categories": "=Tracking!$B$1:$C$1",
                    "values": "=Tracking!$B$2:$C$2",
                },
            )
            _insert_pie_chart(
                workbook,
                worksheet,
                title="Case Ownership",
                cell_pos="M19",
                series={
                    "name": "Case Ownership Last 30",
                    "categories": "=Tracking!$A$3:$A$9",
                    "values": "=Tracking!$D$3:$D$9",
                },
            )
            _insert_pie_chart(
                workbook,
                worksheet,
                title="Case Resolution",
                cell_pos="D19",
                series={
                    "name": "Case Resolution Last 30",
                    "categories": "=Tracking!$A$10:$A$12",
                    "values": "=Tracking!$E$10:$E$12",
                },
            )
            _insert_pie_chart(
                workbook,
                worksheet,
                title="Case Severities",
                cell_pos="M2",
                series={
                    "name": "Severity Last 30",
                    "categories": "=Tracking!$A$13:$A$15",
                    "values": "=Tracking!$F$13:$F$15",
                },
            )

            _data_frame_days_to_excel(
                writer, sheet_name="Cases newer than 30 Days", data_frame_days=self._data_frame_30days,
            )
            _data_frame_days_to_excel(
                writer, sheet_name="Cases older than 60 days", data_frame_days=self._data_frame_60days,
            )
            _data_frame_days_to_excel(
                writer, sheet_name="Cases newer than 90 Days", data_frame_days=self._data_frame_90days,
            )

            self._data_frame_counts.to_excel(writer, sheet_name="Tracking")
            writer.save()

    @staticmethod
    def send_mail():
        """Create email and send excel object

        Vars:
            SENT_TO(str): recipient
            SMTP_SERVER(str): mail server
        """
        msg = MIMEMultipart()
        msg["From"] = "SIRP-Reminders@company.com"
        msg["To"] = SENT_TO
        msg["Subject"] = "The Hive Case Metrics"
        msg.attach(MIMEText("Attached are the requested case metrics in .XLSX format."))
        part = MIMEBase("application", "octet-stream")
        part.set_payload(open("Hive Metrics.xlsx", "rb").read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", 'attachment; filename="Hive Metrics.xlsx"')
        msg.attach(part)
        smtp = smtplib.SMTP(SMTP_SERVER)
        smtp.starttls()
        smtp.sendmail(msg["From"], msg["To"].split(","), msg.as_string())
        smtp.quit()

    def run(self):
        """runtime setup"""
        self._load_data()
        self.make_dataframes()  # may be protected
        self.make_charts()  # may be protected
        self.send_mail()


def main(api):
    """api initilization"""
    pipe = SIRPPipeline(api)
    pipe.run()


main(API)
sys.exit()
