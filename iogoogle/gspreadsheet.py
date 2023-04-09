import dataclasses
import enum

from typing import Any

import googleapiclient.discovery
import googleapiclient.http
import googleapiclient.errors

from google.oauth2 import service_account

ROWS = "ROWS"
COLUMNS = "COLUMNS"


class GSpreadsheetScope(str, enum.Enum):
    """
    Defining scopes for Google Sheets API V4
    https://developers.google.com/identity/protocols/oauth2/scopes#sheets

    drive :: See, edit, create, and delete all of your Google Drive files
    drive.file :: See, edit, create, and delete only the specific Google Drive files you use with this app
    drive.readonly :: See and download all your Google Drive files
    spreadsheets :: See, edit, create, and delete all your Google Sheets spreadsheets
    spreadsheets.readonly :: See all your Google Sheets spreadsheets
    """

    DRIVE = "https://www.googleapis.com/auth/drive"
    DRIVE_FILE = "https://www.googleapis.com/auth/drive.file"
    DRIVE_READONLY = "https://www.googleapis.com/auth/drive.readonly"
    SPREADSHEETS = "https://www.googleapis.com/auth/spreadsheets"
    SPREADSHEETS_READONLY = "https://www.googleapis.com/auth/spreadsheets.readonly"


@dataclasses.dataclass
class GSheet:
    name: str
    sheet_id: str


@dataclasses.dataclass
class Range:
    range_str: str
    values: list[list[Any]] = dataclasses.field(default_factory=list)
    dimension: str = "COLUMNS"


@dataclasses.dataclass
class GSpreadSheet:
    spreadsheet_id: str
    credentials: service_account.Credentials = dataclasses.field(repr=False)
    sheets: dict[str, GSheet] = dataclasses.field(
        init=False, repr=False, default_factory=dict
    )

    def __post_init__(self):
        self.set_sheets_of_the_spreadsheet()

    @property
    def service(self) -> googleapiclient.discovery.Resource:
        return googleapiclient.discovery.build(
            serviceName="sheets", version="v4", credentials=self.credentials
        )

    def execute(self) -> None:
        ...

    def add_sheet(self, gsheet: GSheet) -> None:
        self.sheets[gsheet.name] = gsheet

    def remove_sheet(self, name: str) -> None:
        self.sheets.pop(name)

    def get_sheets_info(self) -> dict[str, GSheet]:
        return self.sheets

    def exists_range(self, name: str) -> bool:
        return True

    def set_sheets_of_the_spreadsheet(self) -> None:
        sheets_metadata = (
            self.service.spreadsheets().get(spreadsheetId=self.spreadsheet_id).execute()
        ).get("sheets", "")

        for i in sheets_metadata:
            property_ = i["properties"]
            self.add_sheet(GSheet(property_["title"], property_["sheetId"]))

    def clear_spreadsheet_contents(self, name: str) -> None:
        self.service.spreadsheets().values().clear(
            spreadsheetId=self.spreadsheet_id,
            range=f"{name}!R1C1:R100C100",
        ).execute()

    def get_values(
        self,
        range_: str,
        dimension: str = "columns",
    ) -> list[list[str]]:
        return (
            self.service.spreadsheets()
            .values()
            .get(
                spreadsheetId=self.spreadsheet_id,
                range=range_,
                majorDimension=dimension.upper(),
            )
            .execute()
            .get("values", [])
        )

    def get_batch_values(
        self,
        ranges: list[str],
        dimension: str = "columns",
    ) -> list[list[list[str]]]:
        values_rangewise: list[list[str]] = (
            self.service.spreadsheets()
            .values()
            .batchGet(
                spreadsheetId=self.spreadsheet_id,
                ranges=ranges,
                majorDimension=dimension.upper(),
            )
            .execute()
            .get("valueRanges", [])
        )
        return [i.get("values", []) for i in values_rangewise]

    def update_values(
        self,
        range_: str,
        values: list[list[int | float | str]],
        dimension: str = "columns",
    ) -> None:
        self.service.spreadsheets().values().update(
            spreadsheetId=self.spreadsheet_id,
            range=range_,
            body={
                "values": values,
                "majorDimension": dimension.upper(),
            },
            valueInputOption="USER_ENTERED",
        ).execute()

    def update_batch_values(
        self,
        ranges: list[Range],
    ) -> None:
        data = [
            {
                "range": i.range_str,
                "values": i.values,
                "majorDimension": i.dimension,
            }
            for i in ranges
        ]

        body = {
            "data": data,
            "valueInputOption": "USER_ENTERED",
        }

        self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheet_id,
            body=body,
        ).execute()

    # def update_cell_format(
    #     self,
    #     sheet_name: str,
    #     start_row_index: int,
    #     end_row_index: int,
    #     start_col_index: int,
    #     end_col_index: int,
    #     bg_color: tuple(int, int, int),
    #     text_color: tuple(int, int, int),
    #     bold: bool = False,
    #     horizontal_alignment: str = "CENTER",
    # ) -> None:
    #     background_color = {
    #         "red": bg_color[0] / 255,
    #         "green": bg_color[0] / 255,
    #         "blue": bg_color[0] / 255,
    #     }

    #     self.service.spreadsheets().batchUpdate(
    #         spreadsheetId=self.spreadsheet_id,
    #         body=(
    #             {
    #                 "requests": [
    #                     {
    #                         "repeatCell": {
    #                             "range": {
    #                                 "sheetId": self.sheets[sheet_name].sheet_id,
    #                                 "startRowIndex": start_row_index,
    #                                 "endRowIndex": end_row_index,
    #                                 "startColumnIndex": start_col_index,
    #                                 "endColumnIndex": end_col_index,
    #                             },
    #                             "cell": {
    #                                 "userEnteredFormat": {
    #                                     "backgroundColor": background_color,
    #                                     "horizontalAlignment": horizontal_alignment,
    #                                     "textFormat": {
    #                                         "bold": bold,
    #                                     },
    #                                 }
    #                             },
    #                             "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
    #                         }
    #                     }
    #                 ]
    #             }
    #         ),
    #     ).execute()
