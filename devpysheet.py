import datetime
from dataclasses import dataclass
from typing import Optional, Union
import logging
import pygsheets

logging.basicConfig(level=logging.DEBUG)
logging.debug('---------- WORKSHEET LOGS ----------')
alphabets = ["A{}", "B{}", "C{}", "D{}", "E{}", "F{}",
             "G{}", "H{}", "I{}", "J{}", "K{}", "L{}",
             "M{}", "N{}", "O{}", "P{}", "Q{}", "R{}",
             "S{}", "T{}", "U{}", "V{}", "W{}", "X{}",
             "Y{}", "Z{}"]


@dataclass
class WorkSheet:
    service_file: str
    spread_sheet_title: str
    work_sheet_title: str

    def get_all_sheets_title(self) -> Optional[list]:
        return self.client.spreadsheet_titles()

    @property
    def client(self):
        return pygsheets.authorize(service_file=self.service_file)

    @property
    def spread_sheet(self):
        return self.client.open(self.spread_sheet_title)

    @property
    def work_sheet(self):
        return self.spread_sheet.worksheet('title', self.work_sheet_title)

    @property
    def get_date_type(self) -> Optional[list]:
        data = list()
        for i in alphabets:
            cell = self.work_sheet.cell(i.format(1)).fetch()
            if cell.value == '':
                break
            if cell.format[0] == 'DATE':
                data.append({
                    'value': cell.value,
                    'label': cell.label
                })
        return data

    def get_label_by_date(self, date: str = None) -> Optional[dict]:
        logging.info("---------- Get label by type ----------")
        data = self.get_date_type
        if date:
            for i in data:
                if i.get('value') == date:
                    return {'label': i.get('label')}
        else:
            today = datetime.date.today().strftime('%d.%m.%Y')
            for i in data:
                if i.get('value') == today:
                    return {'label': i.get('label')}
        return {'label': None}

    def get_address_with_unique_id(self, unique_id: str) -> Union[tuple, None]:
        logging.info("---------- Get Address With Unique Id ----------")
        return self.work_sheet.find(unique_id)[0].address.index if self.work_sheet.find(unique_id) else None

    def set_absent(self, address: Union[tuple, None], date: str = None, absent_chr: str = None):
        logging.info("---------- Set Absent ----------")
        char = absent_chr if absent_chr else " +"
        if date:
            label = self.get_label_by_date(date).get('label')
            address = f'{label[:-1]}{address[0]}'
            self.work_sheet.update_value(address, char)
        else:
            label = self.get_label_by_date().get('label')
            address = f'{label[:-1]}{address[0]}'
            self.work_sheet.update_value(address, char)
