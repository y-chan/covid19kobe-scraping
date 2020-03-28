from util import SUMMARY_INIT, dumps_json, requests_html, get_xlsx, excel_date
import config

from datetime import datetime, timedelta
from typing import Dict, List


class DataJson:
    def __init__(self):
        self.contacts_sheet = get_xlsx(config.contacts_xlsx, "contacts.xlsx")["相談件数"]
        self.patients_html = requests_html("a57337/kenko/health/corona_zokusei.html")
        self.inspections_sheet = get_xlsx(config.inspections_xlsx, "inspections.xlsx")["検査件数・陽性患者"]
        # self.main_summary_html = requests_html("/a73576/kenko/health/infection/protection/covid_19.html")
        self.main_summary_sheet = get_xlsx(config.main_summary_xlsx, "main_summary.xlsx")["kobe"]
        self.inspections_count = 4
        self.contacts_count = 6
        self.summary_count = 2
        self.main_summary_values = []
        self.last_update = datetime.today().strftime("%Y/%m/%d %H:%M")  # TODO: 参照データの最終更新日時を入れる
        self._data_json = {}
        self._window_contacts_json = {}
        self._center_contacts_json = {}
        self._health_center_summary_json = {}
        self._patients_json = {}
        self._patients_summary_json = {}
        self._inspections_summary_json = {}
        self._main_summary_json = {}
        self.get_inspections()
        self.get_contacts()

    def template_json(self) -> Dict:
        return {
            "date": self.last_update,
            "data": []
        }

    def data_json(self) -> Dict:
        if not self._data_json:
            self.make_data()
        return self._data_json

    def window_contacts_json(self) -> Dict:
        if not self._window_contacts_json:
            self.make_contacts()
        return self._window_contacts_json

    def center_contacts_json(self) -> Dict:
        if not self._center_contacts_json:
            self.make_contacts()
        return self._center_contacts_json

    def health_center_summary_json(self) -> Dict:
        if not self._health_center_summary_json:
            self.make_health_center_summary()
        return self._health_center_summary_json

    def patients_json(self) -> Dict:
        if not self._patients_json:
            self.make_patients()
        return self._patients_json

    def patients_summary_json(self) -> Dict:
        if not self._patients_summary_json:
            self.make_patients_summary()
        return self._patients_summary_json

    def inspections_summary_json(self) -> Dict:
        if not self._inspections_summary_json:
            self.make_inspections_summary()
        return self._inspections_summary_json

    def main_summary_json(self) -> Dict:
        if not self._main_summary_json:
            self.make_main_summary()
        return self._main_summary_json

    def make_data(self) -> None:
        self._data_json = {
            "window_contacts": self.window_contacts_json(),
            "center_contacts": self.center_contacts_json(),
            "health_center_summary": self.health_center_summary_json(),
            "patients": self.patients_json(),
            "patients_summary": self.patients_summary_json(),
            "inspections_summary": self.inspections_summary_json(),
            "lastUpdate": self.last_update,
            "main_summary": self.main_summary_json()
        }

    def make_contacts(self) -> None:
        self._window_contacts_json = self.template_json()
        self._center_contacts_json = self.template_json()

        for i in range(6, self.contacts_count):
            window_data = {}
            center_data = {}
            date = excel_date(self.contacts_sheet.cell(row=i, column=1).value)
            window_contacts = self.contacts_sheet.cell(row=i, column=5).value
            center_contacts = self.contacts_sheet.cell(row=i, column=6).value
            if window_contacts is None:
                window_contacts = 0
            if center_contacts is None:
                center_contacts = 0
            window_data["日付"] = center_data["日付"] = date.isoformat() + "Z"
            window_data["小計"] = window_contacts
            center_data["小計"] = center_contacts
            self._window_contacts_json["data"].append(window_data)
            self._center_contacts_json["data"].append(center_data)

    def make_health_center_summary(self) -> None:
        self._health_center_summary_json = {
            "date": self.last_update,
            "data": {
                "保健センター": [],
                "予防衛生課": []
            },
            "labels": []
        }

        for i in range(6, self.contacts_count):
            date = excel_date(self.contacts_sheet.cell(row=i, column=1).value)
            health_center = self.contacts_sheet.cell(row=i, column=2).value
            hygiene_section = self.contacts_sheet.cell(row=i, column=3).value
            if health_center is None:
                health_center = 0
            if hygiene_section is None:
                hygiene_section = 0
            self._health_center_summary_json["data"]["保健センター"].append(health_center)
            self._health_center_summary_json["data"]["予防衛生課"].append(hygiene_section)
            self._health_center_summary_json["labels"].append(date.strftime("%m/%d"))

    def make_patients(self) -> None:
        self._patients_json = self.template_json()

        tables = self.patients_html.find_all("tr")
        for cells in tables:
            data = {}
            for i, cell in enumerate(cells.find_all("td")):
                text = cell.get_text().replace("\u3000", "")
                if text == "番号":
                    break
                if i == 0:
                    continue
                elif i == 1:
                    date = datetime.strptime("2020年" + text, "%Y年%m月%d日") + timedelta(hours=8)
                    data["リリース日"] = date.isoformat() + "Z"
                    data["date"] = date.strftime("%Y-%m-%d")
                elif i == 2:
                    data["年代"] = text + "代"
                elif i == 3:
                    data["性別"] = text
                elif i == 5:
                    data["備考"] = text if text != "\xa0" and text else None
            if data:
                data["退院"] = None  # TODO: 退院データが現状ないため保留
                self._patients_json["data"].append(data)
        self._patients_json["data"].sort(key=lambda x: x['date'])

    def make_patients_summary(self) -> None:
        def make_data(date, value=1):
            return {"日付": date, "小計": value}

        self._patients_summary_json = self.template_json()

        prev_data = {}
        for patients_data in self.patients_json()["data"]:
            date = patients_data["リリース日"]
            if prev_data:
                prev_date = datetime.strptime(prev_data["日付"], "%Y-%m-%dT%H:%M:%SZ")
                patients_zero_days = (datetime.strptime(date, "%Y-%m-%dT%H:%M:%SZ") - prev_date).days
                if prev_data["日付"] == date:
                    prev_data["小計"] += 1
                    continue
                else:
                    self._patients_summary_json["data"].append(prev_data)
                    if patients_zero_days >= 2:
                        for i in range(1, patients_zero_days):
                            self._patients_summary_json["data"].append(
                                make_data((prev_date + timedelta(days=i)).isoformat() + "Z", 0)
                            )
            prev_data = make_data(date)
        self._patients_summary_json["data"].append(prev_data)
        prev_date = datetime.strptime(prev_data["日付"], "%Y-%m-%dT%H:%M:%SZ")
        patients_zero_days = (datetime.now() - prev_date).days
        for i in range(1, patients_zero_days):
            self._patients_summary_json["data"].append(make_data((prev_date + timedelta(days=i)).isoformat() + "Z", 0))

    def make_inspections_summary(self) -> None:
        self._inspections_summary_json = {
            "date": self.last_update,
            "data": {
                "陽性確認者": [],
                "陰性確認者": []
            },
            "labels": []
        }
        prev_date = (
                datetime.strptime("2020/" + self.inspections_sheet.cell(row=4, column=1).value, "%Y/%m/%d") -
                timedelta(days=1)
        )
        for i in range(4, self.inspections_count):
            date = prev_date + timedelta(days=1)
            inspections = self.inspections_sheet.cell(row=i, column=2).value
            patients = self.inspections_sheet.cell(row=i, column=10).value
            if patients is None:
                patients = 0
            self._inspections_summary_json["data"]["陽性確認者"].append(patients)
            self._inspections_summary_json["data"]["陰性確認者"].append(inspections - patients)
            self._inspections_summary_json["labels"].append(date.strftime("%m/%d"))
            prev_date = date

    def make_main_summary(self) -> None:
        self._main_summary_json = SUMMARY_INIT
        # tables = self.main_summary_html.find_all("tr")
        # for i, cells in enumerate(tables):
        #     if i != 3:
        #         continue
        #     for j, cell in enumerate(cells.find_all("td")):
        #         text_list = cell.get_text().split()
        #         try:
        #             value = int(text_list[0])
        #         except Exception:
        #             value = int(text_list[0][:-1])
        #         if len(text_list) == 3 and text_list[2][0] == "(" and text_list[2][-1] == ")":
        #             value -= int(text_list[2][1:-3])
        #         self.main_summary_values.append(value)
        self.main_summary_values = self.get_main_summary_values()
        self.set_summary_values(self._main_summary_json)

    def set_summary_values(self, obj) -> None:
        obj['value'] = self.main_summary_values[0]
        if isinstance(obj, dict) and obj.get('children'):
            for child in obj['children']:
                self.main_summary_values = self.main_summary_values[1:]
                self.set_summary_values(child)

    def get_main_summary_values(self) -> List:
        values = []
        for i in range(2, 9):
            values.append(self.main_summary_sheet.cell(row=self.summary_count - 1, column=i).value)
        return values

    def get_inspections(self) -> None:
        while self.inspections_sheet:
            self.inspections_count += 1
            value = self.inspections_sheet.cell(row=self.inspections_count, column=2).value
            if value is None:
                break

    def get_contacts(self) -> None:
        while self.contacts_sheet:
            self.contacts_count += 1
            value = self.contacts_sheet.cell(row=self.contacts_count, column=1).value
            date = excel_date(value)
            if value is None or date > datetime.today():
                self.contacts_count -= 1
                if (not self.contacts_sheet.cell(row=self.contacts_count, column=2).value and
                        not self.contacts_sheet.cell(row=self.contacts_count, column=3).value and
                        not self.contacts_sheet.cell(row=self.contacts_count, column=5).value and
                        not self.contacts_sheet.cell(row=self.contacts_count, column=6).value):
                    self.contacts_count -= 1
                break

    def get_summary_count(self) -> None:
        while self.main_summary_sheet:
            self.summary_count += 1
            value = self.main_summary_sheet.cell(row=self.summary_count, column=1).value
            if not value:
                break

if __name__ == '__main__':
    dumps_json("data.json", DataJson().data_json())
