from util import SUMMARY_INIT, dumps_json, requests_html, get_xlsx, make_data, excel_date
import config

from datetime import datetime, timedelta
from typing import Dict  # , List


class DataJson:
    def __init__(self):
        self.contacts_sheet = get_xlsx(config.contacts_xlsx, "contacts.xlsx")["相談件数"]
        self.patients_html = requests_html("a57337/kenko/health/corona_zokusei.html")
        self.inspections_sheet = get_xlsx(config.inspections_xlsx, "inspections.xlsx")["検査件数・陽性患者"]
        self.main_summary_html = requests_html("/a73576/kenko/health/infection/protection/covid_19.html")
        # self.main_summary_sheet = get_xlsx(config.main_summary_xlsx, "main_summary.xlsx")["kobe"]
        self.inspections_count = 4
        self.contacts_count = 6
        # self.summary_count = 2
        self.main_summary_values = []
        self.last_update = datetime.today().strftime("%Y/%m/%d %H:%M")  # TODO: 参照データの最終更新日時を入れる
        self._data_json = {}
        # 以下内部変数
        self._window_contacts_json = {}
        self._center_contacts_json = {}
        self._health_center_summary_json = {}
        self._patients_json = {}
        self._patients_summary_json = {}
        self._inspections_summary_json = {}
        self._main_summary_json = {}
        # 初期化
        self.get_inspections()
        self.get_contacts()

    def template_json(self) -> Dict:
        # テンプレート、これをもとにデータを追加していく
        return {
            "date": self.last_update,
            "data": []
        }

    def data_json(self) -> Dict:
        # 内部変数にデータが保管されているか否かを確認し、保管されていなければ生成し、返す。
        # 以下Dictを返す関数はこれに同じ
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
        # window_contactsとcenter_contactsを一緒に生成する。
        # スクリプト実行時間短縮のため、同時に生成している。
        self._window_contacts_json = self.template_json()
        self._center_contacts_json = self.template_json()

        for i in range(6, self.contacts_count):
            window_data = {}
            center_data = {}
            # 日時の取得
            date = excel_date(self.contacts_sheet.cell(row=i, column=1).value)
            # 日別窓口相談者数の取得
            window_contacts = self.contacts_sheet.cell(row=i, column=5).value
            # 日別帰国者・接触者コールセンター相談者数の取得
            center_contacts = self.contacts_sheet.cell(row=i, column=6).value
            # Excelのセル内に0すら入っていないときはNoneが返ってくるので、0を代入しなおす。
            if window_contacts is None:
                window_contacts = 0
            if center_contacts is None:
                center_contacts = 0
            # iso formatで日時を代入
            window_data["日付"] = center_data["日付"] = date.isoformat() + "Z"
            window_data["小計"] = window_contacts
            center_data["小計"] = center_contacts
            self._window_contacts_json["data"].append(window_data)
            self._center_contacts_json["data"].append(center_data)

    def make_health_center_summary(self) -> None:

        # health_center_summaryを生成する
        self._health_center_summary_json = self.template_json()

        for i in range(6, self.contacts_count):
            # 日時がExcel日付で入力されているので、それを変換し代入。
            # 詳しくは https://qiita.com/nezumi/items/23c301c661f5e9653f19 参照
            date = excel_date(self.contacts_sheet.cell(row=i, column=1).value)
            # 保健センター相談者数の取得
            health_center = self.contacts_sheet.cell(row=i, column=2).value
            # 予防衛生課相談者数の取得
            hygiene_section = self.contacts_sheet.cell(row=i, column=3).value
            # Excelのセル内に0すら入っていないときはNoneが返ってくるので、0を代入しなおす。
            if health_center is None:
                health_center = 0
            if hygiene_section is None:
                hygiene_section = 0
            self._health_center_summary_json["data"].append(
                make_data(date.isoformat() + "Z", health_center + hygiene_section)
            )

    def make_patients(self) -> None:
        # patientsを生成する
        self._patients_json = self.template_json()

        # patientsは現状HTMLの表を使用して作成しているので、テーブル(レコード一覧)を取得する
        tables = self.patients_html.find_all("tr")
        # レコード(セル一覧)を取得する
        for cells in tables:
            # データの初期化
            data = {}
            # セルを取得する
            for i, cell in enumerate(cells.find_all("td")):
                # 何もないセルは"\u3000"が埋め込まれているのでリプレースして消す
                text = cell.get_text().replace("\u3000", "")
                # カラム名は飛ばす
                if text == "番号":  # i == 0と同意
                    break
                # 日付を取得する
                elif i == 1:
                    date = datetime.strptime("2020年" + text, "%Y年%m月%d日") + timedelta(hours=8)
                    data["リリース日"] = date.isoformat() + "Z"
                    data["date"] = date.strftime("%Y-%m-%d")
                # 年代を取得する
                elif i == 2:
                    data["年代"] = text + "代"
                # 性別を取得する
                elif i == 3:
                    data["性別"] = text
                # 備考を取得する
                elif i == 5:
                    # 市外在住を除外するため、備考欄を利用
                    if text == "市外在住":
                        # 下のif文で引っかからないようデータを初期化
                        data = {}
                        break
                    data["備考"] = text if text != "\xa0" and text else None
            if data:
                data["退院"] = None  # TODO: 退院データが現状ないため保留
                self._patients_json["data"].append(data)
        # 市外発表者も含むため、日時順でソート
        self._patients_json["data"].sort(key=lambda x: x['date'])

    def make_patients_summary(self) -> None:
        # patients_summaryを生成する
        self._patients_summary_json = self.template_json()

        # 日時の入力が不規則なので、最初のデータを参考に、一日ずつ追加する方式で。
        prev_date = (
                datetime.strptime("2020/" + self.inspections_sheet.cell(row=4, column=1).value, "%Y/%m/%d") -
                timedelta(days=1)
        )
        for i in range(4, self.inspections_count):
            date = prev_date + timedelta(days=1)
            # 陽背患者数を取得する
            patients = self.inspections_sheet.cell(row=i, column=10).value
            if patients is None:
                patients = 0
            self._patients_summary_json["data"].append(make_data(date.isoformat() + "Z", patients))
            prev_date = date

    def make_inspections_summary(self) -> None:
        # inspections_summaryを生成
        self._inspections_summary_json = self.template_json()

        # 日時の入力が不規則なので、最初のデータを参考に、一日ずつ追加する方式で。
        prev_date = (
                datetime.strptime("2020/" + self.inspections_sheet.cell(row=4, column=1).value, "%Y/%m/%d") -
                timedelta(days=1)
        )
        for i in range(4, self.inspections_count):
            date = prev_date + timedelta(days=1)
            # 検査人数を取得する
            inspections = self.inspections_sheet.cell(row=i, column=2).value
            self._inspections_summary_json["data"].append(make_data(date.isoformat() + "Z", inspections))
            prev_date = date

    def make_main_summary(self) -> None:
        # main_summaryの生成
        self._main_summary_json = SUMMARY_INIT

        # main_summaryはHTMLの表を使用して作成しているので、テーブル(レコード一覧)を取得する
        tables = self.main_summary_html.find_all("tr")
        # レコード(セル一覧)を取得する
        for i, cells in enumerate(tables):
            # https://www.city.kobe.lg.jp/a73576/kenko/health/infection/protection/covid_19.html の
            # 検査件数総数(i == 3, j == 0の場所)のデータと神戸市内在住者合計(i == 5, j == 0以降)のデータを使うので、
            # それ以外の行は読み飛ばす
            if i not in [3, 5]:
                continue
            for j, cell in enumerate(cells.find_all("td")):
                # 検査件数総数以外は使わないのでbreakさせる
                if i == 3 and j > 0:
                    break
                # 一番最初に「神戸市内在住者合計」とあって、データではないので読み飛ばす
                if i == 5 and j == 0:
                    continue
                # テキストをリストとして取得
                text_list = cell.get_text().split()
                # text_list[0]が数字のみの場合(「10」など)はtryでそのまま成功するが、「10人」などの場合はexpectで処理させる。
                try:
                    value = int(text_list[0])
                except Exception:
                    value = int(text_list[0][:-1])
                self.main_summary_values.append(value)
        # 県のExcelデータを用いるために使用していたが、市外在住者の扱いを統一するためHPをスクレイピングしたものを使うことになったのでコメントアウト
        # self.main_summary_values = self.get_main_summary_values()
        self.set_summary_values(self._main_summary_json)

    def set_summary_values(self, obj) -> None:
        # リストの先頭の値を"value"にセットする
        obj["value"] = self.main_summary_values[0]
        # objが辞書型で"children"を持っている場合のみ実行
        if isinstance(obj, dict) and obj.get("children"):
            for child in obj["children"]:
                # 再起させて値をセット
                self.main_summary_values = self.main_summary_values[1:]
                self.set_summary_values(child)

    # 県のExcelデータを用いるために使用していたが、市外在住者の扱いを統一するためHPをスクレイピングしたものを使うことになったのでコメントアウト
    # def get_main_summary_values(self) -> List:
    #     values = []
    #     for i in range(2, 9):
    #         values.append(self.main_summary_sheet.cell(row=self.summary_count - 1, column=i).value)
    #      return values

    def get_inspections(self) -> None:
        # 何行分検査数のデータがあるかを取得
        while self.inspections_sheet:
            self.inspections_count += 1
            value = self.inspections_sheet.cell(row=self.inspections_count, column=2).value
            if value is None:
                break

    def get_contacts(self) -> None:
        # 何行分相談数のデータがあるかを取得
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

    # 県のExcelデータを用いるために使用していたが、市外在住者の扱いを統一するためHPをスクレイピングしたものを使うことになったのでコメントアウト
    # def get_summary_count(self) -> None:
    #     while self.main_summary_sheet:
    #         self.summary_count += 1
    #         value = self.main_summary_sheet.cell(row=self.summary_count, column=1).value
    #         if not value:
    #             break


if __name__ == '__main__':
    dumps_json("data.json", DataJson().data_json())
