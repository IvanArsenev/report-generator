"""Скрипт для анализа арендных платежей на основе банковской выписки"""

import os
import fire
import smtplib
import pandas as pd
from email.message import EmailMessage
from datetime import datetime
import logging


logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


def months_between(start_date, end_date):
    result = (end_date.year - start_date.year) * 12 + (end_date.month - start_date.month) - 1
    return result


class RentReportGenerator:
    """Класс для генерации отчета по арендным платежам"""

    def __init__(
        self,
        bank_file_path="docs/bank_report.xlsx",
        rent_file_path="docs/arenda_list.xlsx",
        max_different=0,
        report_dir="reports",
        email=""
    ):
        """
        Инициализация пути к файлам

        Args:
            bank_file_path (str): Путь к файлу банковской выписки
            rent_file_path (str): Путь к файлу с арендными данными
            max_different (int): Максимальная разница для похожих платежей
            report_dir (str): Папка для сохранения отчета
            email (str): Если указан, ответ будет отправлен по email
        """
        self.bank_file_path = bank_file_path
        self.rent_file_path = rent_file_path
        self.max_different = max_different
        self.report_dir = report_dir
        self.email = email
        os.makedirs(self.report_dir, exist_ok=True)

    def _send_email_with_attachment(self, file_path: str):
        """
        Отправляет отчет по электронной почте с вложением

        Args:
            file_path (str): Путь к Excel-файлу отчета
        """
        msg = EmailMessage()
        msg["Subject"] = "Отчет по арендным платежам"
        msg["From"] = "Private Person <hello@demomailtrap.co>"
        msg["To"] = self.email
        msg.set_content("Во вложении отчет по арендным платежам")

        try:
            with open(file_path, "rb") as f:
                file_data = f.read()
                filename = os.path.basename(file_path)
                msg.add_attachment(
                    file_data,
                    maintype="application",
                    subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename=filename
                )

            with smtplib.SMTP("live.smtp.mailtrap.io", 587) as server:
                server.starttls()
                server.login("api", "")
                server.send_message(msg)

            logger.info(f"Отчет успешно отправлен на {self.email}")
        except Exception as e:
            logger.error(f"Не удалось отправить письмо: {e}")

    def generate_report(self):
        """
        Генерирует Excel-отчет по арендным платежам на основе данных из Excel-файлов
        """
        bank_data = pd.read_excel(self.bank_file_path, usecols="A,D,E")
        pattern = r'^\d{2}\.\d{2}\.\d{4} \d{2}\.\d{2}\.\d{4}$'

        transfers = bank_data[bank_data['Unnamed: 0'].str.match(pattern, na=False)].copy()
        transfers.loc[:, 'Unnamed: 0'] = pd.to_datetime(
            transfers['Unnamed: 0'].str.extract(r'(\d{2}\.\d{2}\.\d{4})')[0],
            format='%d.%m.%Y'
        ).dt.date

        transfers.columns = ["Дата операции", "Тип операции", "Сумма"]
        transfers.loc[:, 'Сумма'] = (
            transfers['Сумма']
            .astype(str)
            .str.replace(r'[^\d,]', '', regex=True)
            .str.split(',')
            .str[0]
            .astype(int)
        )

        rent_data = pd.read_excel(self.rent_file_path)
        rent_data['Первоначальная дата'] = pd.to_datetime(rent_data['Первоначальная дата']).dt.date
        report_rows = []

        for _, rent_row in rent_data.iterrows():
            expected_sum = rent_row['Сумма']
            garage = rent_row['Гараж']

            mask = (transfers['Сумма'] >= expected_sum - self.max_different) & \
                   (transfers['Сумма'] <= expected_sum + self.max_different)
            matched = transfers[mask]
            months_count = months_between(rent_row["Первоначальная дата"], datetime.today().date())

            if matched.empty:
                must_pay = expected_sum * months_count
                report_rows.append({
                    "Гараж": garage,
                    "Первоначальная дата": rent_row["Первоначальная дата"],
                    "Тип": "Нет платежей",
                    "Дата": "",
                    "Описание": "",
                    "Сумма": "",
                    "Арендная плата": expected_sum,
                    "Разница": "",
                    "Долг": must_pay
                })
                continue

            exact_matches = []
            similar_matches = []
            actual_sum = matched["Сумма"].sum()
            for _, row in matched.iterrows():

                diff = row["Сумма"] - expected_sum
                must_pay = expected_sum * months_count - actual_sum
                record = {
                    "Гараж": garage,
                    "Первоначальная дата": rent_row["Первоначальная дата"],
                    "Тип": "Оплачено" if diff == 0 else "Похожий платеж",
                    "Дата": row["Дата операции"],
                    "Описание": row["Тип операции"],
                    "Сумма": row["Сумма"],
                    "Арендная плата": expected_sum,
                    "Разница": diff,
                    "Долг": must_pay
                }
                if diff == 0:
                    exact_matches.append(record)
                else:
                    similar_matches.append(record)

            report_rows.extend(exact_matches + similar_matches)

        df_report = pd.DataFrame(report_rows)
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        report_path = os.path.join(self.report_dir, f"rent_report_{timestamp}.xlsx")
        df_report.to_excel(report_path, index=False)
        if self.email:
            self._send_email_with_attachment(report_path)
        else:
            logger.info(f"Отчет сохранен в {report_path}")


if __name__ == '__main__':
    fire.Fire(RentReportGenerator)
