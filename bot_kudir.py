import os
import logging
from from aiogram import Bot, Dispatcher, executor, types
import pandas as pd
from io import BytesIO
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO)

API_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
if not API_TOKEN:
    raise ValueError("Переменная окружения TELEGRAM_BOT_TOKEN не задана!")

bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)

UNP_KOZEL = "291530425"
UNP_KONDRASCHUK = "291156481"

@dp.message_handler(commands=['start'])
async def start(message: types.Message):
    await message.answer("Привет! Отправь выписку из банка (.xls или .xlsx), и я создам КУДиР.")

@dp.message_handler(content_types=types.ContentTypes.DOCUMENT)
async def handle_excel(message: types.Message):
    doc = message.document
    if not doc.file_name.endswith(('.xls', '.xlsx')):
        await message.answer("Пожалуйста, отправьте файл в формате .xls или .xlsx")
        return

    try:
        file_info = await bot.get_file(doc.file_id)
        file = await bot.download_file(file_info.file_path)
        df = pd.read_excel(file, skiprows=10)

        # Определяем колонки по индексу
        date_col = df.columns[1]
        credit_col = df.columns[9]
        debit_col = df.columns[8]
        purpose_col = df.columns[12]
        unp_col = df.columns[11]
        doc_num_col = df.columns[0]

        records = []
        total_income = 0.0
        total_expenses = 0.0
        other_income = 0.0

        for _, row in df.iterrows():
            credit = row[credit_col] if pd.notna(row[credit_col]) else 0.0
            debit = row[debit_col] if pd.notna(row[debit_col]) else 0.0
            purpose = str(row[purpose_col]) if pd.notna(row[purpose_col]) else ""
            unp = str(row[unp_col]) if pd.notna(row[unp_col]) else ""

            # Проценты от банка → иные поступления (графа 7)
            if credit > 0 and 'процент' in purpose.lower():
                other_income += credit
                records.append({'date': row[date_col], 'doc': f"ПП №{row[doc_num_col]} от {row[date_col]}", 'desc': purpose[:100], 'income': 0.0, 'other_income': credit, 'expense': 0.0})
                continue

            # Доход от предпринимательской деятельности → графа 4
            if credit > 0 and 'ОАО "Белагропромбанк"' not in purpose and 'налог' not in purpose.lower() and 'возврат' not in purpose.lower():
                total_income += credit
                records.append({'date': row[date_col], 'doc': f"ПП №{row[doc_num_col]} от {row[date_col]}", 'desc': purpose[:100], 'income': credit, 'other_income': 0.0, 'expense': 0.0})
                continue

            # Расходы (банковские комиссии и услуги контрагентов)
            if debit > 0:
                if 'ОАО "Белагропромбанк"' in purpose and ('Абонентская плата' in purpose or 'Комиссионное вознаграждение' in purpose):
                    total_expenses += debit
                    records.append({'date': row[date_col], 'doc': f"ПП №{row[doc_num_col]} от {row[date_col]}", 'desc': purpose[:100], 'income': 0.0, 'other_income': 0.0, 'expense': debit})
                    continue
                if unp in [UNP_KOZEL, UNP_KONDRASCHUK]:
                    total_expenses += debit
                    records.append({'date': row[date_col], 'doc': f"ПП №{row[doc_num_col]} от {row[date_col]}", 'desc': purpose[:100], 'income': 0.0, 'other_income': 0.0, 'expense': debit})
                    continue

        # Формируем КУДиР
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "КУДиР"

        headers = [
            "Дата записи", "Наименование документа, его номер, дата", "Содержание хоз. операции",
            "Доходы, учитываемые в отчетном периоде (сумма)", 
            "Доходы, учитываемые в отчетном периоде (сумма налогов, сборов, уплаченная из выручки)",
            "Освобождаемые доходы, сумма", "Иные поступления",
            "Расходы, приходящиеся на отчетный период",
            "Расходы по нормативу", "Иные расходы", "Примечание"
        ]
        ws.append(headers)

        for rec in records:
            ws.append([
                rec['date'],
                rec['doc'],
                rec['desc'],
                rec['income'],
                "",
                "",
                rec['other_income'],
                rec['expense'],
                "",
                "",
                ""
            ])

        tax_base = total_income - total_expenses
        tax = round(tax_base * 0.2, 2)

        ws.append(["", "", "ИТОГО за квартал", total_income, "", "", other_income, total_expenses, "", "", ""])
        ws.append(["", "", "Налогооблагаемая база", tax_base, "", "", "", "", "", "", ""])
        ws.append(["", "", "Подоходный налог (20%)", tax, "", "", "", "", "", "", ""])

        wb.save(output)
        output.seek(0)

        await message.answer_document(
            document=types.InputFile(output, filename="KUDiR.xlsx"),
            caption=f"✅ Готово!\n\nДоходы: {total_income:.2f} BYN\nНалог: {tax:.2f} BYN"
        )

    except Exception as e:
        await message.answer(f"Ошибка: {str(e)}")

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
