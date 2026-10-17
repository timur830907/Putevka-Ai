import os
from docx import Document
from fpdf import FPDF

class WaybillGenerator:
    def __init__(self, data: dict = None):
        # Если данные не переданы, используем пустой словарь
        self.data = data if data is not None else {}

    def create_word(self, filename="waybill_empty.docx"):
        """Генерация чистого шаблона путевого листа в формате Word (DOCX)"""
        doc = Document()
        
        # Заголовок
        doc.add_heading("ЖҮК КӨЛІГІНІҢ ЖОЛ ҚАҒАЗЫ", level=1)
        doc.add_heading("ПУТЕВОЙ ЛИСТ ГРУЗОВОГО АВТОМОБИЛЯ", level=2)
        
        # Основные данные (абсолютно пустые)
        doc.add_paragraph(f"Дата: {self.data.get('date', '____________________')}")
        doc.add_paragraph(f"Автомобиль: {self.data.get('car', '____________________')} | Госномер: {self.data.get('plate', '________________')}")
        doc.add_paragraph(f"Водитель: {self.data.get('driver', '________________________________________')}")
        
        # Блоки допуска
        doc.add_heading("Медицинский и технический контроль", level=3)
        doc.add_paragraph(f"Медосмотр: {self.data.get('medical', '____________________')}")
        doc.add_paragraph(f"Техосмотр (механик): {self.data.get('mechanic', '____________________')}")
        doc.add_paragraph(f"Начальный спидометр: {self.data.get('km_start', '__________')} км | Топливо: {self.data.get('fuel_start', '__________')} л")

        # Маршрут
        doc.add_heading("Задание и Маршрут", level=3)
        doc.add_paragraph(f"Заказчик: {self.data.get('customer', '________________________________________')}")
        doc.add_paragraph(f"Пункт погрузки: {self.data.get('point_load', '________________________________')}")
        doc.add_paragraph(f"Пункт разгрузки: {self.data.get('point_unload', '________________________________')}")
        doc.add_paragraph(f"Характер груза: {self.data.get('cargo', '________________________________')}")

        # Итоги
        doc.add_heading("Итоги смены", level=3)
        doc.add_paragraph(f"Конечный спидометр: {self.data.get('km_end', '__________')} км | Топливо при возврате: {self.data.get('fuel_end', '__________')} л")

        doc.save(filename)
        print(f"✅ Пустой файл Word сохранен как: {filename}")

    def create_pdf(self, filename="waybill_empty.pdf"):
        """Генерация чистого шаблона путевого листа в формате PDF"""
        pdf = FPDF()
        pdf.add_page()
        
        pdf.set_font("Helvetica", size=12)
        pdf.cell(200, 10, txt="PUTEVOY LIST GRUZOVOGO AVTOMOBILYA", ln=True, align="C")
        pdf.ln(10)
        
        lines