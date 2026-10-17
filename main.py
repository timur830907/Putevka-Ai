import io
import streamlit as st
from docx import Document
from fpdf import FPDF

st.set_page_config(
    page_title="Универсальный ИИ Путевой лист",
    page_icon="🚛",
    layout="wide"
)

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
        
        # Основные данные
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

        # Сохранение в буфер памяти для Streamlit
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def create_pdf(self, filename="waybill_empty.pdf"):
        """Генерация чистого шаблона путевого листа в формате PDF"""
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", size=12)
        
        # Заголовок
        pdf.cell(0, 10, txt="PUTEVOY LIST GRUZOVOGO AVTOMOBILYA", ln=True, align="C")
        pdf.ln(5)
        
        # Данные
        pdf.cell(0, 8, txt=f"Data: {self.data.get('date', '____________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Avtomobil: {self.data.get('car', '____________________')} | Gosnomer: {self.data.get('plate', '________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Voditel: {self.data.get('driver', '________________________________________')}", ln=True)
        pdf.ln(4)
        
        pdf.cell(0, 8, txt="--- Control ---", ln=True)
        pdf.cell(0, 8, txt=f"Medosmotr: {self.data.get('medical', '____________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Tehosmotr: {self.data.get('mechanic', '____________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Spidometr start: {self.data.get('km_start', '__________')} km | Toplivo: {self.data.get('fuel_start', '__________')} l", ln=True)
        pdf.ln(4)
        
        pdf.cell(0, 8, txt="--- Marshrut ---", ln=True)
        pdf.cell(0, 8, txt=f"Zakazchik: {self.data.get('customer', '________________________________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Pogruzka: {self.data.get('point_load', '________________________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Razgruzka: {self.data.get('point_unload', '________________________________')}", ln=True)
        pdf.cell(0, 8, txt=f"Gruz: {self.data.get('cargo', '________________________________')}", ln=True)
        pdf.ln(4)
        
        pdf.cell(0, 8, txt="--- Itogi ---", ln=True)
        pdf.cell(0, 8, txt=f"Spidometr end: {self.data.get('km_end', '__________')} km | Toplivo end: {self.data.get('fuel_end', '__________')} l", ln=True)

        buffer = io.BytesIO()
        pdf_bytes = pdf.output()
        buffer.write(pdf_bytes)
        buffer.seek(0)
        return buffer


# --- Streamlit Веб-интерфейс ---
st.title("🚛 Генератор путевых листов (AI)")
st.markdown("Заполните форму или оставьте поля пустыми для получения бланка с прочерками.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Основные данные")
    date = st.text_input("Дата", "")
    car = st.text_input("Автомобиль / Спецтехника", "")
    plate = st.text_input("Госномер", "")
    driver = st.text_input("Водитель (ФИО)", "")

    st.subheader("2. Контроль")
    medical = st.text_input("Отметка медосмотра", "")
    mechanic = st.text_input("Отметка механика", "")
    km_start = st.text_input("Спидометр (выезд, км)", "")
    fuel_start = st.text_input("Топливо (выезд, л)", "")

with col2:
    st.subheader("3. Маршрут и задание")
    customer = st.text_input("Заказчик", "")
    point_load = st.text_input("Пункт погрузки", "")
    point_unload = st.text_input("Пункт разгрузки", "")
    cargo = st.text_input("Характер груза", "")

    st.subheader("4. Итоги смены")
    km_end = st.text_input("Спидометр (возврат, км)", "")
    fuel_end = st.text_input("Топливо (возврат, л)", "")

# Сбор данных из формы
form_data = {
    "date": date if date else None,
    "car": car if car else None,
    "plate": plate if plate else None,
    "driver": driver if driver else None,
    "medical": medical if medical else None,
    "mechanic": mechanic if mechanic else None,
    "km_start": km_start if km_start else None,
    "fuel_start": fuel_start if fuel_start else None,
    "customer": customer if customer else None,
    "point_load": point_load if point_load else None,
    "point_unload": point_unload if point_unload else None,
    "cargo": cargo if cargo else None,
    "km_end": km_end if km_end else None,
    "fuel_end": fuel_end if fuel_end else None,
}

# Удаляем ключи со значением None, чтобы сработал `.get(..., '___')` в генераторе
clean_data = {k: v for k, v in form_data.items() if v is not None}

generator = WaybillGenerator(clean_data)

st.markdown("---")
st.subheader("📥 Скачать документ")

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    docx_buffer = generator.create_word()
    st.download_button(
        label="📄 Скачать в формате Word (DOCX)",
        data=docx_buffer,
        file_name="putevoy_list.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col_btn2:
    pdf_buffer = generator.create_pdf()
    st.download_button(
        label="🔴 Скачать в формате PDF",
        data=pdf_buffer,
        file_name="putevoy_list.pdf",
        mime="application/pdf"
    )