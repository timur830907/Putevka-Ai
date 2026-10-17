import io
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from fpdf import FPDF

st.set_page_config(
    page_title="ИИ Путевой лист РК",
    page_icon="🚛",
    layout="wide"
)

def set_cell_border(cell, **kwargs):
    """Настройка границ ячейки таблицы Word"""
    tcPr = cell._element.get_or_add_tcPr()
    tcBorders = parse_xml(r'<w:tcBorders %s/>' % nsdecls('w'))
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = parse_xml(f'<{tag} %s w:val="{edge_data.get("val", "single")}" w:sz="{edge_data.get("sz", "4")}" w:space="0" w:color="{edge_data.get("color", "auto")}"/>' % nsdecls('w'))
            tcBorders.append(element)
    tcPr.append(tcBorders)

class WaybillGenerator:
    def __init__(self, data: dict = None):
        self.data = data if data is not None else {}

    def create_word(self):
        doc = Document()

        # Поля страницы
        for section in doc.sections:
            section.top_margin = Cm(1.5)
            section.bottom_margin = Cm(1.5)
            section.left_margin = Cm(1.5)
            section.right_margin = Cm(1.5)

        # Заголовок бланка РК
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_kz = p.add_run("ЖҮК КӨЛІГІНІҢ ЖОЛ ҚАҒАЗЫ № _____\n")
        run_kz.bold = True
        run_kz.font.size = Pt(13)
        run_ru = p.add_run("ПУТЕВОЙ ЛИСТ ГРУЗОВОГО АВТОМОБИЛЯ № _____")
        run_ru.bold = True
        run_ru.font.size = Pt(13)

        # Дата
        p_date = doc.add_paragraph()
        p_date.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_date.add_run(f"Күні / Дата: {self.data.get('date', '«_____» ____________ 20___ ж./г.')}").font.size = Pt(10)

        # Шапка: Организация и Автомобиль
        header_table = doc.add_table(rows=3, cols=2)
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers_data = [
            [f"Ұйым / Организация:\n{self.data.get('organization', '___________________________________')}",
             f"Автомобиль маркасы / Марка авто:\n{self.data.get('car', '___________________________________')}"],
            [f"Жүргізуші / Водитель (ФИО):\n{self.data.get('driver', '___________________________________')}",
             f"Мемлекеттік нөмір / Гос. номер:\n{self.data.get('plate', '___________________________________')}"],
            [f"Куәлік № / Удостоверение №:\n{self.data.get('license', '___________________________________')}",
             f"Тіркеме / Прицеп гос. номер:\n{self.data.get('trailer', '___________________________________')}"]
        ]

        for i, row in enumerate(header_table.rows):
            for j, cell in enumerate(row.cells):
                cell.text = headers_data[i][j]
                cell.paragraphs[0].runs[0].font.size = Pt(9)
                set_cell_border(cell, top={"sz": 4}, bottom={"sz": 4}, left={"sz": 4}, right={"sz": 4})

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

        # Блок контроля и допуска
        doc.add_heading("1. Рұқсат берілуі / Медицинский и технический контроль", level=3)
        
        control_table = doc.add_table(rows=2, cols=2)
        control_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        c_data = [
            [f"Медициналық тексеру / Медосмотр:\n\nЖүргізуші денсаулығы жарамды / Водитель допущен.\nҚол / Подпись дәрігер: _________________",
             f"Техникалық тексеру / Техосмотр (Механик):\n\nАвтомобиль техникалық сау / Техника исправна.\nҚол / Подпись механик: _________________"],
            [f"Выезд (Шығу) спидометр: {self.data.get('km_start', '____________')} км\nЖанармай (Отпуск топлива): {self.data.get('fuel_start', '____________')} л",
             f"Возврат (Қайту) спидометр: {self.data.get('km_end', '____________')} км\nОстаток топлива: {self.data.get('fuel_end', '____________')} л"]
        ]

        for i, row in enumerate(control_table.rows):
            for j, cell in enumerate(row.cells):
                cell.text = c_data[i][j]
                cell.paragraphs[0].runs[0].font.size = Pt(9)
                set_cell_border(cell, top={"sz": 4}, bottom={"sz": 4}, left={"sz": 4}, right={"sz": 4})

        doc.add_paragraph().paragraph_format.space_after = Pt(4)

        # Таблица задания и маршрута
        doc.add_heading("2. Жүргізушінің тапсырмасы / Задание водителя и маршрут", level=3)

        route_table = doc.add_table(rows=5, cols=4)
        route_table.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = ["Тапсырыс беруші / Заказчик", "Тиеу орны / Пункт погрузки", "Түсіру орны / Пункт разгрузки", "Жүк түрі / Наим. груза"]
        for j, h in enumerate(headers):
            cell = route_table.rows[0].cells[j]
            cell.text = h
            cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].runs[0].font.size = Pt(9)
            set_cell_border(cell, top={"sz": 6}, bottom={"sz": 6}, left={"sz": 4}, right={"sz": 4})

        r_customer = self.data.get('customer', '')
        r_load = self.data.get('point_load', '')
        r_unload = self.data.get('point_unload', '')
        r_cargo = self.data.get('cargo', '')

        for row_idx in range(1, 5):
            row_cells = route_table.rows[row_idx].cells
            if row_idx == 1:
                row_cells[0].text = r_customer if r_customer else " "
                row_cells[1].text = r_load if r_load else " "
                row_cells[2].text = r_unload if r_unload else " "
                row_cells[3].text = r_cargo if r_cargo else " "
            else:
                for c in row_cells:
                    c.text = " "

            for cell in row_cells:
                if len(cell.paragraphs[0].runs) > 0:
                    cell.paragraphs[0].runs[0].font.size = Pt(9)
                set_cell_border(cell, top={"sz": 4}, bottom={"sz": 4}, left={"sz": 4}, right={"sz": 4})

        # Подписи
        doc.add_paragraph().paragraph_format.space_after = Pt(10)
        p_sign = doc.add_paragraph()
        p_sign.add_run("Диспетчер / Жауапты тұлға: ____________________            Жүргізуші / Водитель: ____________________").font.size = Pt(9)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    def create_pdf(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", size=10)
        
        pdf.cell(0, 8, txt="PUTEVOY LIST GRUZOVOGO AVTOMOBILYA (RK STANDARD)", ln=True, align="C")
        pdf.ln(4)
        
        pdf.cell(0, 6, txt=f"Data: {self.data.get('date', '____________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Organizaciya: {self.data.get('organization', '________________________________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Avtomobil: {self.data.get('car', '____________________')} | Gosnomer: {self.data.get('plate', '________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Voditel: {self.data.get('driver', '________________________________________')}", ln=True)
        pdf.ln(4)
        
        pdf.cell(0, 6, txt="--- CONTROL & MEDICAL ---", ln=True)
        pdf.cell(0, 6, txt="Medosmotr: Dopushten | Tehosmotr: Ispravan", ln=True)
        pdf.cell(0, 6, txt=f"Spidometr start: {self.data.get('km_start', '__________')} km | Toplivo start: {self.data.get('fuel_start', '__________')} l", ln=True)
        pdf.cell(0, 6, txt=f"Spidometr end: {self.data.get('km_end', '__________')} km | Toplivo end: {self.data.get('fuel_end', '__________')} l", ln=True)
        pdf.ln(4)
        
        pdf.cell(0, 6, txt="--- MARSHRUT / ZADANIE ---", ln=True)
        pdf.cell(0, 6, txt=f"Zakazchik: {self.data.get('customer', '________________________________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Pogruzka: {self.data.get('point_load', '________________________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Razgruzka: {self.data.get('point_unload', '________________________________')}", ln=True)
        pdf.cell(0, 6, txt=f"Gruz: {self.data.get('cargo', '________________________________')}", ln=True)
        pdf.ln(6)

        pdf.cell(0, 6, txt="Podpis: ____________________", ln=True)

        buffer = io.BytesIO()
        pdf_bytes = pdf.output()
        buffer.write(pdf_bytes)
        buffer.seek(0)
        return buffer


# --- Streamlit Веб-интерфейс ---
st.title("🚛 Официальный путевой лист грузового авто (РК)")
st.info("💡 **Совет для водителей:** Откройте эту ссылку со смартфона. Нажимайте на поля и надиктовывайте информацию через кнопку микрофона на клавиатуре телефона.")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Реквизиты и техника")
    organization = st.text_input("Наименование организации / ИП", "")
    date = st.text_input("Дата путевого листа", "")
    car = st.text_input("Марка автомобиля / спецтехники", "")
    plate = st.text_input("Гос. номер ТС", "")
    trailer = st.text_input("Гос. номер прицепа (при наличии)", "")
    driver = st.text_input("Водитель (ФИО)", "")
    license_no = st.text_input("Удостоверение водителя №", "")

with col2:
    st.subheader("2. Задание и показатели")
    customer = st.text_input("Заказчик", "")
    point_load = st.text_input("Пункт погрузки", "")
    point_unload = st.text_input("Пункт разгрузки", "")
    cargo = st.text_input("Наименование / характер груза", "")
    
    st.markdown("**Показатели спидометра и топлива:**")
    col_k1, col_k2 = st.columns(2)
    with col_k1:
        km_start = st.text_input("Спидометр выезд (км)", "")
        fuel_start = st.text_input("Топливо выезд (л)", "")
    with col_k2:
        km_end = st.text_input("Спидометр возврат (км)", "")
        fuel_end = st.text_input("Топливо возврат (л)", "")

form_data = {
    "organization": organization if organization else None,
    "date": date if date else None,
    "car": car if car else None,
    "plate": plate if plate else None,
    "trailer": trailer if trailer else None,
    "driver": driver if driver else None,
    "license": license_no if license_no else None,
    "customer": customer if customer else None,
    "point_load": point_load if point_load else None,
    "point_unload": point_unload if point_unload else None,
    "cargo": cargo if cargo else None,
    "km_start": km_start if km_start else None,
    "fuel_start": fuel_start if fuel_start else None,
    "km_end": km_end if km_end else None,
    "fuel_end": fuel_end if fuel_end else None,
}

clean_data = {k: v for k, v in form_data.items() if v is not None}
generator = WaybillGenerator(clean_data)

st.markdown("---")
st.subheader("📥 Скачать официальный бланк РК")

col_btn1, col_btn2 = st.columns(2)

with col_btn1:
    docx_buffer = generator.create_word()
    st.download_button(
        label="📄 Скачать Word бланком РК (.docx)",
        data=docx_buffer,
        file_name="putevoy_list_RK.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

with col_btn2:
    pdf_buffer = generator.create_pdf()
    st.download_button(
        label="🔴 Скачать PDF бланком РК (.pdf)",
        data=pdf_buffer,
        file_name="putevoy_list_RK.pdf",
        mime="application/pdf"
    )