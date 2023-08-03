import platform
import mammoth
import streamlit as st
import os
import json

from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from fpdf import FPDF

from app.read import add_float_picture

from docx import Document

from contants import General, Keys, Types
from PIL import Image

TEMPLATE_PATH = f'{os.getcwd()}/files/template.docx'
MAX_CHARS, PART1_LEN, LEFT_SIDE_MARGIN = 280, 268, 52
CM_DEFAULT_VALUES = [4.85, 4.85, 9.37, 0.75]
INCHES_DEFAULT_VALUES = [1.91, 1.91, 3.69, 0.3]
PAGE_SIZE = 110, 210
BOLD_FIELDS = ["Certificate number", 'Apprised Retail Price', "Currency"]
NOT_SPLIT_FIELDS = ["Date", "Additional comments"] + BOLD_FIELDS
PUT_RIGHT_SIDE = False
FONT_NAME, FONT_SIZE = "Arial", 7


class App:
    def __init__(self):
        with open(f"{os.getcwd()}/app/config.json", 'r') as f:
            self.__config = json.loads(f.read())
        self.__doc = Document(TEMPLATE_PATH)
        self.__doc_bold_style = self.__doc.styles.add_style('BoldStyle', WD_STYLE_TYPE.CHARACTER)
        self.__st = st
        self.__st.set_page_config(page_title=General.APP_NAME, layout="wide")
        self.__st.title(General.APP_NAME)
        self.__side_bar_fields = self.__initiate_ui_field_options(sidebar=True)
        self.__center_fields = self.__initiate_ui_field_options(sidebar=False)
        self.__certificate_update = {}

    @property
    def __get_width_object(self):
        pdf = FPDF('p', 'mm', PAGE_SIZE)
        pdf.set_font(FONT_NAME, size=FONT_SIZE)
        pdf.add_page(orientation='l')
        return pdf

    def __set_columns_by_type(self, key, obj: dict):
        obj_type, value, desc = obj[Keys.TYPE], None, obj[Keys.DESC]
        label = key if obj_type == Types.CB else f"{key}{'' if not desc else f' - {desc}'}:"
        match obj_type:
            case Types.SELECT_BOX:
                value = self.__st.selectbox(label=label, options=obj[Keys.OPTIONS])
            case Types.FLOAT:
                value = self.__st.number_input(label=label, min_value=0.0)
            case Types.INT:
                value = self.__st.number_input(label=label, min_value=0)
            case Types.TXT:
                value = self.__st.text_input(label=label)
            case Types.DATE_INPUT:
                # value = str(self.__st.date_input(label=label))
                value = self.__st.date_input(label=label)
                day, month, year = f"0{value.day}" if value.day < 10 else value.day, f"0{value.month}" if value.month < 10 else value.day, value.year
                value = f"{day}-{month}-{year}"
            case Types.CB:
                value = self.__st.checkbox(label=label, value=False)
            case Types.IMAGE:
                value = self.__st.text_input(label).replace(
                    f"file:{'///' if 'Windows' in platform.platform() else '//'}", '')
                if value:
                    if os.path.exists(value):
                        if value:
                            image_data = {}
                            size_type, width, height, pos_x, pos_y, image_col = self.__st.columns(6)
                            with size_type:
                                size_ut = self.__st.radio(label='size_units', options=['CM', 'INCHES'])
                                image_data.setdefault('size_units', size_ut)
                                size_utv = CM_DEFAULT_VALUES if size_ut == 'CM' else INCHES_DEFAULT_VALUES
                                self.__st.info(f"{size_ut} default values: {size_utv}")
                            for col, label, val in zip([width, height, pos_x, pos_y],
                                                       ['width', 'height', 'position_x', 'position_y'],
                                                       size_utv):
                                with col:
                                    key_val = self.__st.number_input(label=f"{label}:", value=val)
                                    image_data.setdefault(label, key_val)
                            with image_col:
                                image = Image.open(value)
                                self.__st.image(image, caption='Your image selection')
                                if image_data:
                                    self.__certificate_update['image_data'] = image_data
                        else:
                            self.__st.error(f"The file path: {value} not exists.")
        if desc:
            value = f'{value} {desc}'
        self.__certificate_update[key] = value

    def __convert_docx_to_markdown(self, input_file, output_file=None):
        self.__doc = Document(input_file)
        paragraphs = [p.text for p in self.__doc.paragraphs]
        markdown_text = '\n'.join(paragraphs)
        output_file = output_file if output_file else input_file.replace('.docx', '_temp.md')
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown_text)
        return f, markdown_text

    def __initiate_ui_field_options(self, sidebar: bool):
        ui_fields = {}
        for key, values in self.__config.items():
            if values[Keys.SIDEBAR] == sidebar:
                ui_fields.setdefault(key, values)
        return ui_fields

    def __set_ui_fields(self, sidebar: bool):
        if not sidebar:
            with self.__st.form(key='CF'):
                for k, o in self.__center_fields.items():
                    self.__set_columns_by_type(key=k, obj=o)
                if self.__st.form_submit_button('Confirm certificate info'):
                    self.__st.info('Certificate items:')
                    self.__st.json(self.__certificate_update)
                    self.__st.session_state['data'] = self.__certificate_update
        else:
            with self.__st.sidebar:
                for k, o in self.__side_bar_fields.items():
                    self.__set_columns_by_type(key=k, obj=o)

    def __set_bold_txt(self, p):
        font = self.__doc_bold_style.font
        font.name = FONT_NAME
        font.size = Pt(FONT_SIZE)
        font.complex_script = True
        font.rtl = False
        p_txt = p.text
        p.clear()
        p.add_run(p_txt, style='BoldStyle').bold = True

    def __create_new_template(self):
        # set unique style
        doc_styles = self.__doc.styles['Normal']
        font = doc_styles.font
        font.name = FONT_NAME
        font.size = Pt(FONT_SIZE)
        update_width_values, cn_pi = {}, 0
        # step 1 - replace doc values
        for pi, p in enumerate(self.__doc.paragraphs, 1):
            p.paragraph_format.alignment = 2
            p.style.font.bold = False
            for field, field_value in self.__certificate_update.items():
                if field_value is not None and len(str(field_value)) > 0:
                    field_key, line_field_margin = '{' + field + '}', 0
                    if field_key in p.text:
                        inline = p.runs
                        for i in range(len(inline)):
                            if field_key in inline[i].text:
                                inline[i].text = inline[i].text.replace(field_key, str(field_value))
                                p.text = inline[i].text
                                if field not in NOT_SPLIT_FIELDS:
                                    update_width_values.setdefault(pi, p.text)
                                elif field in BOLD_FIELDS:
                                    self.__set_bold_txt(p)
                                    if field == BOLD_FIELDS[0]:
                                        cn_pi = pi
                        # print(len(p.text), p.text)

            # add new photo
            if pi == 1 and self.__certificate_update.get('Image Path'):  # Inches
                add_float_picture(p=p, image_path_or_stream=self.__certificate_update.get('Image Path'),
                                  width=self.__certificate_update.get('image_data').get('width'),
                                  height=self.__certificate_update.get('image_data').get('height'),
                                  pos_x=self.__certificate_update.get('image_data').get('position_x'),
                                  pos_y=self.__certificate_update.get('image_data').get('position_y'),
                                  size_units=self.__certificate_update.get('image_data').get('size_units'))

        # step 2 - get min width
        width_obj, min_width = self.__get_width_object, 3000
        for pi, p in enumerate(self.__doc.paragraphs, 1):
            if pi in update_width_values.keys() and p.text in update_width_values.get(pi):
                p_width = width_obj.get_string_width(p.text)
                if p_width < min_width:
                    min_width = p_width

        # step 3 - update paragraph width using FPDF package
        for pi, p in enumerate(self.__doc.paragraphs, 1):
            if pi in update_width_values.keys() and p.text in update_width_values.get(pi):
                p_width = width_obj.get_string_width(p.text)
                p_txt = p.text
                if p_width > min_width:
                    while int(p_width) != int(min_width):
                        p_txt = p_txt.replace(',', '', 1)
                        p_width = width_obj.get_string_width(p_txt)
                p.text = p_txt
            elif pi == cn_pi:
                cn_txt = p.text.replace(' ', '')
                cn_width = width_obj.get_string_width(cn_txt)
                if cn_width < min_width:
                    while int(cn_width) != int(min_width) and cn_width < min_width:
                        cn_txt = f" {cn_txt} "
                        cn_width = width_obj.get_string_width(cn_txt)
                p.text = cn_txt
                self.__set_bold_txt(p)
                p.paragraph_format.alignment = 2
                p.paragraph_format.left_indent = 0

        # last update for doc ph
        doc_path, doc_name = f"{os.getcwd()}/files/certificates/", f"{self.__certificate_update['Date']}_{self.__certificate_update['Certificate number']}.docx"
        doc_full_path = f"{doc_path}{doc_name}"
        self.__doc.save(doc_full_path)
        self.__convert_to_html(path=doc_full_path)

        # read data for download file
        with open(doc_full_path, 'rb') as d:
            data = d.read()
        download_button = self.__st.download_button(label='Download Certificate', data=data, file_name=doc_name)

    def __convert_to_html(self, path: str):
        with open(path, "rb") as docx_file:
            result = mammoth.convert_to_html(docx_file)
            html = result.value
        html_path = path.replace('docx', 'html')
        with open(html_path, "w") as html_file:
            html_file.write(html)
            html_file.close()
        self.__st.success("Your new certificate is ready to use")
        self.__st.write("Click here [link](%s) to print or download file" % html_path)

    def run(self):
        # step 1 - set sidebar params
        self.__set_ui_fields(sidebar=True)
        # step 2 - set center params and waiting for user action
        self.__set_ui_fields(sidebar=False)

        # if self.__st.session_state.get('data') is not None and self.__st.button('Save'):
        if self.__st.session_state.get('data') is not None and self.__st.button('Build'):
            self.__create_new_template()
            self.__st.session_state['data'] = None


if __name__ == "__main__":
    App().run()
