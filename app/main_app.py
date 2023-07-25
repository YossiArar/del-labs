import platform
import mammoth
import streamlit as st
import os
import json

from docx.shared import Inches, Cm

from app.read import add_float_picture

from docx import Document

from contants import General, Keys, Types
from PIL import Image

TEMPLATE_PATH = f'{os.getcwd()}/files/template.docx'
MAX_CHARS, PART1_LEN, LEFT_SIDE_MARGIN = 280, 268, 52
CM_DEFAULT_VALUES = [4.85, 4.85, 9.37, 0.75]
INCHES_DEFAULT_VALUES = [1.91, 1.91, 3.69, 0.3]
NOT_SPLIT_FIELDS = ["Date", "Certificate number", "Additional comments", 'Apprised Retail Price', "Currency"]
PUT_RIGHT_SIDE = False


class App:
    def __init__(self):
        with open(f"{os.getcwd()}/app/config.json", 'r') as f:
            self.__config = json.loads(f.read())
        self.__doc = Document(TEMPLATE_PATH)
        self.__st = st
        self.__st.set_page_config(page_title=General.APP_NAME, layout="wide")
        self.__st.title(General.APP_NAME)
        self.__side_bar_fields = self.__initiate_ui_field_options(sidebar=True)
        self.__center_fields = self.__initiate_ui_field_options(sidebar=False)
        self.__certificate_update = {}

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
            # with self.__st.form(key='SF'):
            with self.__st.sidebar:
                for k, o in self.__side_bar_fields.items():
                    self.__set_columns_by_type(key=k, obj=o)

    def __create_new_template(self):
        left_side_values, counter = [], 0
        for pi, p in enumerate(self.__doc.paragraphs, 1):
            p.paragraph_format.alignment = 2
            p.paragraph_format.left_indent = Inches(2.75)
            # p.paragraph_format.line_spacing = 0.85
            # p.paragraph_format.keep_together = True
            for field, field_value in self.__certificate_update.items():
                if field_value is not None and len(str(field_value)) > 0:
                    field_key, line_field_margin = '{' + field + '}', 0
                    # field_key, line_field_margin = field, 0
                    if field_key in p.text:
                        inline = p.runs
                        # Loop added to work with runs (strings with same style)
                        for i in range(len(inline)):
                            if field_key in inline[i].text:
                                if not PUT_RIGHT_SIDE:
                                    inline[i].text = inline[i].text.replace(field_key, str(field_value), 1)
                                else:
                                    inline[i].text = inline[i].text.replace(field_key, str(field_value))
                                # if len(text) > 0:
                                #     split_values = text.split('  ')
                                #     if len(split_values) > 0:
                                #         start_p, index = None, 0
                                #         # found start p value
                                #         for val in split_values:
                                #             index += 1
                                #             if str(field_value) in val:
                                #                 start_p = str(val[(0 if str(val)[0] != ' ' else 1):(
                                #                     -1 if str(val)[-1] == ' ' else len(val))])
                                #                 break
                                #         if start_p and PUT_RIGHT_SIDE:
                                #             end_p, field_value_len = '', len(field_key)
                                #             for val in split_values[index:]:
                                #                 index += 1
                                #                 if str(field_value) in val:
                                #                     end_p = str(val[(0 if str(val)[0] != ' ' else 1):(
                                #                         -1 if str(val)[-1] == ' ' else len(val))])
                                #                     break
                                #             space_margin, space_value = PART1_LEN - (len(start_p) + len(end_p)), ''
                                #             while space_margin != 0:
                                #                 space_value += ' '
                                #                 space_margin = space_margin - 1
                                #             part1 = f"{start_p}{space_value}"
                                #             print(f"p1 = {len(part1)}, {field_value_len}")
                                #             text = f"{part1}{end_p}"
                                #             if len(text) > MAX_CHARS:
                                #                 text = text.replace(space_value[:len(text) - MAX_CHARS], '', 1)
                                # if pi != 1 and not PUT_RIGHT_SIDE and len(text) > PART1_LEN / 2:
                                #     text = text[:int(PART1_LEN / 2)]
                                # update left side fields margin
                                if len(inline[i].text) != LEFT_SIDE_MARGIN and not PUT_RIGHT_SIDE \
                                        and field not in NOT_SPLIT_FIELDS:
                                    margin = LEFT_SIDE_MARGIN - len(inline[i].text)
                                    if margin >= 0:
                                        inline[i].text = inline[i].text.replace(str(field_value),
                                                                                f"{',' * margin}{str(field_value)}")
                                    else:
                                        inline[i].text = inline[i].text.replace(',', '', abs(margin))
                                    # field_text = inline[i].text.split(field_key)[0]
                                    # # inline[i].text = inline[i].text.split('  ')[0]
                                    # margin = LEFT_SIDE_MARGIN - len(field_text) + len(str(field_value))
                                    # # margin = LEFT_SIDE_MARGIN - (len(f"{field}{str(field_value)}"))
                                    # # margin = LEFT_SIDE_MARGIN - (len(field) - len(str(field_value)))
                                    # # inline[i].text = f"{field}{',' * margin}{field_value}"
                                    # if margin > 0:
                                    #     # inline[i].text = f"{inline[i].text.split(str(field_value))[0]}{',' * (margin + len(str(field_value)))}{field_value}"
                                    #     inline[i].text = f"{field_text}{',' * margin}{field_value}"
                                    # elif margin < 0:
                                    #     inline[i].text = f"{field_text[:margin]}{field_value}"
                                    # else:
                                    #     inline[i].text = f"{inline[i].text}{field_value}"
                                    left_side_values.append(inline[i].text)
                                p.text = inline[i].text  # .replace(',', '.')
                                counter += 1
                                # print(pi, p.text, len(p.text), len(p.text) - len(str(field_value)))
                                print(inline[i].text)
            if pi == 1 and self.__certificate_update.get('Image Path'):  # Inches
                add_float_picture(p=p, image_path_or_stream=self.__certificate_update.get('Image Path'),
                                  width=self.__certificate_update.get('image_data').get('width'),
                                  height=self.__certificate_update.get('image_data').get('height'),
                                  pos_x=self.__certificate_update.get('image_data').get('position_x'),
                                  pos_y=self.__certificate_update.get('image_data').get('position_y'),
                                  size_units=self.__certificate_update.get('image_data').get('size_units'))

        # last update for doc ph
        # print(left_side_values, counter)
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
