import platform

import streamlit as st
import os
import json

from docx.shared import Inches, Cm

from app.read import add_float_picture

from docx import Document

from contants import General, Keys, Types
from PIL import Image

TEMPLATE_PATH = f'{os.getcwd()}/files/template.docx'
MAX_CHARS, PART1_LEN = 280, 268


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
                value = str(self.__st.date_input(label=label))
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
                                size_utv = [4.85, 4.85, 9.37, 0.25] if size_ut == 'CM' else [1.91, 1.91, 3.69, 0.1]
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
        ignore_right_side, left_size_margin, max_left_margin = self.__certificate_update.get(
            'Ignore right side data'), {}, 0
        for pi, p in enumerate(self.__doc.paragraphs, 1):
            p.paragraph_format.alignment = 2
            p.paragraph_format.left_indent = 1
            p.paragraph_format.line_spacing = 0.85
            p.paragraph_format.keep_together = True
            for field, field_value in self.__certificate_update.items():
                if field_value is not None and len(str(field_value)) > 0:
                    field_key, line_field_margin = '{' + field + '}', 0
                    # field_key, line_field_margin = field, 0
                    if field_key in p.text:
                        inline = p.runs
                        # Loop added to work with runs (strings with same style)
                        for i in range(len(inline)):
                            if field_key in inline[i].text:
                                if ignore_right_side:
                                    text = inline[i].text.replace(field_key, str(field_value), 1)
                                else:
                                    text = inline[i].text.replace(field_key, str(field_value))
                                if len(text) > 0:
                                    split_values = text.split('  ')
                                    if len(split_values) > 0:
                                        start_p, index = None, 0
                                        # found start p value
                                        for val in split_values:
                                            index += 1
                                            if str(field_value) in val:
                                                start_p = str(val[(0 if str(val)[0] != ' ' else 1):(
                                                    -1 if str(val)[-1] == ' ' else len(val))])
                                                if '……' in start_p:
                                                    left_size_margin[field_key[1:-1]] = [len(start_p), start_p]
                                                    if len(start_p) > max_left_margin:
                                                        max_left_margin = len(start_p)
                                                break
                                        if start_p and not ignore_right_side:
                                            end_p, field_value_len = '', len(field_key)
                                            for val in split_values[index:]:
                                                index += 1
                                                if str(field_value) in val:
                                                    end_p = str(val[(0 if str(val)[0] != ' ' else 1):(
                                                        -1 if str(val)[-1] == ' ' else len(val))])
                                                    break
                                            space_margin, space_value = PART1_LEN - (len(start_p) + len(end_p)), ''
                                            while space_margin != 0:
                                                space_value += ' '
                                                space_margin = space_margin - 1
                                            part1 = f"{start_p}{space_value}"
                                            print(f"p1 = {len(part1)}, {field_value_len}")
                                            text = f"{part1}{end_p}"
                                            if len(text) > MAX_CHARS:
                                                text = text.replace(space_value[:len(text) - MAX_CHARS], '', 1)
                                if pi != 1 and ignore_right_side and len(text) > PART1_LEN / 2:
                                    text = text[:int(PART1_LEN / 2)]
                                inline[i].text = text
                                print(text, len(text))
            if pi == 1 and self.__certificate_update.get('Image Path'):  # Inches
                add_float_picture(p=p, image_path_or_stream=self.__certificate_update.get('Image Path'),
                                  width=self.__certificate_update.get('image_data').get('width'),
                                  height=self.__certificate_update.get('image_data').get('height'),
                                  pos_x=self.__certificate_update.get('image_data').get('position_x'),
                                  pos_y=self.__certificate_update.get('image_data').get('position_y'),
                                  size_units=self.__certificate_update.get('image_data').get('size_units'))
        # last update for doc ph
        new_p_data = {}
        for key, values in left_size_margin.items():
            if values[0] < max_left_margin:
                space_margin = max_left_margin - values[0]
                split_p, space_value = values[1].split(key), ''
                while space_margin > 0:
                    space_value += '…'
                    space_margin = space_margin - 1
                new_p_value = f"{key}{space_value}{split_p[-1]}"
                new_p_data[key] = [values[1], new_p_value, len(new_p_value)]
                for pi, p in enumerate(self.__doc.paragraphs, 1):
                    if new_p_data[key][0] in p.text:
                        inline = p.runs
                        for i in range(len(inline)):
                            if new_p_data[key][0] in inline[i].text:
                                text = inline[i].text.replace(new_p_data[key][0], new_p_data[key][1])
                                inline[i].text = text
                                print(text, len(text))
                                break

        doc_path, doc_name = f"{os.getcwd()}/files/certificates/", f"{self.__certificate_update['Date']}_{self.__certificate_update['Certificate number']}.docx"
        doc_full_path = f"{doc_path}{doc_name}"
        self.__doc.save(doc_full_path)

        with open(doc_full_path, 'rb') as d:
            data = d.read()
        # # check option for present word document
        # markdown_text, doc = self.__convert_docx_to_markdown(input_file=doc_full_path), self.__doc
        # markdown_file = doc_full_path.replace('.docx', '_temp.md')
        # with open(markdown_file, 'w', encoding='utf-8') as f:
        #     f.write(markdown.markdown(str(data)))
        # st.markdown(doc.paragraphs, unsafe_allow_html=False)

        download_button = self.__st.download_button(label='Download Certificate', data=data, file_name=doc_name)
        if download_button:
            self.__st.success("Your new certificate is ready to use")

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
