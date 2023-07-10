import streamlit as st
import os
import json

from docx import Document

from contants import General, Keys, Types
from PIL import Image

TEMPLATE_PATH = f'{os.getcwd()}/files/template.docx'


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
                value = self.__st.file_uploader(label)
        self.__certificate_update[key] = value

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
                if self.__st.form_submit_button('Create a new certificate'):
                    self.__st.info('Certificate info:')
                    self.__st.json(self.__certificate_update)
                    self.__st.session_state['data'] = self.__certificate_update
        else:
            with self.__st.sidebar:
                for k, o in self.__side_bar_fields.items():
                    self.__set_columns_by_type(key=k, obj=o)

    def __create_new_template(self):
        for p in self.__doc.paragraphs:
            p.paragraph_format.alignment = 2
            p.paragraph_format.left_indent = 1
            for field, field_value in self.__certificate_update.items():
                field_key = '{' + field + '}'
                if field_key in p.text:
                    inline = p.runs
                    # Loop added to work with runs (strings with same style)
                    for i in range(len(inline)):
                        if field_key in inline[i].text:
                            text = inline[i].text.replace(field_key, str(field_value))
                            inline[i].text = text
            print(p.text)
        doc_path = f"{os.getcwd()}/files/certificates/"
        doc_full_name = f"{doc_path}{self.__certificate_update['Date']}_{self.__certificate_update['Certificate number']}"

        self.__doc.save(f'{doc_full_name}.docx')
        self.__st.success(f"Your new certificate location: {doc_full_name}")

    def run(self):
        # step 1 - set sidebar params
        self.__set_ui_fields(sidebar=True)
        # step 2 - set center params and waiting for user action
        self.__set_ui_fields(sidebar=False)

        if self.__st.session_state.get('data') is not None and self.__st.button('Save'):
            self.__create_new_template()
            self.__st.session_state['data'] = None


if __name__ == "__main__":
    App().run()
