import os
from docx import Document

REPLACE_OPTIONS = '..', '…', ' .', ':'  # , "\""
SPLIT_OPTIONS = '.', ':'
IGNORE_ROWS = []  # ['Additional Comments', 'This Jewelry has been tested and verified by a GG']
CERTIFICATE = 'Certificate'

DATA = {'Date': '2023-05-06', 'Type of Certificate': 'Big', 'Certificate number': '', 'certificate validation': False,
        'Additional Comments': '', 'Apprised Retail Price': 0.0, 'Currency': '$', 'Image Path': None,
        'Type of jewelry': 'Ring', 'Metal Type': 'Platinum\xa0(328)', 'Total Metal Wight': 0.0,
        'Stone Type': 'Natural Diamond', 'Stone Total Weight': 0.0, 'Amount of Stones': 0, 'Color Grade': 'D',
        'Clarity Grade': 'IF', 'Cut Grade': 'Excellent'}


class Word:
    def __init__(self, doc_path: str):
        if os.path.exists(path=doc_path):
            self.__main_document = Document(docx=doc_path)
        else:
            raise print(f"{doc_path} is not exists")
        self.__metadata, self.__new_document = {}, None
        self.__initiate_document_object()

    @property
    def metadata(self):
        return self.__metadata

    @staticmethod
    def __extract_value(values: list):
        value = ''
        iterations = 0
        # first check
        for v in values:
            if v:
                value += f" {v}"
            if value:
                iterations += 1
            if iterations > 3:
                break
        # second check
        if value:
            value = value.replace('  ', ' ')
            value = value if value[0] != ' ' else value[1:]
            value = value if value[-1] != ' ' else value[:-1]
        return value

    def __initiate_document_object(self):
        self.__new_document, paragraphs = self.__main_document, self.__main_document.paragraphs
        update_values = {}

        for i, p in enumerate(paragraphs, 1):
            if p.text:
                # ignore specific chars
                temp_txt = p.text
                for rp in REPLACE_OPTIONS:
                    temp_txt = temp_txt.replace(rp, ' ')
                txt_split = temp_txt.replace('  ', '/').split('/')
                if len(txt_split) > 2:
                    key = None if not txt_split[0] and i > 1 else CERTIFICATE if i < 2 else txt_split[0]
                    if key and key not in IGNORE_ROWS:
                        # check all values
                        value = self.__extract_value(values=txt_split[1:])
                        if value not in update_values.keys():
                            update_values.setdefault(value, key)
                        self.__metadata.setdefault(key if key is not None else str(value), value)
            paragraphs[i - 1] = p
        # update paragraphs
        for i, p in enumerate(paragraphs):
            self.__new_document.paragraphs.insert(i, paragraphs[i])

    def display_template(self, save_path=None):
        for p in self.__new_document.paragraphs:
            print(p.text)

        if save_path:
            save_path = save_path if '.docx' in save_path else f"{save_path}.docx"
            self.__main_document.save(save_path)

    # paragraphs = self.__doc.paragraphs
    # for i, p in enumerate(self.__doc.paragraphs, 1):
    #     if p.text:
    #         # ignore specific chars
    #         for rp in REPLACE_OPTIONS:
    #             p.text = p.text.replace(rp, ' ')
    #         txt_split = p.text.replace('  ', '/').split('/')
    #         if len(txt_split) > 2:
    #             key = None if not txt_split[0] else txt_split[0]
    #             if key:
    #                 # check all values
    #                 value = ''
    #                 for v in txt_split[1:]:
    #                     if v:
    #                         value += v
    #                 metadata.setdefault(key if key is not None else str(value),
    #                                     value if value[0] != ' ' else value[1:])
    #             print(f"{i}. {p.text}\n")  # \n{metadata}\n")
    #             paragraphs[i - 1] = p


if __name__ == '__main__':
    path = '../files/template.docx'
    word_obj = Word(doc_path=path)
    print(word_obj.metadata)
    # word_obj.display_template(save_path='new_template')
