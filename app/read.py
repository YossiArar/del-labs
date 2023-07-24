from docx.shared import Inches, Cm

# -*- coding: utf-8 -*-

'''
Implement floating image based on python-docx.

- Text wrapping style: BEHIND TEXT <wp:anchor behindDoc="1">
- Picture position: top-left corner of PAGE `<wp:positionH relativeFrom="page">`.

Create a docx sample (Layout | Positions | More Layout Options) and explore the
source xml (Open as a zip | word | document.xml) to implement other text wrapping
styles and position modes per `CT_Anchor._anchor_xml()`.
'''

from docx.oxml import parse_xml, register_element_cls
from docx.oxml.ns import nsdecls
from docx.oxml.shape import CT_Picture
from docx.oxml.xmlchemy import BaseOxmlElement, OneAndOnlyOne


# refer to docx.oxml.shape.CT_Inline
class CT_Anchor(BaseOxmlElement):
    """
    ``<w:anchor>`` element, container for a floating image.
    """
    extent = OneAndOnlyOne('wp:extent')
    docPr = OneAndOnlyOne('wp:docPr')
    graphic = OneAndOnlyOne('a:graphic')

    @classmethod
    def new(cls, cx, cy, shape_id, pic, pos_x, pos_y):
        """
        Return a new ``<wp:anchor>`` element populated with the values passed
        as parameters.
        """
        anchor = parse_xml(cls._anchor_xml(pos_x, pos_y))
        anchor.extent.cx = cx
        anchor.extent.cy = cy
        anchor.docPr.id = shape_id
        anchor.docPr.name = 'Picture %d' % shape_id
        anchor.graphic.graphicData.uri = (
            'http://schemas.openxmlformats.org/drawingml/2006/picture'
        )
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def new_pic_anchor(cls, shape_id, rId, filename, cx, cy, pos_x, pos_y):
        """
        Return a new `wp:anchor` element containing the `pic:pic` element
        specified by the argument values.
        """
        pic_id = 0  # Word doesn't seem to use this, but does not omit it
        pic = CT_Picture.new(pic_id, filename, rId, cx, cy)
        anchor = cls.new(cx, cy, shape_id, pic, pos_x, pos_y)
        anchor.graphic.graphicData._insert_pic(pic)
        return anchor

    @classmethod
    def _anchor_xml(cls, pos_x, pos_y):
        return (
                '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="0" \n'
                '           behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1" \n'
                '           %s>\n'
                '  <wp:simplePos x="0" y="0"/>\n'
                '  <wp:positionH relativeFrom="page">\n'
                '    <wp:posOffset>%d</wp:posOffset>\n'
                '  </wp:positionH>\n'
                '  <wp:positionV relativeFrom="page">\n'
                '    <wp:posOffset>%d</wp:posOffset>\n'
                '  </wp:positionV>\n'
                '  <wp:extent cx="914400" cy="914400"/>\n'
                '  <wp:wrapNone/>\n'
                '  <wp:docPr id="666" name="unnamed"/>\n'
                '  <wp:cNvGraphicFramePr>\n'
                '    <a:graphicFrameLocks noChangeAspect="1"/>\n'
                '  </wp:cNvGraphicFramePr>\n'
                '  <a:graphic>\n'
                '    <a:graphicData uri="URI not set"/>\n'
                '  </a:graphic>\n'
                '</wp:anchor>' % (nsdecls('wp', 'a', 'pic', 'r'), int(pos_x), int(pos_y))
        )


# refer to docx.parts.story.BaseStoryPart.new_pic_inline
def new_pic_anchor(part, image_descriptor, width, height, pos_x, pos_y):
    """Return a newly-created `w:anchor` element.

    The element contains the image specified by *image_descriptor* and is scaled
    based on the values of *width* and *height*.
    """
    rId, image = part.get_or_add_image(image_descriptor)
    cx, cy = image.scaled_dimensions(width, height)
    shape_id, filename = part.next_id, image.filename
    return CT_Anchor.new_pic_anchor(shape_id, rId, filename, cx, cy, pos_x, pos_y)


# refer to docx.text.run.add_picture
# def add_float_picture(p, image_path_or_stream, width: Inches = Inches(1.91), height: Inches = Inches(1.91),
#                       pos_x: Inches = Inches(3.69),  # Pt(300),
#                       pos_y: Inches = Inches(0.1)):
def add_float_picture(p, image_path_or_stream, width=Inches(1.91), height=Inches(1.91),
                      pos_x=Inches(3.69),  # Pt(300),
                      pos_y=Inches(0.1), size_units: str = 'CM'):
    """Add float picture at fixed position `pos_x` and `pos_y` to the top-left point of page.
    """
    if size_units == 'CM':
        width, height, pos_x, pos_y = Cm(width), Cm(height), Cm(pos_x), Cm(pos_y)
    else:  # inches
        width, height, pos_x, pos_y = Inches(width), Inches(height), Inches(pos_x), Inches(pos_y)
    run = p.add_run()
    anchor = new_pic_anchor(run.part, image_path_or_stream, width, height, pos_x, pos_y)
    run._r.add_drawing(anchor)


# refer to docx.oxml.shape.__init__.py
register_element_cls('wp:anchor', CT_Anchor)



# max_size = 0
# pdf_file = PyPDF2.PdfReader(open(TEMPLATE_PATH, "rb"))
#
# # Count number of pages in our pdf file
# number_of_pages = len(pdf_file.pages)
# # print number of pages in the pdf file
# print("Number of pages in this pdf: " + str(number_of_pages))
#
# # Read first page
# page = pdf_file.pages[0]
#
# # print entire text of first page of the pdf
# text = page.extract_text().replace(' ', '')
# # print(f"before: {text}")
#
# for field, field_value in FIELDS_DATA.items():
#     field_key = '{' + field.replace(' ', '') + '}'
#     if field not in IGNORE_FIELDS and field_key in text and len(str(field_value)) > 0:
#         # replace text in pdf
#         field_margin = len(field_key) - len(str(field_value))
#         text = text.replace(field_key, str(field_value))
#         # update title
#         text = text.replace(field_key[1:-1], field)
#         # print(text)
# pdf_lines = text.splitlines()
# for pl in pdf_lines:
#     if len(pl) > max_size:
#         max_size = len(pl)
# print(f"\n\nafter: {text}")
# # # # page_size = (page.mediabox.width, page.mediabox.height)
# #
# pdf_text = ''
# for field, field_value in FIELDS_DATA.items():
#     if field not in IGNORE_FIELDS and len(str(field_value)) > 0:
#         field_margin, value_margin = len(field), len(str(field_value))
#         values_margin = field_margin - value_margin + value_margin
#         margin = max_size - values_margin  # (field_margin + value_margin)
#         margin = '.' * margin
#         row_text = f"{field}{margin}{field_value}"
#         # w = len(row_text) + 6
#         # new_p = (210 - w) / 2
#         # value_whitespace_margin = int(str(field_value).count(' '))
#         # row_text = f"{row_text[:-value_margin if values_margin > 0 else -value_margin]}{field_value}"
#         # row_text = f"{row_text[:len(row_text) if value_whitespace_margin == 0 else -value_whitespace_margin]}{field_value}"
#         # row_text = textwrap.fill(text=row_text, width=max_size, tabsize=0)
#         # row_text = textwrap.shorten(row_text, width=len(row_text))
#         # if len(row_text) == max_size:
#         pdf_text += f"{row_text}\n"
#         print(len(row_text), field_margin, value_margin, str(field_value))
# # pdf_text = textwrap.indent('\n', pdf_text)
# # pdf_text = textwrap.fill(pdf_text)
# print(f"\n\nafter update:\n"
#       f"{pdf_text}")