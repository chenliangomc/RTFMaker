"""
core.py is part of RTFMaker, a simple RTF document generation package

Copyright (C) 2019  Liang Chen

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU Affero General Public License for more details.

You should have received a copy of the GNU Affero General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

from __future__ import absolute_import

class RTFDocument(object):
    """RTF document container"""

    KEY_TYPE = 'type'
    KEY_VALUE = 'value'
    KEY_FONT = 'font'
    KEY_STYLE = 'style'
    DEFAULT_LANGUAGE = 'EnglishUS'
    DEFAULT_FONT_NAME = 'Arial'
    DEFAULT_FONT_SIZE = '9'
    MODIFIER_REGULAR = 'Regular'
    MODIFIER_BOLD = 'Bold'
    MODIFIER_ITALIC = 'Italic'

    def __init__(self, **kwargs):
        self._element_cache = list()
        self._default_p_style = None
        self._doc = None

    def append(self, element):
        """
        add new element to the document

        @param element (dict)
        """
        self._element_cache.append(element)
        return None

    def _get_font_style(self, data, **kwargs):
        """generate font and text style object

        @param data basic text style information (dict)

        @return (string, font_obj, text_style_obj)
        """

    def _parse_css_font(self, css_font_def, **kwargs):
        """convert CSS font directives to internal representation

        @param css_font_def (string)
        """
        ret = dict()
        rules = css_font_def.strip().split(';')
        rules = [ i.strip() for i in rules if len(i.strip()) ]
        attrs = dict([ i.split(':',1) for i in rules ])

        ret['size'] = int(attrs.get('font-size', self.DEFAULT_FONT_SIZE).replace('pt', ''))
        ret['font'] = attrs.get('font-family').split(',')[0]
        ret['modifier'] = None
        if attrs.get('font-weight'):
            ret['modifier'] = self.MODIFIER_BOLD
        if attrs.get('font-style'):
            ret['modifier'] = self.MODIFIER_ITALIC
        return ret

    def _collect_styles(self, **kwargs):
        """get all the registered styles

        @rtype `PyRTF.Elements.StyleSheet`
        """
        from PyRTF.Elements import StyleSheet
        from PyRTF.Styles import TextStyle, ParagraphStyle
        from PyRTF.PropertySets import Font, TextPropertySet
        from .utils import StyleSet

        font_set = StyleSet(Font)
        t_style_set = StyleSet(TextStyle)
        p_style_set = StyleSet(ParagraphStyle)

        _default_font_ts = self._get_font_style(
            data={
                'font': self.DEFAULT_FONT_NAME,
                'size': int(self.DEFAULT_FONT_SIZE),
                'modifier': self.MODIFIER_REGULAR,
            },
            **kwargs
        )
        f_arial = Font('Arial', 'swiss', 0, 2, '020b0604020202020204')
        ts_arial_9pt_regular = TextStyle(
            TextPropertySet(
                font=f_arial,
                size=2*9,
                bold=False,
                italic=False,
                underline=False,
            ),
            name='Arial 9pt Regular'
        )
        ps_normal = ParagraphStyle('Normal', ts_arial_9pt_regular)
        self._default_p_style = ps_normal

        # insert the default one at the beginning;
        font_set.add(f_arial)
        # rvalue;
        _doc_style = StyleSheet(fonts=font_set)
        _doc_style.TextStyle.append(ts_arial_9pt_regular)
        _doc_style.ParagraphStyles.append(ps_normal)
        return _doc_style

    def _collect_elements(self, **kwargs):
        """get all the elements

        @rtype `PyRTF.document.section.Section`
        """
        from PyRTF.document.section import Section
        from PyRTF.document.paragraph import Paragraph
        from PyRTF.document.paragraph import Table

        # go through element list and add to section;
        ret = Section()
        for a_element in self._element_cache:
            e_type = a_element.get(self.KEY_TYPE, None)
            e_ctx = a_element.get(self.KEY_VALUE, '')
            e_style = a_element.get(self.KEY_STYLE, None)
            # use captured styles to create document element;
            if e_type == 'paragraph':
                element_obj = Paragraph()
                if e_style is not None:
                    element_obj = Paragraph(e_style)
                element_obj.append(e_ctx)
                ret.append(element_obj)
            elif e_type == 'table':
                element_obj = Table()
                ret.append(element_obj)
            else:
                element_obj = None
        return ret

    def _to_rtf(self, **kwargs):
        """convert internal representation of document structure into RTF stream

        @rtype `PyRTF.Elements.Document`
        """
        from PyRTF.Constants import Languages
        from PyRTF.Elements import Document

        # create document object, and capture all the styles;
        _doc = Document(
            style_sheet=self._collect_styles(**kwargs),
            default_language=getattr(Languages, self.DEFAULT_LANGUAGE),
        )

        # parse element objects and add to document;
        _sect = self._collect_elements(**kwargs)
        _doc.Sections.append(_sect)

        return _doc

    def _write(self, file, **kwargs):
        """dump the full document into the file"""
        if not self._doc:
            self._doc = self._to_rtf(**kwargs)

        return self._doc.write(file, **kwargs)

    def to_string(self, **kwargs):
        """
        return the string stream of the full document

        @param strip_newline:
        @type strip_newline: boolean

        return `string`
        """

        from StringIO import StringIO
        cache = StringIO()
        self._write(cache, **kwargs)
        ret = cache.getvalue()
        return ret

    def __repr__(self):
        ret = "<RTF document of {ec} element(s) at {addr}>".format(
            ec=len(self._element_cache),
            addr="0x%x"%(id(self)),
        )
        return ret


#---eof---#
