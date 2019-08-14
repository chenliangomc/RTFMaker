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

    FORMAT_FONT_FULL_NAME = '{name} {size}pt {modifier}'
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

        @note font definition data is extracted from `PyRTF` package

        @param data basic text style information (dict)

        @return (string, font_obj, text_style_obj)
        """
        font_short_name = data.get('font', self.DEFAULT_FONT_NAME)
        font_size = data.get('size', int(self.DEFAULT_FONT_SIZE))
        font_decor = data.get('modifier', self.MODIFIER_REGULAR)
        #
        _FONT_ARG_HUB = {
            'Arial': ('swiss', 0, 2, '020b0604020202020204'),
            #self.DEFAULT_FONT_NAME: ('swiss', 0, 2, '020b0604020202020204'),
            'Arial Black': ('swiss', 0, 2, '020b0a04020102020204'),
            'Courier New': ('modern', 0, 1, '02070309020205020404'),
            #('Bitstream Vera Sans Mono', 'modern', 0, 1, '020b0609030804020204'),
            #('Monotype Corsiva', 'script', 0, 2, '03010101010201010101'),
            #('Tahoma', 'swiss', 0, 2, '020b0604030504040204'),
            #('Trebuchet MS', 'swiss', 0, 2, '020b0603020202020204'),
        }
        #
        font_obj = None
        txt_style_obj = None
        font_full_name = self.FORMAT_FONT_FULL_NAME.format(
            name=font_short_name,
            size=font_size,
            modifier=font_decor
        )
        #
        from PyRTF.Styles import TextStyle
        from PyRTF.PropertySets import Font
        from PyRTF.PropertySets import TextPropertySet
        font_args = _FONT_ARG_HUB.get(
            font_short_name,
            _FONT_ARG_HUB[self.DEFAULT_FONT_NAME]
        )
        font_obj = Font(font_short_name, *font_args)
        txt_style_obj = TextStyle(
            TextPropertySet(
                font=font_obj,
                size=2*font_size,
                bold=True if font_decor.find(self.MODIFIER_BOLD) > -1 else False,
                italic=True if font_decor.find(self.MODIFIER_ITALIC) > -1 else False,
                underline=False,
            ),
            name=font_full_name
        )
        return (font_full_name, font_obj, txt_style_obj)

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
        from PyRTF.PropertySets import Font
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
        f_arial = _default_font_ts[1]
        ts_arial_9pt_regular = _default_font_ts[2]
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

        def _inject_blankline(element, **kwargs):
            line = None
            line_text = kwargs.get('alt.line.text', '')
            if element.get('append_newline', False):
                line = Paragraph(self._default_p_style)
                line.append(line_text)
            return line

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
            #
            trailing = _inject_blankline(a_element, **kwargs)
            if trailing:
                ret.append(trailing)
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
