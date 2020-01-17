"""
core.py is part of RTFMaker, a simple RTF document generation package

Copyright (C) 2019, 2020  Liang Chen

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
    KEY_ADD_NEWLINE = 'append_newline'
    DEFAULT_LANGUAGE = 'EnglishUS'
    DEFAULT_FONT_NAME = 'Arial'
    DEFAULT_FONT_SIZE = '9'
    DEFAULT_PSTYLE_NAME = 'Normal'
    DEFAULT_LIST_STYLE_NAME = 'List 1'
    ELEMENT_PARAGRAPH = 'paragraph'
    ELEMENT_TABLE = 'table'
    ELEMENT_LIST = 'list'
    ELEMENT_PARTIAL = 'partial'
    MODIFIER_REGULAR = 'Regular'
    MODIFIER_BOLD = 'Bold'
    MODIFIER_ITALIC = 'Italic'

    DEFAULT_EM_WIDTH = 90  # (PyRTF.PropertySets.TabPropertySet.DEFAULT_WIDTH/8)
    DEFAULT_LIST_INDENT = 4
    DEFAULT_LIST_HANGING = 2

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
        additional_font_mapping = kwargs.get('alt.font.map', None)
        if isinstance(additional_font_mapping, dict):
            _FONT_ARG_HUB.update(additional_font_mapping)
        #
        if font_decor is None:
            font_decor = self.MODIFIER_REGULAR
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
        ret['font'] = attrs.get('font-family', self.DEFAULT_FONT_NAME).split(',')[0]
        ret['modifier'] = self.MODIFIER_REGULAR
        css_font_weight = attrs.get('font-weight')
        if css_font_weight is not None and css_font_weight.lower().find('normal') == -1:
            ret['modifier'] = self.MODIFIER_BOLD
        css_font_style = attrs.get('font-style')
        if css_font_style is not None and css_font_style.lower().find('italic') > -1:
            ret['modifier'] = self.MODIFIER_ITALIC
        return ret

    def _get_bold_style_name(self, name, **kwargs):
        ret = name
        style_cache = getattr(self, '_style_cache')
        assert style_cache is not None, 'style cache unavailable'

        if name.find(self.MODIFIER_BOLD) > -1 :
            pass
        else:
            if name == self.DEFAULT_PSTYLE_NAME:
                ret = 'ps_{ts}'.format(ts=self._default_p_style.TextStyle.name).replace(self.MODIFIER_REGULAR, self.MODIFIER_BOLD)
            else:
                ret = name.replace(self.MODIFIER_REGULAR, self.MODIFIER_BOLD)
        # append the bold style to stylesheet if it is not included already;
        new_pstyle = style_cache.ParagraphStyles.get_by_name(ret)
        if new_pstyle is None:
            raise NotImplementedError('TODO: construct the new style and append')
            #from PyRTF.Styles import ParagraphStyle
        return ret

    def _collect_styles(self, **kwargs):
        """get all the registered styles

        @rtype `PyRTF.Elements.StyleSheet`
        """
        from PyRTF.Elements import StyleSheet
        from PyRTF.Styles import TextStyle, ParagraphStyle
        from PyRTF.PropertySets import Font
        from PyRTF.PropertySets import ParagraphPropertySet
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
        ps_normal = ParagraphStyle(self.DEFAULT_PSTYLE_NAME, ts_arial_9pt_regular)
        self._default_p_style = ps_normal

        # insert the default one at the beginning;
        font_set.add(f_arial)
        t_style_set.append(ts_arial_9pt_regular)
        p_style_set.append(ps_normal)

        doc_has_list = False
        # then go through element list to collect all other styles;
        for a_element in self._element_cache:
            # extract style information;
            e_font = a_element.get(self.KEY_FONT, None)
            if a_element.get(self.KEY_TYPE, None) == self.ELEMENT_LIST:
                doc_has_list = True
            # try to match any registered style;
            if e_font:
                font_arg = self._parse_css_font(e_font, **kwargs)
                new_font_obj = self._get_font_style(data=font_arg, **kwargs)
                font_set.add(new_font_obj[1])
                t_style_set.add(new_font_obj[2])
                p_style = ParagraphStyle(
                    'ps_{ts}'.format(ts=new_font_obj[0]),
                    new_font_obj[2]
                )
                p_style_set.add(p_style)
            else:
                p_style = ps_normal
            # replace raw style info with internal style cache reference;
            a_element[self.KEY_STYLE] = p_style.name
            if a_element.get(self.KEY_TYPE, None) == self.ELEMENT_PARTIAL:
                for a_sub in a_element[ self.KEY_VALUE ]:
                    if a_sub is None:
                        continue
                    sub_font = a_sub.get(self.KEY_FONT, None)
                    if sub_font:
                        sub_font_arg = self._parse_css_font(sub_font, **kwargs)
                        sub_new_font_obj = self._get_font_style(data=sub_font_arg, **kwargs)
                        font_set.add(sub_new_font_obj[1])
                        t_style_set.add(sub_new_font_obj[2])
                        sub_p_style = ParagraphStyle(
                            'ps_{ts}'.format(ts=sub_new_font_obj[0]),
                            sub_new_font_obj[2]
                        )
                        p_style_set.add(sub_p_style)
                    else:
                        sub_p_style = ps_normal
                    a_sub[self.KEY_STYLE] = sub_p_style.name

        # put in list style when needed;
        if doc_has_list:
            list_item_indent = self.DEFAULT_EM_WIDTH * self.DEFAULT_LIST_INDENT
            hanging_indent = -(self.DEFAULT_EM_WIDTH * self.DEFAULT_LIST_HANGING)
            ps_for_list_item = ParagraphStyle(
                self.DEFAULT_LIST_STYLE_NAME,
                ts_arial_9pt_regular,
                ParagraphPropertySet(
                    space_before=60,
                    space_after=60,
                    first_line_indent=hanging_indent,
                    left_indent=list_item_indent
                )
            )
            p_style_set.append(ps_for_list_item)

        # rvalue;
        _doc_style = StyleSheet(fonts=font_set)
        # overwrite default values;
        _doc_style.TextStyle = t_style_set
        _doc_style.ParagraphStyles = p_style_set
        return _doc_style

    def _collect_elements(self, **kwargs):
        """get all the elements

        @rtype `PyRTF.document.section.Section`
        """
        from PyRTF.document.section import Section
        from .utils import RPar, RTable, RList

        def _inject_blankline(element, **kwargs):
            line = None
            line_text = kwargs.get('alt.line.text', '')
            if element.get(self.KEY_ADD_NEWLINE, False):
                line = RPar(line_text, style=self._default_p_style).getParagraph(**kwargs)
            return line

        # go through element list and add to section;
        ret = Section()
        for a_element in self._element_cache:
            e_type = a_element.get(self.KEY_TYPE, None)
            e_ctx = a_element.get(self.KEY_VALUE, '')
            e_style = a_element.get(self.KEY_STYLE, None)

            # use captured styles to create document element;
            element_obj = None
            if e_type == self.ELEMENT_PARAGRAPH:
                style_obj = self._style_cache.ParagraphStyles.get_by_name(e_style)
                element_obj = RPar(e_ctx, style=style_obj).getParagraph(**kwargs)
            elif e_type == self.ELEMENT_PARTIAL:
                style_obj = self._style_cache.ParagraphStyles.get_by_name(e_style)
                rp = RPar(None, style=style_obj)
                rp.append(*e_ctx)
                element_obj = rp.getParagraph(**kwargs)
            elif e_type == self.ELEMENT_TABLE:
                cell_s_obj = self._style_cache.ParagraphStyles.get_by_name(e_style)
                head_s_obj = self._style_cache.ParagraphStyles.get_by_name(self._get_bold_style_name(cell_s_obj.name))
                element_obj = RTable(e_ctx, style=cell_s_obj, header_style=head_s_obj).getTable(**kwargs)
            elif e_type == self.ELEMENT_LIST:
                fallback_style_obj = self._style_cache.ParagraphStyles.get_by_name(self.DEFAULT_LIST_STYLE_NAME)
                # try to use 'e_style', and fall back to default style;
                style_obj = self._style_cache.ParagraphStyles.get_by_name(e_style, fallback_style_obj)
                element_obj = RList(e_ctx, style=style_obj).getList(**kwargs)
            else:
                pass
            # push the element object to cache;
            if element_obj:
                if e_type == self.ELEMENT_LIST or (e_type == self.ELEMENT_TABLE and isinstance(element_obj, tuple)):
                    ret.extend(element_obj)
                else:
                    ret.append(element_obj)
                # optional blank line;
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

        # capture all the styles;
        self._style_cache = self._collect_styles(**kwargs)
        # create document object;
        _doc = Document(
            style_sheet=self._style_cache,
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

        @param strip_newline whether the newline character needs to be removed from the output (boolean)

        @rtype string
        """
        _need_strip = kwargs.pop('strip_newline', False)
        _debug_out = kwargs.pop('debug_output', False)

        from StringIO import StringIO
        cache = StringIO()
        self._write(cache, **kwargs)
        ret = cache.getvalue()
        # post-generation manipulation;
        if _need_strip:
            ret = ret.replace('\n', '')
        if _debug_out:
            _ez_args = [
                ('}\\paperw',      '}\n\\paperw'),
                ('footer}{',       'footer}\n{'),
                ('{\\colortbl',    '\n{\\colortbl'),
                ('}{\\fonttbl',    '}\n{\\fonttbl'),
                ('}{\\stylesheet', '}\n{\\stylesheet'),
            ]
            for i in _ez_args:
                ret = ret.replace(*i)
        return ret

    def __repr__(self):
        ret = "<RTF document of {ec} element(s) at {addr}>".format(
            ec=len(self._element_cache),
            addr="0x%x"%(id(self)),
        )
        return ret


#---eof---#
