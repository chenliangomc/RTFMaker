"""
utils.py is part of RTFMaker, a simple RTF document generation package

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

from PyRTF.PropertySets import AttributedList

class StyleSet(AttributedList):
    """generic style object pool"""
    def __init__(self, *args):
        super(StyleSet, self).__init__(*args)

    def get_by_name(self, name, default=None):
        """query registered style object pool with style name

        @param name the name of the style object (string)
        @param default the default value if lookup fails
        """
        ret = default
        # go through object pool and check attribute value;
        ref = None
        for i in self:
            i_name = getattr(i, 'name', None)
            if i_name == name:
                ref = i
                break
        if ref is not None:
            ret = ref
        # rvalue;
        return ret

    def add(self, *values):
        """register new style object into the object pool"""
        for value in values:
            name = getattr(value, 'name', None)
            existing_item = None
            if name:
                existing_item = self.get_by_name(name)
            if existing_item is not None:
                continue
            self.append(value)

    def get_names(self, **kwargs):
        names = list()
        for i in self:
            i_name = getattr(i, 'name', None)
            names.append(i_name)
        return names


def _text_strip(x, **kwargs):
    """
    @param x text object (string or obj)
    """
    LINE_SEP = ' '
    sep_char = unicode(kwargs.get('', LINE_SEP))
    ret = unicode(x)
    if not isinstance(x, (basestring, unicode)):
        cache = list()
        lines = x.get_text(strip=True).splitlines()
        for line in lines:
            striped_line = line.strip()
            if len(striped_line) > 0:
                cache.append(unicode(striped_line))
        ret = sep_char.join(cache)
    return ret


def _htmlify(x, **kwargs):
    """
    @param x text object (string)
    """
    from bs4 import BeautifulSoup
    html_obj = BeautifulSoup(x, 'html.parser')
    return html_obj


class RPar(object):
    """internal representation of the paragraph"""

    DELIMITER_PREFIX = '   '
    DELIMITER_INLINE = ' '

    def __init__(self, content, style=None, **kwargs):
        self._html_content = content
        self._style = style

    def _convert_text(self, **kwargs):
        if getattr(self, '_text_elements', None):
            pass
        else:
            self._text_elements = _text_strip(self._html_content)

    def append(self, *values):
        self._text_elements = list()
        from PyRTF.document.character import Text

        _idx = 0
        for value in values:
            if value is not None:
                a_text = _text_strip(value['value'])
                a_style = value.get('style', self._style)
                new_item = Text()
                if isinstance(a_style, type(self._style)):
                    new_item.Style = a_style.TextStyle
                else:
                    #from PyRTF.Styles import ParagraphStyle
                    #from PyRTF.PropertySets import ParagraphPropertySet
                    # TODO: parse and get the actual text_style here, wrapped with ParagraphStyle object;
                    pass
                if _idx > 0:
                    new_item.SetData(unicode(self.DELIMITER_INLINE))
                new_item.SetData(a_text)
                self._text_elements.append(new_item)
            _idx += 1

    def getParagraph(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph

        prefix_element = kwargs.pop('prefix', None)

        self._convert_text(**kwargs)

        element_obj = Paragraph()
        if self._style is not None:
            element_obj.Style = self._style
        if prefix_element:
            element_obj.append(prefix_element)
            element_obj.append(unicode(self.DELIMITER_PREFIX))
        if isinstance(self._text_elements, (list, tuple)):
            for atext in self._text_elements:
                element_obj.append(atext)
        else:
            element_obj.append(self._text_elements)
        return element_obj


class RTable(object):
    """internal representation of the table"""

    EMPTY = ''

    TABLE_COLUMN_PRESET = {
        2: (3810,5080),
        3: (1270,2540,5080),
        4: (1270,2540,2540,2540),
        5: (2540,2540,1270,1270,1270),
        6: (1270,1270,1270,1270,1270,1270),
        #6: (1905,1905,1270,1270,1270,1270),
        #6: (2540,1905,1270,1270,635,1270),
        #6: (2540,1905,1270,1270,635,635+3*90),
        #6: (1905+4*90,1905,1270,1270,635,635+3*90),
        #6: (1905+4*90,1905+1270,1270,1270,635,635+3*90),
    }
    TEXT_WIDTH = 9420 # 1270*6+1800=9420; 1270*7+7*90=9520; left_offset=108;

    def __init__(self, content, style=None, header_style=None, foot_style=None, **kwargs):
        self._html_content = content
        self._cell_style = style
        self._head_style = header_style
        self._foot_style = foot_style
        self.EMPTY_CELL = kwargs.get('blank_cell', self.EMPTY)

    def _convert_table(self, **kwargs):
        self._table_elements = {
            'head': list(),
            'body': list(),
            'foot': list(),
            'col.cnt': 0,
        }
        # parse HTML here;
        if isinstance(self._html_content, dict):
            self._table_elements.update(self._html_content)
        else:
            obj = self._html_content
            if isinstance(self._html_content, (basestring, unicode)):
                obj = _htmlify(self._html_content, **kwargs)
            html_head = getattr(obj, 'thead')
            if html_head:
                for a_col in html_head.find_all('th'):
                    tmp_col_head = {
                        'value': _text_strip(a_col),
                    }
                    self._table_elements['head'].append(tmp_col_head)
            html_body = getattr(obj, 'tbody')
            if html_body:
                for a_row in html_body.find_all('tr'):
                    new_row = list()
                    for a_cell in a_row.find_all('td'):
                        tmp_cell = {
                            'value': _text_strip(a_cell),
                        }
                        new_row.append(tmp_cell)
                    self._table_elements['body'].append(new_row)
            html_foot = getattr(obj, 'tfoot')
            if html_foot:
                for a_foot in html_foot.find_all('td'):
                    foot_cell = {
                        'value': _text_strip(a_foot),
                    }
                    self._table_elements['foot'].append(foot_cell)

        # normalize the header and body;
        hdr_cnt = len(self._table_elements['head'])
        row_cnt_set = [ len(row) for row in self._table_elements['body'] ]
        assert len(row_cnt_set) > 0, 'empty table'
        self._table_elements['col.cnt'] = max(hdr_cnt, max(row_cnt_set))
        #
        col_count = self._table_elements['col.cnt']
        trailing = [ {'value': self.EMPTY_CELL,}, ] * col_count
        if hdr_cnt > 0 and hdr_cnt < col_count:
            self._table_elements['head'] = (self._table_elements['head'] + trailing[:])[:col_count]
        self._table_elements['body'] = [ (row+trailing[:])[:col_count] for row in self._table_elements['body'] ]

    def _get_column_layout(self, colcnt, **kwargs):
        """
        @param colcnt column count (positive integer)
        @param table_column_layout (dict)
        """
        assert colcnt >= 1, 'no columns in the table'
        additional_layout = kwargs.get('table_column_layout', None)
        HUB = dict()
        HUB.update(self.TABLE_COLUMN_PRESET)
        if isinstance(additional_layout, dict):
            HUB.update(additional_layout)
        cell_width = int(self.TEXT_WIDTH/float(colcnt))
        evenly_split = tuple([ cell_width for i in range(colcnt) ])
        ret = HUB.get(colcnt, evenly_split)
        return ret

    def getTable(self, **kwargs):
        """
        @param table_left_offset (integer)
        @param merged_footer whether combine the content of all the footer cells into one paragraph (boolean)
        @param space_before_footer insert blank line before the merged footer paragraph (boolean)
        """
        from PyRTF.document.paragraph import Paragraph, Table, Cell
        #from PyRTF.PropertySets import ParagraphPropertySet

        self._convert_table(**kwargs)
        col_count = self._table_elements['col.cnt']

        tbl_left_offset = kwargs.get('table_left_offset', 108)
        ret = Table(left_offset=tbl_left_offset)
        tbl_layout = self._get_column_layout(col_count, **kwargs)
        ret.SetColumnWidths(*(tbl_layout))

        if len(self._table_elements['head']) > 0:
            header_row = list()
            for a_head in self._table_elements['head'][:col_count]:
                head_p = Paragraph(a_head['value'])
                if self._head_style:
                    head_p.Style = self._head_style
                rhead = Cell(head_p)
                header_row.append(rhead)
            ret.AddRow(*header_row)

        for row in self._table_elements['body']:
            single_row = list()
            for a_cell in row[:col_count]:
                cell_p = Paragraph(a_cell['value'])
                if self._cell_style:
                    cell_p.Style = self._cell_style
                rcell = Cell(cell_p)
                single_row.append(rcell)
            ret.AddRow(*single_row)

        if len(self._table_elements['foot']) > 0:
            if kwargs.get('merged_footer', True):
                combined = list()
                combined.append(ret)
                if kwargs.get('space_before_footer', True):
                    spacer_p = Paragraph('')
                    combined.append(spacer_p)
                foot_p = Paragraph(self._table_elements['foot'][0]['value'])
                if self._foot_style:
                    foot_p.Style = self._foot_style
                combined.append(foot_p)
                # override rvalue;
                ret = tuple(combined)
            else:
                foot_row = list()
                for a_foot in self._table_elements['foot'][:col_count]:
                    foot_p = Paragraph(a_foot['value'])
                    if self._foot_style:
                        foot_p.Style = self._foot_style
                    rfoot = Cell(foot_p)
                    foot_row.append(rfoot)
                ret.AddRow(*foot_row)
        return ret


class RList(object):
    """internal representation of the list"""

    ITEM_TYPE_NORMAL = 'li'
    ITEM_TYPE_PLAIN = 'plain'

    def __init__(self, content, style=None, **kwargs):
        self._html_content = content
        self._style = style

    def _convert_list(self, **kwargs):
        self._list_elements = list()
        # parse HTML here;
        obj = self._html_content
        if isinstance(self._html_content, (basestring, unicode)):
            obj = _htmlify(self._html_content, **kwargs)
        for item in obj.find_all():
            item_name = getattr(item, 'name', None)
            item_type = self.ITEM_TYPE_PLAIN
            if str(item_name).lower() in ('li',):
                item_type = self.ITEM_TYPE_NORMAL
            item_text = _text_strip(item)
            tmp_dic = {
                'text': item_text,
                'type': item_type,
            }
            self._list_elements.append(tmp_dic)

    def _bullet_point(self, **kwargs):
        from PyRTF.document.base import RawCode
        hub = {
            'bullet': RawCode(r'\u9679\'3f'),#RawCode(r'\u8729\'b7'),#RawCode(r'\u8226'),
            'star': '*',
            'minus': '-',
            'plus': '+',
            'circle': 'o',
        }
        return hub

    def getList(self, **kwargs):
        """
        @param list_symbol_name (string)
        """
        #from PyRTF.document.paragraph import Enumerate, Item

        self._convert_list(**kwargs)

        symbol_name = kwargs.get('list_symbol_name', 'bullet')
        prefix_symbol = self._bullet_point(**kwargs)[symbol_name]

        #ret = Enumerate()
        ret = list()
        for item in self._list_elements:
            tmp_dic = dict()
            if item['type'] == self.ITEM_TYPE_NORMAL:
                tmp_dic['prefix'] = prefix_symbol
            tmp_dic.update(kwargs)
            item_par = RPar(item['text'], style=self._style, **kwargs).getParagraph(**tmp_dic)
            ret.append(item_par)
        return ret


#class RFigure(object):
#    def __init__(self, content, style=None, **kwargs):
#        self._html_content = content
#        self._style = style
#
#    def getFigure(self, **kwargs):
#        from PyRTF.document.base import TAB, LINE
#        from PyRTF.object.picture import Image


#--eof--#
