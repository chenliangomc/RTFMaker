"""
utils.py is part of RTFMaker, a simple RTF document generation package

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


class RPar(object):
    """internal representation of the paragraph"""
    def __init__(self, content, style=None, **kwargs):
        self._html_content = content
        self._style = style

    def _convert_text(self, **kwargs):
        self._text_elements = unicode(self._html_content)
        try:
            if not isinstance(self._html_content, (basestring, unicode)):
                self._text_elements = unicode(
                    self._html_content.get_text(
                        separator=' ',
                        strip=True
                    )
                )
        except:
            pass

    def getParagraph(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph

        prefix_element = kwargs.pop('prefix', None)

        self._convert_text(**kwargs)

        element_obj = Paragraph()
        if self._style is not None:
            element_obj.Style = self._style
        if prefix_element:
            element_obj.append(prefix_element)
            element_obj.append(unicode(' '))
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
    }

    def __init__(self, content, style=None, header_style=None, **kwargs):
        self._html_content = content
        self._cell_style = style
        self._head_style = header_style
        self.EMPTY = kwargs.get('blank_cell', self.EMPTY)

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
                from bs4 import BeautifulSoup
                obj = BeautifulSoup(self._html_content)
            html_head = getattr(obj, 'thead')
            if html_head:
                for a_col in html_head.find_all('th'):
                    tmp_col_head = {
                        'value': a_col.get_text(strip=True),
                    }
                    self._table_elements['head'].append(tmp_col_head)
            html_body = getattr(obj, 'tbody')
            if html_body:
                for a_row in html_body.find_all('tr'):
                    new_row = list()
                    for a_cell in a_row.find_all('td'):
                        tmp_cell = {
                            'value': a_cell.get_text(strip=True),
                        }
                        new_row.append(tmp_cell)
                    self._table_elements['body'].append(new_row)
            html_foot = getattr(obj, 'tfoot')
            if html_foot:
                for a_foot in html_foot.find_all('td'):
                    foot_cell = {
                        'value': a_foot.get_text(strip=True),
                    }
                self._table_elements['foot'].append(foot_cell)

        # normalize the header and body;
        hdr_cnt = len(self._table_elements['head'])
        row_cnt_set = [ len(row) for row in self._table_elements['body'] ]
        assert len(row_cnt_set) > 0, 'empty table'
        self._table_elements['col.cnt'] = max(hdr_cnt, max(row_cnt_set))
        #
        col_count = self._table_elements['col.cnt']
        trailing = [ {'value': self.EMPTY,}, ] * col_count
        if hdr_cnt > 0 and hdr_cnt < col_count:
            self._table_elements['head'] = (self._table_elements['head'] + trailing[:])[:col_count]
        self._table_elements['body'] = [ (row+trailing[:])[:col_count] for row in self._table_elements['body'] ]

    def getTable(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph, Table, Cell

        self._convert_table(**kwargs)
        col_count = self._table_elements['col.cnt']

        ret = Table(left_offset=108)
        ret.SetColumnWidths(*(self.TABLE_COLUMN_PRESET[col_count]))

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
            foot_p = Paragraph(self._table_elements['foot'][0]['value'])
            combined = (ret, foot_p)
            ret = combined
        return ret


class RList(object):
    """internal representation of the list"""
    def __init__(self, content, style=None, **kwargs):
        self._html_content = content
        self._style = style

    def _convert_list(self, **kwargs):
        self._list_elements = list()

    def getList(self, **kwargs):
        from PyRTF.document.base import RawCode

        self._convert_list(**kwargs)

        ret = list()
        for item in self._list_elements:
            tmp_dic = {
                'prefix': RawCode(r'\u8226'),
            }
            tmp_dic.update(kwargs)
            item_par = RPar(item, style=self._style, **kwargs).getParagraph(**tmp_dic)
            ret.append(item_par)
        return ret


#--eof--#
