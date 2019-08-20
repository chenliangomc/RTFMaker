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
                lines = unicode(self._html_content.text).splitlines()
                lines = [ i.strip() for i in lines if len(i.strip()) ]
                self._text_elements = unicode(' ').join(lines)
        except:
            pass

    def getParagraph(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph

        self._convert_text(**kwargs)

        element_obj = Paragraph()
        if self._style is not None:
            element_obj.Style = self._style
        element_obj.append(self._text_elements)
        return element_obj


class RTable(object):
    """internal representation of the table"""

    EMPTY = ''

    def __init__(self, content, **kwargs):
        self._html_content = content
        # TODO: may need document stylesheet reference here to query styles;
        self.EMPTY = kwargs.get('blank_cell', self.EMPTY)

    def _convert_table(self, **kwargs):
        self._table_elements = {
            'head': list(),
            'body': list(),
            'col.cnt': 0,
        }
        # TODO: parse HTML here;
        # normalize the header and body;
        hdr_cnt = len(self._table_elements['head'])
        row_cnt_set = [ len(row) for row in self._table_elements['body'] ]
        self._table_elements['col.cnt'] = max(hdr_cnt, max(row_cnt_set))
        #
        col_count = self._table_elements['col.cnt']
        trailing = [ {'value': self.EMPTY,}, ] * col_count
        if hdr_cnt < col_count:
            self._table_elements['head'] = (self._table_elements['head'] + trailing[:])[:col_count]
        self._table_elements['body'] = [ (row+trailing[:])[:col_count] for row in self._table_elements['body'] ]

    def getTable(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph, Table, Cell

        self._convert_table(**kwargs)
        col_count = self._table_elements['col.cnt']

        ret = Table(*(1270,2540,5080), left_offset=108)

        if len(self._table_elements['head']) > 0:
            header_row = list()
            for a_head in self._table_elements['head'][:col_count]:
                rhead = Cell(
                    Paragraph(a_head['value'], ) # TODO: add p_style and p_prop_set on-demand;
                )
                header_row.append(rhead)
            ret.AddRow(*header_row)

        for row in self._table_elements['body']:
            single_row = list()
            for a_cell in row[:col_count]:
                rcell = Cell(
                    Paragraph(a_cell['value'], ) # TODO: add p_style and p_prop_set on-demand;
                )
                single_row.append(rcell)
            ret.AddRow(*single_row)
        return ret


#--eof--#
