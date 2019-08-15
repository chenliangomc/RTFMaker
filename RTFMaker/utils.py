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
    def __init__(self, content, style = None, **kwargs):
        self._html_content = content
        self._style = style

    def _convert_text(self, **kwargs):
        self._text_elements = str(self._html_content)

    def getParagraph(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph

        self._convert_text(**kwargs)

        element_obj = Paragraph()
        return element_obj


class RTable(object):
    """internal representation of the table"""
    def __init__(self, content, **kwargs):
        self._html_content = content
        # TODO: may need document stylesheet reference here to query styles;

    def _convert_table(self, **kwargs):
        self._table_elements = list()

    def getTable(self, **kwargs):
        from PyRTF.document.paragraph import Paragraph, Table, Cell

        self._convert_table(**kwargs)

        ret = Table(*(1270,2540,5080), left_offset=108)
        return ret


#--eof--#
