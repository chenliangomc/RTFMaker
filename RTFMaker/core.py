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

    DEFAULT_LANGUAGE = 'EnglishUS'

    def __init__(self, **kwargs):
        self._element_cache = list()
        self._default_p_style = None
        self._doc = None

    def append(self, element):
        """add new element to the document"""
        self._element_cache.append(element)
        return None


    def _collect_styles(self, **kwargs):
        """get all the registered styles

        return `PyRTF.Elements.StyleSheet`
        """

    def _collect_elements(self, **kwargs):
        """get all the elements

        return `PyRTF.document.section.Section`
        """

    def _to_rtf(self, **kwargs):
        """convert internal representation of document structure into RTF stream

        return `PyRTF.Elements.Document`
        """
        from PyRTF.Constants import Languages
        from PyRTF.Elements import Document

        # create document object, and capture all the styles;
        _doc = Document(
            style_sheet=self._collect_styles(**kwargs),
            default_language=getattr(Languages, self.DEFAULT_LANGUAGE),
        )

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
