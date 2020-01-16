#!/usr/bin/env python
# -*- coding:utf-8 -*-
#
# This file is part of RTFMaker, a simple RTF document generation package
# Copyright (C) 2019, 2020  Liang Chen
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""
RTFMaker
========

RTFMaker is a simple RTF document generation Python package. It aims to be a
light solution, easy to use.

"""

classifiers = """\
Development Status :: 4 - Beta
Topic :: Text Editors :: Text Processing
Topic :: Software Development :: Libraries :: Python Modules
Intended Audience :: Developers
Programming Language :: Python
Operating System :: OS Independent
License :: OSI Approved :: GNU Affero General Public License v3 or later (AGPLv3+)
"""

from setuptools import setup, find_packages

def find_version(*file_paths):
    import codecs
    import re
    import os

    def read(*parts):
        here = os.path.abspath(os.path.dirname(__file__))
        with codecs.open(os.path.join(here, *parts), 'r') as fp:
            return fp.read()

    version_file = read(*file_paths)
    version_match = re.search(
        r"^__version__ = ['\"]([^'\"]*)['\"]",
        version_file,
        re.M
    )
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")


setup(
    name='RTFMaker',
    version=find_version('RTFMaker', '__init__.py'),
    description='simple RTF document generation package',
    #long_description=__doc__,
    author="Liang Chen",
    author_email="liangchenomc@gmail.com",
    url="https://github.com/chenliangomc/RTFMaker",
    license='https://www.gnu.org/licenses/agpl-3.0',
    platforms="Any",
    packages=find_packages(),
    install_requires=['PyRTF3', 'beautifulsoup4'],
    keywords=('RTF', 'Rich Text', 'Rich Text Format'),
    python_requires=">=2.7",
    classifiers=[_f for _f in classifiers.split('\n') if _f],
    long_description_markdown_filename='readme.md',
)
