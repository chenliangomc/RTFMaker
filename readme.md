RTFMaker
========

A simple RTF document generation package

RTFMaker aims to be a light solution, easy to use.

Quick Start
-----------

```python
from RTFMaker import RTFDocument

r = RTFDocument()

cache = [
    {
        'type':'paragraph',
        'value': 'a simple line in this document.',
        'append_newline': True,
    },
    {
        'type':'paragraph',
        'value': 'This is a line that has the style is different than the default style: bold.',
        'append_newline': True,
        'font': 'font-family:Arial;font-weight:bold;font-size:10pt;',
    },
    {
        'type':'paragraph',
        'value': 'This is a line that has the style is different than the default style: italic.',
        'append_newline': True,
        'font': 'font-family:Arial;font-style:italic;font-size:10pt;',
    },
]

for i in cache:
    r.append(i)

print r.to_string()
```

TODO
----

- implement parsing logic for table
- add support for paragraph properties
- add support for image
- add config for PyLint, Flask8
- add unittest

License
-------

This software is licensed under GNU Affero General Public License Version 3 or any later version.

You should have received a copy of the GNU Affero General Public License along with this program.  If not, see [here](https://www.gnu.org/licenses/).
