"""
htmlconv.py is part of RTFMaker, a simple RTF document generation package

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

class _empty(object):
    """special placeholder"""
    pass


def get_html_translator(base_cls, **kwargs):
    '''
    factory method for HTML-to-RTF translator class

    @param base_cls the base class for the translator class (class/type)
    '''
    from inspect import isclass
    from bs4 import BeautifulSoup

    assert isclass(base_cls), 'invalid argument value'

    class HTMLRTF(base_cls):

        @staticmethod
        def _span_wrap(inner_html, **kw):
            outer_html = '<span>{x}</span>'.format(x=inner_html)
            span_obj = BeautifulSoup(outer_html, 'html.parser').span
            return span_obj

        @staticmethod
        def _get_extraction_directive(node, **kw):
            node_directives = dict()
            attr_directive = kw.get('directive_attibute_name', 'data-rtf-directive')
            attr_value = _empty()
            try:
                attr_txt = node.get(attr_directive)
                attr_token = [ i.strip() for i in attr_txt.split(' ') ]
                attr_value = [ i for i in attr_token if len(i) > 0 ]
            except:
                pass
            if isinstance(attr_value, (list,tuple)):
                for a_directive in attr_value:
                    if a_directive.find('=') > -1:
                        token = a_directive.split('=', 1)
                        node_directives[ token[0] ] = token[1]
                    else:
                        node_directives[ a_directive ] = True
            return node_directives

        @staticmethod
        def _collect_cls(*args):
            '''
            @return combined list or None
            '''
            cache  = list()
            for a_cls in args:
                if isinstance(a_cls, (list,tuple)):
                    cache.extend(a_cls)
            if len(cache) > 0:
                return cache
            return None

        @staticmethod
        def _extract_tag(doc, tag_list, **kw):
            ret = list()

            _add_na = kw.get('add.na', False)
            _na_str = kw.get('na.str', '&nbsp;')

            _EMPTY_SPAN = HTMLRTF._span_wrap(_na_str, **kw)

            placeholder = kw.get('placeholder', _EMPTY_SPAN)

            for a_attr in tag_list:
                if isinstance(a_attr, dict):
                    tag_obj = doc.findAll(attrs=a_attr)
                    if len(tag_obj) == 0 and _add_na:
                        tag_obj.append(placeholder)
                    ret.extend(tag_obj)
            return ret

        @staticmethod
        def _font_def_validator(font_def, **kw):
            '''
            validate the font definition

            @param font_def (dict)
            '''
            valid = False
            assert isinstance(font_def, dict), 'invalid data type'
            FONT_FIELD_HUB = {
                'font-family': {
                    'required': True,
                    'type': (basestring, unicode),
                    'enumerate': ['Arial','Courier New','Times','Tahoma',],
                },
                'font-size': {
                    'required': True,
                    'type': (basestring, unicode, int),
                    'unit': 'pt',
                    'enumerate': ['8pt','9pt',],
                },
                'font-weight': {
                    'required': False,
                    'type': (basestring, unicode),
                    'enumerate': ['bold',],
                },
                'font-style': {
                    'required': False,
                    'type': (basestring, unicode),
                    'enumerate': ['italic',],
                },
            }
            try:
                if len(font_def) == 0:
                    _msg = "empty font definition"
                    raise ValueError(_msg)
                for a_field, a_action in FONT_FIELD_HUB.items():
                    a_def = font_def.get(a_field, _empty)
                    if a_action['required'] and a_def == _empty:
                        _msg = "required field '{f}' is missing".format(f=a_field)
                        raise ValueError(_msg)
                    if not isinstance(a_def, a_action['type']):
                        _msg = "invalid data type for field '{f}'".format(f=a_field)
                        raise ValueError(_msg)
                    if not a_def in a_action['enumerate']:
                        _msg = "invalid data type for field '{f}'".format(f=a_field)
                        raise ValueError(_msg)
                valid = True
            except:
                pass
            return valid

        def _map_css_cls_to_font(self, names, default=None, **kw):
            ret = default

            replacements = [
                ('med-font',   'font-family:Arial;font-size:9pt;'),
                ('bold-font',  'font-family:Arial;font-size:9pt;font-weight:bold;'),
                ('small-font', 'font-family:Arial;font-size:8pt;'),
                ('ref-text',   'font-family:Arial;font-size:8pt;font-style:italic;'),
            ]
            font_hub = dict(replacements)

            if names:
                for a_cls_name in names:
                    ret = font_hub.get(a_cls_name, None)
                    if ret:
                        break
            return ret

        @staticmethod
        def _expand_tag(node, **kw):
            '''
            @spec if the node cannot be expanded, return empty list
            '''
            ret = list()


            _parent_cls = kw.get('parent.cls', None)

            from bs4.element import NavigableString

            _is_exempted_cls = False
            if isinstance(node, NavigableString):
                _is_exempted_cls = True
            elif str(getattr(node, 'name')).lower() in ('u', 'i', 'br', 'hr', 'ul', 'ol'):
                _is_exempted_cls = True
            else:
                pass

            EXEMPT_COUNT = -1
            cnt_children = EXEMPT_COUNT
            if not _is_exempted_cls:
                try:
                    children = getattr(node, 'children')
                    cnt_children = children.__length_hint__()
                except:
                    pass

            if cnt_children > EXEMPT_COUNT:
                _idx = 0
                for child in node.children:
                    child_name = child.name

                    if child_name:
                        child_cls = child.get('class')
                        new_cls = list()
                        for cls_set in (child_cls, _parent_cls):
                            if cls_set:
                                new_cls.extend(cls_set)
                        if len(new_cls):
                            child['class'] = new_cls
                    ret.append(child)
                    _idx += 10
            else:
                pass
            return ret

        @staticmethod
        def _flatten_tag(tag, **kw):
            ROOT_LEVEL = 0
            STEP = 1
            PARAM_DEPTH = 'depth'

            flat_list = list()

            _recursive = kw.get('recursive', True)
            _depth = kw.get(PARAM_DEPTH, ROOT_LEVEL)

            expand_param = dict()
            expand_param.update(**kw)
            try:
                t_cls = tag.get('class')
                expand_param['parent.cls'] = t_cls
            except:
                pass

            tag_directives = HTMLRTF._get_extraction_directive(tag, **kw)
            this_tag_do_expand = tag_directives.get('expand', None)
            if not isinstance(this_tag_do_expand, bool):
                this_tag_do_expand = False
            children = HTMLRTF._expand_tag(tag, **expand_param)
            if this_tag_do_expand and len(children) > 0:
                expanded = list()
                if _recursive:
                    call_param = dict()
                    call_param.update(kw)
                    call_param[PARAM_DEPTH] = _depth + STEP

                    for child_tag in children:
                        child_expand = HTMLRTF._flatten_tag(child_tag, **call_param)
                        expanded.extend(child_expand)
                else:
                    expanded.extend(children)
                flat_list.extend(expanded)
            else:
                flat_list.append(tag)

            if _depth == ROOT_LEVEL:
                flat_list = HTMLRTF._merge_tag(flat_list, **kw)
            return flat_list

        @staticmethod
        def _merge_tag(tags, **kw):
            new_tags = list()

            MARK = 1

            def _detect_br_blank(tag, **kw):
                ret = 0
                if str(tag.name).lower() == 'br':
                    ret = MARK
                elif len(unicode(tag).strip()) == 0:
                    ret = MARK
                return ret

            def _rollover(stack, callback=None, **kw):
                cnt = len(stack)
                if callable(callback):
                    if cnt == 1:
                        callback(stack[0])
                    elif cnt > 1:
                        callback(stack[:])
                    else:
                        pass
                return list()

            flag = [ _detect_br_blank(i, **kw) for i in tags ]
            cnt_flag = sum(flag)
            _cb = new_tags.append
            if cnt_flag > 0:
                if cnt_flag < len(flag):
                    stak = list()
                    for i in range(len(flag)):
                        if flag[i] == MARK:
                            stak = _rollover(stak, callback=_cb)
                        else:
                            stak.append(tags[i])
                        pass
                    stak = _rollover(stak, callback=_cb)
            else:
                new_tags.extend(tags)
            return new_tags

        @staticmethod
        def _filter_tag(tags, **kw):
            tag_cache = list()
            for tag in tags:
                new_tags = HTMLRTF._flatten_tag(tag, **kw)
                tag_cache.extend(new_tags)
            return tag_cache

        def _get_text_from_tag(self, tag, **kw):
            txt_obj = [0, None]

            _use_exc = kw.get('use_exc', False)
            _func = kw.get('callback.text.extraction', None)

            t_name = getattr(tag, 'name', _empty)
            if isinstance(tag, (list,tuple)):
                tt_cache = list()
                for tt in tag:
                    tt_cache.append(self._get_text_from_tag(tt, **kw)[1])

                tmp_dic = {
                    'type': 'partial',
                    'value': tt_cache,
                    #
                    'append_newline': True,
                }
                txt_obj[1] = tmp_dic
                txt_obj[0] = 1
                pass
            elif t_name == _empty:
                pass
            else:
                if t_name in ('ul', 'ol'):
                    # TODO: text style not propgate to children tags;
                    tmp_dic = {
                        'type': 'list',
                        'value': tag,
                        #
                        'append_newline': True,
                    }
                    txt_obj[1] = tmp_dic
                    txt_obj[0] = 1
                elif t_name in ('table',):
                    tmp_dic = {
                        'type': 'table',
                        'value': tag,
                        #
                        'append_newline': True,
                    }
                    txt_obj[1] = tmp_dic
                    txt_obj[0] = 1
                elif t_name in ('p','span', 'div'):
                    t_cls = tag.get('class')
                    t_font = self._map_css_cls_to_font(t_cls, None, **kw)
                    tmp_dic = {
                        'type': 'paragraph',
                        'value': tag,
                        'font': t_font,
                        #
                        'append_newline': False if t_name in ('span',) else True,
                    }
                    if callable(_func):
                        tmp_dic = _func(tmp_dic, **kw)
                    txt_obj[1] = tmp_dic
                    txt_obj[0] = 1
                elif t_name in ('br',):
                    pass
                elif t_name in (None,'u','i',):
                    span_tag = self._span_wrap(unicode(tag), **kw)

                    tmp_dic = {
                        'type': 'paragraph',
                        'value': span_tag,
                        #
                        'append_newline': False,
                    }
                    txt_obj[1] = tmp_dic
                    txt_obj[0] = 1
                    pass
                else:
                    _msg = "cannot handle element type:{t}".format(t=t_name)
                    if _use_exc:
                        raise RuntimeError(_msg)

            return txt_obj

        def _tag2txt(self, tags, **kw):
            txt_list = list()

            for tag in tags:
                txt = self._get_text_from_tag(tag, **kw)
                txt_def = txt[1]
                if txt_def is not None:
                    txt_list.append(txt_def)
            return txt_list

        def translate(self, raw_html, tag_set, **kw):
            '''
            @param raw_html (string)
            @param tag_set (list)

            @return RTF stream (string)
            '''
            dom = BeautifulSoup(raw_html, 'html.parser')
            raw_tags = self._extract_tag(dom, tag_set, **kw)
            final_tags = self._filter_tag(raw_tags, **kw)
            txt_cache = self._tag2txt(final_tags, **kw)
            from . import RTFDocument
            r = RTFDocument(**kw)
            for i in txt_cache:
                r.append(i)
            return r.to_string(**kw)

        def demo(self, **kw):
            '''
            try parameter 'strip_newline=True' and see the differences of the output
            '''
            ret = {
                'html': '''<!doctype html>
<html>
<head>
    <title>demo page</title>
</head>
<body>
    <div class="wrapper">
        <div class="large-font" data-rtf-extract="page-title">Sample page title</div>
        <div class="content-body col-lg-12">
            <div class="row">
                <div class="bold-font" data-rtf-extract="introduction-section-title">Introduction</div>
                <div class="normal-font" data-rtf-extract="introdcution-section-body" data-rtf-directive="expand">
                    <p>This is the first line of introduction.</p>
                    <br>
                    <p>This is the second line.</p>
                </div>
                <div class="bold-font" data-rtf-extract="simple-list-title">A simple list</div>
                <div class="normal-font" data-rtf-extract="simple-list-body" data-rtf-directive="expand style=bold">
                    <ul>
                        <li>First item</li>
                        <li>Second item</li>
                        <li>Third and the last item</li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="bold-font" data-rtf-extract="table-title">A table</div>
                <table data-rtf-extract="table-body">
                    <thead>
                        <tr>
                            <th>A</th>
                            <th>b</th>
                            <th>C</th>
                        </tr>
                    </thead>
                    <tbody>
                        <tr>
                            <td>one</td>
                            <td>two</td>
                            <td>3</td>
                        </tr>
                    </tbody>
                </table>
            </div>
        </div>
    </div>
</body>
</html>''',
                'tags': [
                    {'data-rtf-extract':'page-title'},
                    {'data-rtf-extract':'introduction-section-title'},
                    {'data-rtf-extract':'introdcution-section-body'},
                    {'data-rtf-extract':'simple-list-title'},
                    {'data-rtf-extract':'simple-list-body'},
                    {'data-rtf-extract':'table-title'},
                    {'data-rtf-extract':'table-body'},
                    {'data-rtf-extract':'end-note'},
                ],
                'rtf': None,
            }
            demo_param = {
                'add.na': True,
                'na.str': '&lt;no text here&gt;',
            }
            demo_param.update(kw)
            ret['rtf'] = self.translate(ret['html'], ret['tags'], **demo_param)
            return ret

    return HTMLRTF

