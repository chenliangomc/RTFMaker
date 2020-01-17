"""
htmlconv.py is part of RTFMaker, a simple RTF document generation package

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
        ATTR_FONT_DEF = 'FONT_HUB'

        DEFAULT_FONT_DEF = (
            ('med-font',   'font-family:Arial;font-size:9pt;'),
            ('bold-font',  'font-family:Arial;font-size:9pt;font-weight:bold;'),
            ('small-font', 'font-family:Arial;font-size:8pt;'),
            ('ref-text',   'font-family:Arial;font-size:8pt;font-style:italic;'),
        )

        DEFAULT_EXPAND_DIRECTIVE_LABEL = 'expand'
        DEFAULT_NOLINEFEED_DIRECTIVE_LABEL = 'nolinefeed'
        DEFAULT_NOPARENTCLS_DIRECTIVE_LABEL = 'maskparentclass'
        DEFAULT_HTML_ATTR_NAME = 'data-rtf-directive'

        @staticmethod
        def _span_wrap(inner_html, **kw):
            outer_html = '<span>{x}</span>'.format(x=inner_html)
            span_obj = BeautifulSoup(outer_html, 'html.parser').span
            return span_obj

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
            '''
            @param doc ()
            @param tag_list (list,tuple)
            '''
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

        def _load_font_def(self, user_font_def, **kw):
            '''
            @param user_font_def (dict,list,tuple)
            '''
            font_hub = getattr(self, self.ATTR_FONT_DEF, None)
            if (not isinstance(font_hub, dict)):
                font_hub = dict()
            _update_hub = False

            if isinstance(user_font_def, (list,tuple)):
                if len(user_font_def):
                    for font_cls, font_def in user_font_def:
                        font_hub[font_cls] = font_def
                    _update_hub = True
            if isinstance(user_font_def, dict):
                if len(user_font_def):
                    font_hub.update(user_font_def)
                    _update_hub = True
            else:
                if kw.get('debug.use.exc', False):
                    _msg = 'invalid data type: {c}'.format(c=type(user_font_def))
                    raise ValueError(_msg)
            if _update_hub:
                setattr(self, self.ATTR_FONT_DEF, font_hub)
            return font_hub

        def _load_default_font_def(self, **kw):
            font_hub = self._load_font_def(self.DEFAULT_FONT_DEF, **kw)
            return font_hub

        def _map_css_cls_to_font(self, names, default=None, **kw):
            ret = default

            font_hub = getattr(self, self.ATTR_FONT_DEF, None)
            if names:
                for a_cls_name in names:
                    ret = font_hub.get(a_cls_name, None)
                    if ret:
                        break
            return ret

        def _get_extraction_directive(self, node, **kw):
            '''
            @param node

            @return (dict)
            '''
            node_directives = dict()
            attr_directive = kw.get('directive_attribute_name', self.DEFAULT_HTML_ATTR_NAME)
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

        def _get_node_expand_policy(self, tag, **kw):
            '''
            @param tag

            @return boolean or None
            '''
            decision = None

            from bs4.element import NavigableString

            tag_directives = self._get_extraction_directive(tag, **kw)
            attr_expand = tag_directives.get(self.DEFAULT_EXPAND_DIRECTIVE_LABEL, None)
            if isinstance(attr_expand, bool):
                decision = attr_expand
            else:
                if isinstance(tag, NavigableString):
                    decision = False
                elif str(getattr(tag, 'name')).lower() in ('u', 'i', 'br', 'hr', 'ul', 'ol'):
                    decision = False
                else:
                    pass
            return decision

        def _expand_tag(self, node, **kw):
            '''
            @spec if the node cannot be expanded, return empty list

            @param parent.cls (list)
            '''
            ret = list()

            EXEMPT_COUNT = -1
            _parent_cls = kw.get('parent.cls', None)

            cnt_children = EXEMPT_COUNT
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
                        child_directive = self._get_extraction_directive(child, **kw)
                        child_cls = child.get('class')
                        new_cls = HTMLRTF._collect_cls(child_cls, _parent_cls)
                        if child_directive.get(self.DEFAULT_NOPARENTCLS_DIRECTIVE_LABEL, False):
                            new_cls = HTMLRTF._collect_cls(child_cls, None)
                        if isinstance(new_cls, list):
                            child['class'] = new_cls
                    ret.append(child)
                    _idx += 10
            else:
                pass
            return ret

        def _flatten_tag(self, tag, **kw):
            '''
            @param recursive (bool)
            @param parent.cls (list)
            '''
            ROOT_LEVEL = 0
            STEP = 1
            PARAM_DEPTH = 'depth'
            PARAM_PARENT_CLASS = 'parent.cls'

            flat_list = list()

            _recursive = kw.get('recursive', True)
            _depth = kw.get(PARAM_DEPTH, ROOT_LEVEL)

            expand_param = dict()
            expand_param.update(**kw)
            #this_tag_directive = self._get_extraction_directive(tag, **kw)
            try:
                caller_cls = kw.get(PARAM_PARENT_CLASS, None)
                t_cls = tag.get('class')
                combined_cls = HTMLRTF._collect_cls(t_cls, caller_cls)
                #if this_tag_directive.get(self.DEFAULT_NOPARENTCLS_DIRECTIVE_LABEL, False):
                #    combined_cls = HTMLRTF._collect_cls(t_cls, None)
                expand_param[PARAM_PARENT_CLASS] = combined_cls
            except:
                pass

            expanded = list()
            this_tag_do_expand = self._get_node_expand_policy(tag, **kw)
            if this_tag_do_expand == True:
                children = self._expand_tag(tag, **expand_param)
                if len(children) > 0:
                    if _recursive:
                        call_param = dict()
                        call_param.update(kw)
                        call_param[PARAM_PARENT_CLASS] = expand_param[PARAM_PARENT_CLASS]
                        call_param[PARAM_DEPTH] = _depth + STEP

                        for child_tag in children:
                            child_expand = self._flatten_tag(child_tag, **call_param)
                            expanded.extend(child_expand)
                    else:
                        expanded.extend(children)
                else:
                    expanded.append(tag)
            else:
                expanded.append(tag)
            flat_list.extend(expanded)

            if _depth == ROOT_LEVEL:
                flat_list = self._merge_tag(flat_list, **kw)
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

        def _filter_tag(self, tags, **kw):
            tag_cache = list()
            for tag in tags:
                new_tags = self._flatten_tag(tag, **kw)
                tag_cache.extend(new_tags)
            return tag_cache

        def _get_text_from_tag(self, tag, **kw):
            txt_obj = [0, None]

            from bs4.element import Comment

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
            elif isinstance(tag, Comment) or t_name == _empty:
                pass
            else:
                tag_directive = self._get_extraction_directive(tag, **kw)
                if t_name in ('ul', 'ol'):
                    t_cls = tag.get('class')
                    t_font = self._map_css_cls_to_font(t_cls, None, **kw)
                    tmp_dic = {
                        'type': 'list',
                        'value': tag,
                        'font': t_font,
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
                    t_style = tag.get('style')
                    if tag_directive.get(self.DEFAULT_NOPARENTCLS_DIRECTIVE_LABEL, None):
                        if t_font is None and t_style is not None:
                            t_font = t_style
                    need_linefeed = False if t_name in ('span',) else True
                    if tag_directive.get(self.DEFAULT_NOLINEFEED_DIRECTIVE_LABEL, None):
                        need_linefeed = False
                    tmp_dic = {
                        'type': 'paragraph',
                        'value': tag,
                        'font': t_font,
                        #
                        'append_newline': need_linefeed,
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
            '''
            @param tags (list)
            '''
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
            @param css_font_def (dict/list)

            @return RTF stream (string)
            '''
            self._load_default_font_def(**kw)
            user_font = kw.pop('css_font_def', None)
            self._load_font_def(user_font, **kw)

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
                <div class="bold-font" data-rtf-extract="simple-list-body" data-rtf-directive="expand style=bold">
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
                'css_font_def': [ ('large-font',   'font-family:Arial;font-size:11pt;'), ],
            }
            demo_param.update(kw)
            ret['rtf'] = self.translate(ret['html'], ret['tags'], **demo_param)
            return ret

    return HTMLRTF

