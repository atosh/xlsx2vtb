# coding: utf-8
import re

def xml2dict(elem):
    ret = {
        'tag': re.sub(r'^{.+}', '', elem.tag),
        'text': elem.text,
        'tail': elem.tail,
        'children': [xml2dict(child) for child in elem.getchildren()],
    }
    ret.update({re.sub(r'^{.+}', '', k): v for k, v in elem.attrib.items()})
    return ret


