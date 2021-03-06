#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
python3.5 works as-is
json lib does not work as expected with python2.7
seems utf-8 string related
pandoc -t docx -o <docx filename> <input filename>
    --filter=pandoc-crossref --filter=pandoc-docx-utils
"""

import os
import hashlib
import panflute as pf
import subprocess as sp
from shutil import which


class UnnumberHeadings(object):
    """
    converts level 1~4 headers in 'unnumbered' class to unnumbered headers
    * works with docx output only
    * Level 5 header is unnumbered regardless to the class
    * "Heading Unnumbered x" must be prepared in template

    | Level | Numbered  | Unnumbered            |
    |-------------------------------------------|
    | 1     | Heading 1 | Heading Unnumbered 1  |
    | 2     | Heading 2 | Heading Unnumbered 2  |
    | 3     | Heading 3 | Heading Unnumbered 3  |
    | 4     | Heading 4 | Heading Unnumbered 4  |
    | 5     |           | Heading 5             |
    """

    def action(self, elem, doc):
        if (doc.format == "docx"):
            default_style = [doc.get_metadata("heading-unnumbered.1", "Heading Unnumbered 1"),
                             doc.get_metadata("heading-unnumbered.2", "Heading Unnumbered 2"),
                             doc.get_metadata("heading-unnumbered.3", "Heading Unnumbered 3"),
                             doc.get_metadata("heading-unnumbered.4", "Heading Unnumbered 4"),
                             ]
            # pf.debug(default_style)
            if isinstance(elem, pf.Header) and "unnumbered" in elem.classes:
                if elem.level < 5:
                    style = elem.attributes.get("custom-style", default_style[elem.level - 1])
                    elem.attributes.update({"custom-style": style})
                    elem = pf.Div(pf.Para(*elem.content), attributes=elem.attributes,
                                  identifier=elem.identifier, classes=elem.classes)
                    # pf.debug(elem)
        return elem


class InlineFigureCentered(object):
    """
    moves a Para containing an Image into to "Image Caption" style div
    * "Image Caption" custom style should be prepared in template
    """

    def action(self, elem, doc):
        if (doc.format == "docx"):
            default_style = doc.get_metadata("image-div-style", "Image Caption")
            if isinstance(elem, pf.Para) and len(elem.content) == 1:
                for subelem in elem.content:
                    if isinstance(subelem, pf.Image):
                        style = subelem.attributes.get("custom-style", default_style)
                        subelem.attributes["custom-style"] = style
                        subelem.attributes.pop("custom-style")
                        d = pf.Div(elem, attributes={"custom-style": style})
                        return d


class Svg2Png(object):
    def __init__(self):
        self.dir_to = "svg"
        assert which("rsvg-convert"), "rsvg-convert does not exist in path"

    def action(self, elem, doc):
        if doc.format not in ["html", "html5"]:
            if isinstance(elem, pf.Image):
                fn = elem.url
                if fn.endswith(".svg"):
                    counter = hashlib.sha1(fn.encode("utf-8")).hexdigest()[:8]

                    if not os.path.exists(self.dir_to):
                        pf.debug("mkdir")
                        os.mkdir(self.dir_to)

                    basename = "/".join([self.dir_to, str(counter)])

                    if doc.format in ["latex"]:
                        format = "pdf"
                    else:
                        format = "png"

                    linkto = os.path.abspath(".".join([basename, format]))
                    elem.url = linkto
                    command = "rsvg-convert {} -f {} -o {}".format(fn, format, linkto)
                    sp.Popen(command, shell=True, stdin=sp.PIPE, stdout=sp.PIPE, stderr=sp.PIPE)
                    pf.debug("[inline] convert a svg file to {}".format(format))


class ExtractBulletList(object):
    """
    Extracts bullet list to "Bullet List 1" and "Bullet List 2" styles
    * "Bullet List 1" and "Bullet List 2" must be prepared in template
    * Level-3 and lower lists are forced to be Level-2
    """

    def get_depth(self, elem):
        depth = 0
        # pf.debug("elem: {}".format(elem))
        while not isinstance(elem, pf.Doc):
            if isinstance(elem, pf.ListItem):
                depth += 1
            elem = elem.parent
        if depth in [0, 1]:
            return depth
        else:
            return 2

    def action(self, elem, doc):
        if (doc.format == "docx"):
            # pf.debug(elem.__class__)
            if isinstance(elem, pf.BulletList):
                pf.debug("Bullet list found")
                bullets = [doc.get_metadata("bullet-style.1", "Bullet List 1"),
                           doc.get_metadata("bullet-style.2", "Bullet List 2"),
                           doc.get_metadata("bullet-style.3", "Bullet List 3"),
                           ]

                depth = self.get_depth(elem)
                # pf.debug(depth)
                converted = []

                for se in elem.content:
                    converted.extend(
                        [pf.Div(e, attributes={"custom-style": bullets[depth]}) for e in se.content if
                         not isinstance(e, pf.BulletList)])

                ret = []
                for content in converted:
                    while isinstance(content, pf.Div):
                        content = content.content[0]
                    ret.append(content.parent)
                # pf.debug(ret)
                return ret


def main(doc=None):
    uh = UnnumberHeadings()
    ifc = InlineFigureCentered()
    ebl = ExtractBulletList()
    s2p = Svg2Png()
    pf.run_filters([uh.action, ifc.action, ebl.action, s2p.action], doc=doc)
    return doc


def extract_bullet_list(doc=None):
    ebl = ExtractBulletList()
    pf.run_filter(ebl.action, doc=doc)


if __name__ == "__main__":
    main()
