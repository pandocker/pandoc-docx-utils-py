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

    def action(self, elem, doc):
        if (doc.format == "docx"):
            if isinstance(elem, pf.BulletList):
                p = elem
                depth = 0
                pf.debug("elem: {}".format(elem))
                while not isinstance(p, pf.Doc):
                    if isinstance(p, pf.ListItem):
                        pf.debug("content: {}".format(*p.content))
                        depth += 1
                    p = p.parent
                pf.debug(depth)

                # if depth >= 2:
                #     d = pf.Div(*elem.content, attributes={"custom-style": "Bullet List 2"})
                # else:
                #     d = pf.Div(*elem.content, attributes={"custom-classes": "Bullet List 1"})
                # pf.debug(d.attributes)
                return []


def main(doc=None):
    uh = UnnumberHeadings()
    ifc = InlineFigureCentered()
    pf.run_filters([uh.action, ifc.action], doc=doc)
    return doc


if __name__ == "__main__":
    main()
