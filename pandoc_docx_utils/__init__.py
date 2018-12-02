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
import panflute as pf


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

    unnumbered = {1: "Heading Unnumbered 1",
                  2: "Heading Unnumbered 2",
                  3: "Heading Unnumbered 3",
                  4: "Heading Unnumbered 4",
                  }

    def action(self, elem, doc):
        if (doc.format == "docx"):
            if isinstance(elem, pf.Header) and "unnumbered" in elem.classes:
                if elem.level < 5:
                    elem.attributes.update({"custom-style": self.unnumbered[elem.level]})
                    elem = pf.Div(pf.Para(*elem.content), attributes=elem.attributes,
                                  identifier=elem.identifier, classes=elem.classes)
                    # pf.debug(elem)
        return elem


class InlineFigureCentered(object):
    """
    moves a Para containing an Image into to "Centered" custom style div
    * "Centered" custom style must be prepared in template
    """

    def action(self, elem, doc):
        if (doc.format == "docx"):
            default_style = doc.get_metadata("image-div-style", "Image Div")
            if isinstance(elem, pf.Para) and len(elem.content) == 1:
                for subelem in elem.content:
                    if isinstance(subelem, pf.Image):
                        style = subelem.attributes.get("custom-style", default_style)
                        subelem.attributes["custom-style"] = style
                        d = pf.Div(elem, attributes=subelem.attributes)
                        return d


def main(doc=None):
    uh = UnnumberHeadings()
    ifc = InlineFigureCentered()
    pf.run_filters([uh.action, ifc.action], doc=doc)
    return doc


if __name__ == "__main__":
    main()
