#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from pandoc_docx_utils import ExtractBulletList
import panflute as pf


def main():
    with open("markdown.md", "r") as f:
        md = f.read()
    doc = pf.Doc(*pf.convert_text(md), format="docx")
    pf.debug("doc: {}".format(*doc.content))
    ebl = ExtractBulletList()
    pf.run_filter(ebl.action, doc=doc)


if __name__ == "__main__":
    main()
