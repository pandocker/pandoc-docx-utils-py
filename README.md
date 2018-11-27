# pandoc-docx-utils

Experimental set of pandoc filters to output DOCX file

## unnumbered headings

* Template is set up to use numbered Headings while also required to use unnumbered ones
* `unnumbered` class does not work for docx output
* `UnnumberHeadings` class works to _unnumber_ Headings
* Limited to level-1 to 4 headings
* Requires `Heading Unnumbered 1` to `Heading Unnumbered 4` heading styles in template otherwise
  these headers appear as in `Body` style

## centering images

* an image link in paragraph will be centered
* blank lines required before and after image link
* Requires `Centered` paragraph style in template otherwise no effect can be seen
* Can change style afterwards

# samples

![image](qr.png){width=100mm #fig:centered}

# another heading {.unnumbered}

a paragraph
