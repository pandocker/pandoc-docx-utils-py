all: docx

docx:
	pandoc -t docx --reference-doc=ref.docx -F ../pandoc_docx_utils/__init__.py markdown.md -o docx.docx
