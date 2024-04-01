# wordoffice

It is a python library to only write Microsoft Word (docx) files.

## Examples

```python
from wordoffice.docx.document import Docx
from wordoffice.ooxml.run import Run

docx = Docx()

run = Run('Hello, World!', rPr={'bold': True, 'color': 'FF0000'})
docx.add_p(run)

docx.save('hello.docx')
```