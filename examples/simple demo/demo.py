from wordoffice.common import Measure
from wordoffice.docx.document import Docx
from wordoffice.docx.image import Image, ImagePart, ImageInline, ImageAnchor

from wordoffice.ooxml.paragraph import Paragraph
from wordoffice.ooxml.run import Run
from wordoffice.ooxml.section import Header, HeaderReference
from wordoffice.ooxml.styles import StylesBuiltinPart, Style
from wordoffice.docx.table import Table
from wordoffice.utils import cn_font

# Create a new document
docx = Docx()

# Basic usage
run = Run('Hello, World!', rPr={'bold': True, 'color': 'FF0000'})
paragraph = Paragraph(run)
docx.add_p(paragraph)

# Use style
style_normal_c = {'fonts': cn_font('SimSun'), 'fontSize': Measure.PT(12), 'bold': False}
style_color_c = {'fonts': cn_font('SimSun'), 'fontSize': Measure.PT(24), 'bold': False, 'color': 'FF0000'}
style_center_p = {'jc': 'center'}

# Use style: add default style and custom style
styles = StylesBuiltinPart()
styles.docDefaults.rPrDefault = {'rPr': style_normal_c}

styles.character('normal_c', rPr=style_normal_c)
styles.character('color_c', rPr=style_color_c)
styles.paragraph('center_p', pPr=style_center_p)

# Use Style: add style to paragraph and run
docx.add_p(Run('Hello, Normal!', 'normal_c'))
docx.add_p(Run('Hello, Color!', 'color_c'))
docx.add_p(Paragraph('Hello, Center!', 'center_p'))

# Table, you can add border in style
table = Table(size=(5, 5))
table.merge_cell((0, 0), (0, 3))
table.merge_cell((1, 0), (3, 0))
table.merge_cell((2, 2), (3, 3))

count = 0
for rows in table:
	for cell in rows:
		count += 1
		cell.add_p(str(count))

docx.add_table(table)

# Header and Footer
headerPart = Header(word=docx)
headerRef = HeaderReference.DEFAULT(headerPart)
docx.lastSection.sectPr.setHeaderReference(headerRef)

headerPart.add_p('this is Header')

# Image
imagePart = ImagePart('./test.jpg', word=docx)

image_inline = ImageInline(imagePart)
image_inline.size = (Measure.PX(300), Measure.PX(150))

run_image_inline = Run()
run_image_inline.add_image(image_inline)
docx.add_p(['inline', run_image_inline])

image_anchor = ImageAnchor(imagePart)
image_anchor.overlap = True
image_anchor.behindDoc = True
image_anchor.size = (Measure.PX(300), Measure.PX(150))
image_anchor.position = (
	{'relativeFrom': 'page', 'posOffset': Measure.CM(11)},
	{'relativeFrom': 'page', 'posOffset': Measure.CM(1)},
)

run_image_anchor = Run()
run_image_anchor.add_image(image_anchor)
docx.add_p(['anchor, the top image', run_image_anchor])

docx.set_template('styles', styles)
# Save to file
docx.save('hello.docx')
