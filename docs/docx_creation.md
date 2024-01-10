# Create DOCX file

Microsoft Word files are a standard for exchanging rich text documents. Images, tables and lists, among other types of resources, can be included in them. In addition, it is a format that the rector can easily edit, unlike PDF files. So having the possibility of creating Word documents with Python can make it easier for us to create reports in this format.

## Install libraries

To handle word documents we need to install the library docx.

```bash
pip install python-docx
```

## Create word file

### Import libraries

Here import the libraries needed.

```python
from docx import Document
from docx.shared import Cm, Pt
```

### Create Microsoft word document

Follow we need to add the document object to code, after add this line you able to add header, title, subtitle, images, text, and styles to word document.

#### Create main object

```python
doc = Document()
```

#### Add header to file

With the next code add the header to file.

```python
# Header
header = doc.sections[0].header
paragraph = header.paragraphs[0]
paragraph.add_run("Retrogemu docx file")
```

#### Add title to file

Like you see in the next part of the code we able to insert a title inside the file and set the title font size.

```python
# Title
title = doc.add_heading("Title 1", level = 1)
title_font = title.runs[0].font
title_font.size = Pt(18)
```

#### Add paragraph to file

With follow code add a paragraph to file.

```python
# Paragraph
paragraph_1 = doc.add_paragraph("Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip")
```

#### Add bold text

if you need to remark some text, the follow code would help you to do that.

```python
# Bold text
paragraph_1.add_run("ex ea commodo consequat.").bold = True
paragraph_1.add_run("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.")
```

#### Add subtitle and image

Also we able to add subtitle and images in the document.

```python
# Subtitle
subtitle_1 = doc.add_heading("Profile image", level = 2)
doc.add_paragraph()
doc.add_picture('assets/img/profile_image.png', width=Cm(2), height=Cm(2))
paragraph_2 = doc.add_paragraph()
paragraph_2.add_run("@Retrogemu").bold = True
```

#### Save file

As last step in this code, is to save the file.

```python
# Save file
doc.save("profile_example.docx")
```

## Complete file

Now here is the complete code in one file, don't forget put the image in the folder **assets\img\profile_image.png** to get from code.

```python
from docx import Document
from docx.shared import Cm, Pt

doc = Document()

# Header
header = doc.sections[0].header
paragraph = header.paragraphs[0]
paragraph.add_run("Retrogemu docx file")

# Title
title = doc.add_heading("Title 1", level=1)
title_font = title.runs[0].font
title_font.size = Pt(18)

# Paragraph
paragraph_1 = doc.add_paragraph(
    "Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip")

# Bold text
paragraph_1.add_run(" ex ea commodo consequat.").bold = True
paragraph_1.add_run("Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.")

# Subtitle
subtitle_1 = doc.add_heading("Profile image", level=2)
doc.add_paragraph()
doc.add_picture('assets/img/profile_image.png', width=Cm(2), height=Cm(2))
paragraph_2 = doc.add_paragraph()
paragraph_2.add_run("@Retrogemu").bold = True

# Save file
doc.save("profile_example.docx")
```

## Result