# RPA_SAP
Python module delivers some actions to manipulate with PDF files.
The module is compatibile with the Robocorp.

## Installation
To install the package run:

```
pip install rpa-sap
```

## Example
### Generating pdf file from text
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.text_to_pdf('some text', 'c:/temp/somefile.pdf)
```
### Compress pdf file
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.compress('c:/temp/somefile.pdf')
```
### Add Code39 stamp
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.add_code39_stamp('c:/temp/input_file.pdf', 'c:/temp/output_file.pdf', '12345678')
```
### Merging pdf files
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.merge(['c:/temp/file1.pdf', 'c:/temp/file2.pdf'], 'c:/temp/merged.pdf')
```
### Add text stamp
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.add_text_stamp('c:/temp/input_file.pdf', 'c:/temp/output_file.pdf', 'some text')
```
### Print PDF document on the printer
```
from rpa_pdf import Pdf

pdf = Pdf()

pdf.print('c:/temp/document.pdf', 'printer1')
```

### Dependencies
Python packages: PyPDF2 >= 1.28.5, fpdf2 >= 2.5.6, python-barcode >= 0.14.0
External dependencies: DejaVu font, SumatraPdf.exe (both included in the package)
