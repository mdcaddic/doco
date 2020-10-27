This doco repository will enable testing of using markdown documents as sections of a greater document and then convert that to docx format using pandoc

https://pandoc.org/MANUAL.html

conversion syntax
pandoc -f markdown -s RBAC.md -o RBAC.docx --reference-doc=oobe.docx

Word document template

--reference-doc=FILE
Use the specified file as a style reference in producing a docx or ODT file.

Docx
For best results, the reference docx should be a modified version of a docx file produced using pandoc. The contents of the reference docx are ignored, but its stylesheets and document properties (including margins, page size, header, and footer) are used in the new docx. If no reference docx is specified on the command line, pandoc will look for a file reference.docx in the user data directory (see --data-dir). If this is not found either, sensible defaults will be used.

To produce a custom reference.docx, first get a copy of the default reference.docx: pandoc -o custom-reference.docx --print-default-data-file reference.docx. Then open custom-reference.docx in Word, modify the styles as you wish, and save the file. For best results, do not make changes to this file other than modifying the styles used by pandoc:

Paragraph styles:

* Normal
* Body Text
* First Paragraph
* Compact
* Title
* Subtitle
* Author
* Date
* Abstract
* Bibliography
* Heading 1
* Heading 2
* Heading 3
* Heading 4
* Heading 5
* Heading 6
* Heading 7
* Heading 8
* Heading 9
* Block Text
* Footnote Text
* Definition Term
* Definition
* Caption
* Table Caption
* Image Caption
* Figure
* Captioned Figure
* TOC Heading

Character styles:
* Default Paragraph Font
* Body Text Char
* Verbatim Char
* Footnote Reference
* Hyperlink
* Section Number
  
Table style:
* Table
