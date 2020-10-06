# WORT - *Wor*d *T*ransformer

## What does it do

The WORd Transformer tool allows you to drag/drop files onto a batch file that will transform your queue into a format of your choosing.

## Requirements

- You must have Microsoft Word installed on your machine (will not work with cloud solutions)
- Input files must be of a format that MS Word can open
- Output files can only be in a format that MS Word can open

## Useage

Select all of the files you wish to convert and drag & drop them onto the `.bat` file. This will spawn a Command Window that will ask you to enter an ID of the format you want for the output file (see Output Format IDs) followed by the `ENTER` key. If you do not specify an ID and hit enter, the tool will default to an ID of 16 which will produce "wdFormatDocumentDefault" `.DOCX` files as this is the most common use of the tool at the time of development.

## Valid Output Format IDs

Table taken from [Microsoft Docs](https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveformat)

| Name                                | Value  | Description                                                               |
| ----------------------------------- | ------ | ------------------------------------------------------------------------- |
| wdFormatDocument                    | 0      | Microsoft Office Word 97 - 2003 binary file format.                       |
| wdFormatDOSText                     | 4      | Microsoft DOS text format.                                                |
| wdFormatDOSTextLineBreaks           | 5      | Microsoft DOS text with line breaks preserved.                            |
| wdFormatEncodedText                 | 7      | Encoded text format.                                                      |
| **wdFormatFilteredHTML**            | **10** | **Filtered HTML format.**                                                 |
| wdFormatFlatXML                     | 19     | Open XML file format saved as a single XML file.                          |
| wdFormatFlatXML                     | 20     | Open XML file format with macros enabled saved as a single XML file.      |
| wdFormatFlatXMLTemplate             | 21     | Open XML template format saved as a XML single file.                      |
| wdFormatFlatXMLTemplateMacroEnabled | 22     | Open XML template format with macros enabled saved as a single XML file.  |
| wdFormatOpenDocumentText            | 23     | OpenDocument Text format.                                                 |
| wdFormatHTML                        | 8      | Standard HTML format.                                                     |
| wdFormatRTF                         | 6      | Rich text format (RTF).                                                   |
| wdFormatStrictOpenXMLDocument       | 24     | Strict Open XML document format.                                          |
| wdFormatTemplate                    | 1      | Word template format.                                                     |
| wdFormatText                        | 2      | Microsoft Windows text format.                                            |
| wdFormatTextLineBreaks              | 3      | Windows text format with line breaks preserved.                           |
| wdFormatUnicodeText                 | 7      | Unicode text format.                                                      |
| wdFormatWebArchive                  | 9      | Web archive format.                                                       |
| wdFormatXML                         | 11     | Extensible Markup Language (XML) format.                                  |
| wdFormatDocument97                  | 0      | Microsoft Word 97 document format.                                        |
| **wdFormatDocumentDefault**         | **16** | **Word default document file format. For Word, this is the DOCX format.** |
| **wdFormatPDF**                     | **17** | **PDF format.**                                                           |
| wdFormatTemplate97                  | 1      | Word 97 template format.                                                  |
| wdFormatXMLDocument                 | 12     | XML document format.                                                      |
| wdFormatXMLDocumentMacroEnabled     | 13     | XML document format with macros enabled.                                  |
| wdFormatXMLTemplate                 | 14     | XML template format.                                                      |
| wdFormatXMLTemplateMacroEnabled     | 15     | XML template format with macros enabled.                                  |
| wdFormatXPS                         | 18     | XPS format.                                                               |
