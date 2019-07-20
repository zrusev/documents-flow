Attribute VB_Name = "Core_Enums"
Option Explicit

Public Enum File_Formats
    wdFormatDocument = 0 'Microsoft Office Word 97 - 2003 binary file format.
    wdFormatDOSText = 4 'Microsoft DOS text format.
    wdFormatDOSTextLineBreaks = 5 'Microsoft DOS text with line breaks preserved.
    wdFormatEncodedText = 7 'Encoded text format.
    wdFormatFilteredHTML = 10 'Filtered HTML format.
    wdFormatFlatXML = 19 'Open XML file format saved as a single XML file.
    wdFormatFlatXMLMacroEnabled = 20 'Open XML file format with macros enabled saved as a single XML file.
    wdFormatFlatXMLTemplate = 21 'Open XML template format saved as a XML single file.
    wdFormatFlatXMLTemplateMacroEnabled = 22 'Open XML template format with macros enabled saved as a single XML file.
    wdFormatOpenDocumentText = 23 'OpenDocument Text format.
    wdFormatHTML = 8 'Standard HTML format.
    wdFormatRTF = 6 'Rich text format (RTF).
    wdFormatStrictOpenXMLDocument = 24 'Strict Open XML document format.
    wdFormatTemplate = 1 'Word template format.
    wdFormatText = 2 'Microsoft Windows text format.
    wdFormatTextLineBreaks = 3 'Windows text format with line breaks preserved.
    wdFormatUnicodeText = 7 'Unicode text format.
    wdFormatWebArchive = 9 'Web archive format.
    wdFormatXML = 11 'Extensible Markup Language (XML) format.
    wdFormatDocument97 = 0 'Microsoft Word 97 document format.
    wdFormatDocumentDefault = 16 'Word default document file format. For Word, this is the DOCX format.
    wdFormatPDF = 17 'PDF format.
    wdFormatTemplate97 = 1 'Word 97 template format.
    wdFormatXMLDocument = 12 'XML document format.
    wdFormatXMLDocumentMacroEnabled = 13 'XML document format with macros enabled.
    wdFormatXMLTemplate = 14 'XML template format.
    wdFormatXMLTemplateMacroEnabled = 15 'XML template format with macros enabled.
    wdFormatXPS = 18 'XPS format.
End Enum

Public Enum Break_Types
    wdColumnBreak = 8 'Column break at the insertion point.
    wdLineBreak = 6 'Line break.
    wdLineBreakClearLeft = 9 'Line break.
    wdLineBreakClearRight = 10 'Line break.
    wdPageBreak = 7 'Page break at the insertion point.
    wdSectionBreakContinuous = 3 'New section without a corresponding page break.
    wdSectionBreakEvenPage = 4 'Section break with the next section beginning on the next even-numbered page. If the section break falls on an even-numbered page, Word leaves the next odd-numbered page blank.
    wdSectionBreakNextPage = 2 'Section break on next page.
    wdSectionBreakOddPage = 5 'Section break with the next section beginning on the next odd-numbered page. If the section break falls on an odd-numbered page, Word leaves the next even-numbered page blank.
    wdTextWrappingBreak = 11 'Ends the current line and forces the text to continue below a picture, table, or other item. The text continues on the next blank line that does not contain a table aligned with the left or right margin.
End Enum

Public Enum ObjectTypes
    [_First] = 1
    Shape = 1
    TextBox = 2
    Slicer = 3
    DropDown = 4
    SpinButton = 5
    Table = 6
    Picture = 7
    MergedCells = 8
    [_Last] = 8
End Enum

Public Enum ShapeTypes
    msoShapeTypeMixed = -2   'Mixed shape type
    msoAutoShape = 1         'AutoShape.
    msoCallout = 2           'Callout.
    msoChart = 3             'Chart.
    msoComment = 4           'Comment.
    msoFreeform = 5          'Freeform.
    msoGroup = 6             'Group.
    msoEmbeddedOLEObject = 7 'Embedded OLE object.
    msoFormControl = 8       'Form control.
    msoLine = 9              'Line
    msoLinkedOLEObject = 10  'Linked OLE object
    msoLinkedPicture = 11    'Linked picture
    msoOLEControlObject = 12 'OLE control object
    msoPicture = 13          'Picture
    msoPlaceholder = 14      'Placeholder
    msoTextEffect = 15       'Text effect
    msoMedia = 16            'Media
    msoTextBox = 17          'Text box
    msoScriptAnchor = 18     'Script anchor
    msoTable = 19            'Table
    msoCanvas = 20           'Canvas.
    msoDiagram = 21          'Diagram.
    msoInk = 22              'Ink
    msoInkComment = 23       'Ink comment
    msoIgxGraphic = 24       'SmartArt graphic
    msoWebVideo = 26         'Web video
    msoContentApp = 27       'Content Office Add-in
    msoGraphic = 28          'Graphic
    msoLinkedGraphic = 29    'Linked graphic
End Enum

Public Enum CustomOperators
     Yes = True
     No = False
     Not_Applicable = True
End Enum

Public Enum VarTypes
    vbEmpty = 0 'Uninitialized (default)
    vbNull = 1 'Contains no valid data
    vbInteger = 2 'Integer
    vbLong = 3 'Long integer
    vbSingle = 4 'Single-precision floating-point number
    vbDouble = 5 'Double-precision floating-point number
    vbCurrency = 6 'Currency
    vbDate = 7 'Date
    vbString = 8 'String
    vbObject = 9 'Object
    vbError = 10 'Error
    vbBoolean = 11 'Boolean
    vbVariant = 12 'Variant (used only for arrays of variants)
    vbDataObject = 13 'Data access object
    vbDecimal = 14 'Decimal
    vbByte = 17 'Byte
    vbLongLong = 20 'LongLong integer (valid on 64-bit platforms only)
    vbUserDefinedType = 36 'Variants that contain user-defined types
    vbArray = 8192 'Array
End Enum
