' PdfParser.cls
Private m_txt As PdfTXT
Private m_wrt As PdfWRT
Private m_ocr As WinOCR
Private m_wrd As WdCOM
Private m_sTier As String
Private m_lStatus As Long

Private Sub Class_Initialize()
    Set m_txt = New PdfTXT
    Set m_wrt = New PdfWRT
    Set m_ocr = New WinOCR
    Set m_wrd = New WdCOM
End Sub

Public Property Get LastTier() As String:   LastTier = m_sTier:   End Property
Public Property Get LastStatus() As Long:   LastStatus = m_lStatus: End Property

Public Property Get TXT() As PdfTXT: Set TXT = m_txt: End Property
Public Property Get WRT() As PdfWRT: Set WRT = m_wrt: End Property
Public Property Get OCR() As WinOCR: Set OCR = m_ocr: End Property
Public Property Get WRD() As WdCOM:  Set WRD = m_wrd: End Property

Public Function ExtractText(filePath As String) As String
    ' TODO: Implement text extraction logic
End Function

Public Function ExtractPages(filePath As String) As Collection
    ' TODO: Implement per-page extraction
End Function