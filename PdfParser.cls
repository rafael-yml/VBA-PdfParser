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

Public Function ExtractText( _
    ByVal filePath As String, _
    Optional ByVal useWord As Boolean = False) As String

    m_sTier = ""
    m_lStatus = 0
    
    Dim sText As String
    sText = m_txt.ExtractText(filePath)
    m_lStatus = m_txt.LastStatus
    
    If m_lStatus = PDFTXT_OK Then
        m_sTier = "PdfTXT"
        ExtractText = sText
        Exit Function
    End If
    
    Dim oPages As Collection
    Set oPages = m_wrt.RenderPDFToBytes(filePath, 2480)
    
    If oPages.Count > 0 Then
        Dim vPage As Variant
        For Each vPage In oPages
            Dim aPage() As Byte
            aPage = vPage
            sText = sText & m_ocr.BytesToText(aPage) & vbCrLf
        Next
        m_sTier = "OCR"
        m_lStatus = 0
        ExtractText = sText
        Exit Function
    End If
    
    If useWord Then
        sText = m_wrd.ExtractText(filePath)
        If Len(sText) > 0 Then
            m_sTier = "WdCOM"
            m_lStatus = 0
            ExtractText = sText
            Exit Function
        End If
    End If
    
    ExtractText = ""
End Function

Public Function ExtractPages( _
    ByVal filePath As String, _
    Optional ByVal useWord As Boolean = False) As Collection

    m_sTier = ""
    m_lStatus = 0
    
    Dim oPages As New Collection
    Dim sText As String
    
    sText = m_txt.ExtractText(filePath)
    m_lStatus = m_txt.LastStatus
    
    If m_lStatus = PDFTXT_OK Then
        m_sTier = "PdfTXT"
        oPages.Add sText
        Set ExtractPages = oPages
        Exit Function
    End If
    
    Dim oRenderedPages As Collection
    Set oRenderedPages = m_wrt.RenderPDFToBytes(filePath, 2480)
    
    If oRenderedPages.Count > 0 Then
        Dim vPage As Variant
        For Each vPage In oRenderedPages
            Dim aPage() As Byte
            aPage = vPage
            oPages.Add m_ocr.BytesToText(aPage)
        Next
        m_sTier = "OCR"
        m_lStatus = 0
        Set ExtractPages = oPages
        Exit Function
    End If
    
    If useWord Then
        sText = m_wrd.ExtractText(filePath)
        If Len(sText) > 0 Then
            m_sTier = "WdCOM"
            m_lStatus = 0
            oPages.Add sText
            Set ExtractPages = oPages
            Exit Function
        End If
    End If
    
    Set ExtractPages = oPages
End Function