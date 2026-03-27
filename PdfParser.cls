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
    
    If m_lStatus = m_txt.PDFTXT_OK Then
        m_sTier = "PdfTXT"
        ExtractText = sText
        Exit Function
    End If
    
    Dim oPages As Collection
    Set oPages = TryRenderAndOCR(filePath)
    
    If oPages.Count > 0 Then
        Dim vPage As Variant
        For Each vPage In oPages
            If Len(vPage) > 0 Then
                sText = sText & CStr(vPage) & vbCrLf
            End If
        Next
        
        If Len(Trim$(sText)) > 0 Then
            m_sTier = "OCR"
            m_lStatus = 0
            ExtractText = sText
            Exit Function
        End If
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
    
    If m_lStatus = m_txt.PDFTXT_OK Then
        m_sTier = "PdfTXT"
        oPages.Add sText
        Set ExtractPages = oPages
        Exit Function
    End If
    
    Dim oRenderedPages As Collection
    Set oRenderedPages = TryRenderAndOCR(filePath)
    
    If oRenderedPages.Count > 0 Then
        Dim vPage As Variant
        For Each vPage In oRenderedPages
            If Len(CStr(vPage)) > 0 Then
                oPages.Add vPage
            End If
        Next
        
        If oPages.Count > 0 Then
            m_sTier = "OCR"
            m_lStatus = 0
            Set ExtractPages = oPages
            Exit Function
        End If
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

Private Function TryRenderAndOCR(ByVal filePath As String) As Collection
    Set TryRenderAndOCR = New Collection
    
    Dim oPages As Collection
    Dim widths(2) As Long
    widths(0) = 0
    widths(1) = 1240
    widths(2) = 800
    
    Dim i As Long
    Dim vPage As Variant
    Dim sText As String
    
    For i = LBound(widths) To UBound(widths)
        Set oPages = m_wrt.RenderPDFToBytes(filePath, widths(i))
        
        If oPages.Count > 0 Then
            sText = ""
            For Each vPage In oPages
                Dim aPage() As Byte
                aPage = vPage
                Dim pageText As String
                pageText = m_ocr.BytesToText(aPage)
                If Len(pageText) > 0 Then
                    sText = sText & pageText & vbCrLf
                End If
            Next
            
            If Len(Trim$(sText)) > 0 Then
                Exit For
            End If
        End If
    Next
    
    If Len(Trim$(sText)) > 0 Then
        Dim allPagesText As String
        allPagesText = sText
        Dim pageStart As Long
        Dim pageEnd As Long
        Dim pageStr As String
        
        pageStart = 1
        Do While pageStart <= Len(allPagesText)
            pageEnd = InStr(pageStart, allPagesText, vbCrLf)
            If pageEnd = 0 Then pageEnd = Len(allPagesText) + 1
            pageStr = Mid$(allPagesText, pageStart, pageEnd - pageStart)
            If Len(pageStr) > 0 Then
                TryRenderAndOCR.Add pageStr
            End If
            pageStart = pageEnd + 2
        Loop
    End If
End Function