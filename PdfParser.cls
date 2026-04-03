VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PdfParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

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
' ---------------------------------------------------------------------------
' Status codes
' ---------------------------------------------------------------------------
Private Const PDFTXT_OK      As Long = 0  ' clean text returned
Private Const PDFTXT_NO_TEXT As Long = 1  ' no text operators found (image-only or blank PDF)
Private Const PDFTXT_NO_CMAP As Long = 2  ' hex-encoded glyphs found but no ToUnicode CMap
Private Const PDFTXT_GARBLED As Long = 3  ' CMap present but unmapped-glyph ratio too high
Private Const PDFTXT_FAIL    As Long = 4  ' file missing, not a PDF, encrypted, or parse error

Private Const WDCOM_OK           As Long = 0
Private Const WDCOM_FILE_MISSING As Long = 1
Private Const WDCOM_ENCRYPTED    As Long = 2
Private Const WDCOM_WORD_FAILED  As Long = 3
Private Const WDCOM_OPEN_FAILED  As Long = 4
Private Const WDCOM_EMPTY        As Long = 5
Private Const WDCOM_GARBLED      As Long = 6

Private Const WINOCR_OK        As Long = 0  ' text recognised successfully
Private Const WINOCR_NO_ENGINE As Long = 1  ' language not installed / engine unavailable
Private Const WINOCR_TOO_LARGE As Long = 2  ' image exceeds OcrEngine.MaxImageDimension
Private Const WINOCR_EMPTY     As Long = 3  ' OCR completed but returned no text
Private Const WINOCR_FAIL      As Long = 4  ' file missing, decode error, or init failure

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
        m_sTier = "OCR"
        m_lStatus = 0
        ExtractText = JoinCollection(oPages, vbCrLf)
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
    
    If m_lStatus = m_txt.PDFTXT_OK Then
        m_sTier = "PdfTXT"
        oPages.Add sText
        Set ExtractPages = oPages
        Exit Function
    End If
    
    Dim oRenderedPages As Collection
    Set oRenderedPages = TryRenderAndOCR(filePath)
    
    If oRenderedPages.Count > 0 Then
        m_sTier = "OCR"
        m_lStatus = 0
        Set ExtractPages = oRenderedPages
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

Private Function TryRenderAndOCR(ByVal filePath As String) As Collection
    Set TryRenderAndOCR = New Collection
    
    Dim widths(2) As Long
    widths(0) = 0
    widths(1) = 1240
    widths(2) = 800
    
    Dim i As Long
    Dim oRenderedPages As Collection
    Dim vPage As Variant
    Dim pageText As String
    Dim foundText As Boolean
    
    For i = LBound(widths) To UBound(widths)
        foundText = False
        Set oRenderedPages = m_wrt.RenderPDFToBytes(filePath, widths(i))
        
        If oRenderedPages.Count > 0 Then
            For Each vPage In oRenderedPages
                Dim aPage() As Byte
                aPage = vPage
                pageText = m_ocr.BytesToText(aPage)
                If Len(pageText) > 0 Then
                    TryRenderAndOCR.Add pageText
                    foundText = True
                End If
            Next
            
            If foundText Then
                Exit For
            End If
        End If
    Next
End Function

Private Function JoinCollection(ByVal oCol As Collection, ByVal delimiter As String) As String
    JoinCollection = ""
    Dim v As Variant
    For Each v In oCol
        If Len(JoinCollection) > 0 Then
            JoinCollection = JoinCollection & delimiter
        End If
        JoinCollection = JoinCollection & CStr(v)
    Next
End Function