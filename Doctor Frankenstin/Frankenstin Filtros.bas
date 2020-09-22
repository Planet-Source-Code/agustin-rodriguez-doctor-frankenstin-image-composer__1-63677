Attribute VB_Name = "mFilterG"
'+--------------------------------------------------------+
'| Name            : mFilterG - Graphic Filters           |
'| Author          : Manuel Augusto Nogueira dos Santos   |
'| Dates           : 23/03/2001                           |
'| Description     : Apply effects to images              |
'| Notes           : If you include this module in your   |
'|                   program, include me in the About box |
'+--------------------------------------------------------+
'| FilterG(ByVal Filtro As iFilterG                       |
'|               > one of iFilterG Enum                   |
'|         ByVal Pic As Long,                             |
'|               > PictureBox.Image                       |
'|         ByVal Factor As Long,                          |
'|               > depends upon Filtro (see below)        |
'|         ByRef pProgress As Long)                       |
'|               > % progress done                        |
'+--------------------------------------------------------+
'| Factor                                                 |
'|  iSHARPEN    : 0..N for Sharpen + to Sharpen -         |
'|  iNEGATIVE   : no effect                               |
'|  iBLUR       : no effect                               |
'|  iDIFFUSE    : diffuse radius, 6 normal / 12 diffuse + |
'|  iSMOOTH     : no effect                               |
'|  iEDGE       : 1..N for EdgeEnhance + to EdgeEnhance - |
'|  iCONTOUR    : RGB BackColor                           |
'|  iEMBOSS     : RGB BackColor                           |
'|  iEMBOSSMORE : RGB BackColor                           |
'|  iENGRAVE    : RGB BackColor                           |
'|  iENGRAVEMORE: RGB BackColor                           |
'|  iGREYSCALE  : no effect                               |
'|  iRELIEF     : no effect                               |
'|  iBRIGHTNESS : >0 to increase, <0 to decrease          |
'|  iPIXELIZE   : size of each pixel                      |
'|  iSWAPBANK   : 1..5 RGB to (BRG,GBR,RBG,BGR,GRB)       |
'|  iCONTRAST   : >0 to increase, <0 to decrease          |
'|  iCOLDEPTH1  : RGB color to set black below            |
'|  iCOLDEPTH2  : no effect                               |
'|  iCOLDEPTH3  : no effect                               |
'|  iCOLDEPTH4  : 1..n Palette colors weight              |
'|  iCOLDEPTH5  : 1..n Palette colors weight              |
'|  iCOLDEPTH6  : 1..n Palette colors weight              |
'|  iAQUA       : no effect                               |
'|  iDILATE     : no effect                               |
'|  iERODE      : no effect                               |
'|  iCONNECTION : no effect                               |
'|  iSTRETCH    : no effect                               |
'|  iADDNOISE   : noise intensity                         |
'|  iSATURATION : >0 to increase, <0 to decrease          |
'+--------------------------------------------------------+
Option Explicit

'-------------------------------------------Windows API
Public Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetDIBits Lib "GDI32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function LoadImage Lib "User32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function GetObject Lib "GDI32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetDIBits Lib "GDI32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

'-------------------------------------------Public var
Public WorkFilterG As Boolean

Public Const SRCCOPY = &HCC0020

Public Enum iFilterG
  iSHARPEN = 1
  iNEGATIVE = 2
  iBLUR = 3
  iDIFFUSE = 4
  iSMOOTH = 5
  iEDGE = 6
  iCONTOUR = 7
  iEMBOSS = 8
  iEMBOSSMORE = 9
  iENGRAVE = 10
  iENGRAVEMORE = 11
  iGREYSCALE = 12
  iRELIEF = 13
  iBRIGHTNESS = 14
  iPIXELIZE = 15
  iSWAPBANK = 16
  iCONTRAST = 17
  iCOLDEPTH1 = 18
  iCOLDEPTH2 = 19
  iCOLDEPTH3 = 20
  iCOLDEPTH4 = 21
  iCOLDEPTH5 = 22
  iCOLDEPTH6 = 23
  iAQUA = 24
  iDILATE = 25
  iERODE = 26
  iCONNECTION = 27
  iSTRETCH = 28
  iADDNOISE = 29
  iSATURATION = 30
  iGAMMA = 31
End Enum

'-------------------------------------------Private var
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0&
Private Const LR_LOADFROMFILE = &H10
Private Const IMAGE_BITMAP = 0&

Private iDATA() As Byte           'holds bitmap data
Private bDATA() As Byte           'holds bitmap backup
Private PicInfo As BITMAP         'bitmap info structure
Private DIBInfo As BITMAPINFO     'Device Ind. Bitmap info structure
Public mProgress As Long         '% filter progress
Private Speed(0 To 765) As Long   'Speed up values

Private Type BITMAP
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Type BITMAPINFOHEADER   '40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type

Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbReserved As Byte
End Type

Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER
  bmiColors As RGBQUAD
End Type

Public Sub FilterG(ByVal Filtro As iFilterG, ByVal Pic As Long, ByVal Factor As Long, ByRef pProgress As Long)
  Dim hdcNew As Long
  Dim oldhand As Long
  Dim Ret As Long
  Dim BytesPerScanLine As Long
  Dim PadBytesPerScanLine As Long
    
  If WorkFilterG = True Then Exit Sub
  WorkFilterG = True
  'On Error GoTo FilterError:
  'get data buffer
  Call GetObject(Pic, Len(PicInfo), PicInfo)
  hdcNew = CreateCompatibleDC(0&)
  oldhand = SelectObject(hdcNew, Pic)
  With DIBInfo.bmiHeader
    .biSize = 40
    .biWidth = PicInfo.bmWidth
    .biHeight = -PicInfo.bmHeight     'bottom up scan line is now inverted
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_RGB
    BytesPerScanLine = ((((.biWidth * .biBitCount) + 31) \ 32) * 4)
    PadBytesPerScanLine = _
       BytesPerScanLine - (((.biWidth * .biBitCount) + 7) \ 8)
    .biSizeImage = BytesPerScanLine * Abs(.biHeight)
  End With
  'redimension  (BGR+pad,x,y)
  ReDim iDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  ReDim bDATA(1 To 4, 1 To PicInfo.bmWidth, 1 To PicInfo.bmHeight) As Byte
  'get bytes
  Ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  Ret = GetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, bDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  'do it
  Select Case Filtro
    Case iSHARPEN:     Call Sharpen(pProgress, Factor)
    Case iNEGATIVE:    Call NegativeImage(pProgress)
    Case iBLUR:        Call Blurs(pProgress)
    Case iDIFFUSE:     Call Diffuse(pProgress, Factor)
    Case iSMOOTH:      Call Smooth(pProgress)
    Case iEDGE:        Call EdgeEnhance(pProgress, Factor)
    Case iCONTOUR:     Call Contour(pProgress, Factor)
    Case iEMBOSS:      Call Emboss(pProgress, Factor)
    Case iEMBOSSMORE:  Call EmbossMore(pProgress, Factor)
    Case iENGRAVE:     Call Engrave(pProgress, Factor)
    Case iENGRAVEMORE: Call EngraveMore(pProgress, Factor)
    Case iGREYSCALE:   Call GreyScale(pProgress)
    Case iRELIEF:      Call Relief(pProgress)
    Case iBRIGHTNESS:  Call Brightness(pProgress, Factor)
    Case iPIXELIZE:    Call Pixelize(pProgress, Factor)
    Case iSWAPBANK:    Call SwapBank(pProgress, Factor)
    Case iCONTRAST:    Call Contrast(pProgress, Factor)
    Case iCOLDEPTH1:   Call NearestColorBW(pProgress, Factor)
    Case iCOLDEPTH2:   Call EnhancedDiffusionBW(pProgress)
    Case iCOLDEPTH3:   Call OrderedDitherBW(pProgress)
    Case iCOLDEPTH4:   Call FloydSteinbergBW(pProgress, Factor)
    Case iCOLDEPTH5:   Call BurkeBW(pProgress, Factor)
    Case iCOLDEPTH6:   Call StuckiBW(pProgress, Factor)
    Case iAQUA:        Call Aqua(pProgress)
    Case iDILATE:      Call Dilate(pProgress)
    Case iERODE:       Call Erode(pProgress)
    Case iCONNECTION:  Call ConnectedContour(pProgress)
    Case iSTRETCH:     Call StretchHistogram(pProgress)
    Case iADDNOISE:    Call AddNoise(pProgress, Factor)
    Case iSATURATION:  Call Saturation(pProgress, Factor)
    Case iGAMMA:       Call GammaCorrection(pProgress, Factor)
  End Select
  'copy bytes to device
  Ret = SetDIBits(hdcNew, Pic, 0, PicInfo.bmHeight, iDATA(1, 1, 1), DIBInfo, DIB_RGB_COLORS)
  SelectObject hdcNew, oldhand
  DeleteDC hdcNew
  ReDim iDATA(1 To 4, 1 To 2, 1 To 2) As Byte
  ReDim bDATA(1 To 4, 1 To 2, 1 To 2) As Byte
  WorkFilterG = False
  Exit Sub
FilterError:
  MsgBox "Filter Error"
  WorkFilterG = False
End Sub

'-------------------------------------------AUXILIARY
Private Sub GetRGB(ByVal col As Long, ByRef r As Long, ByRef G As Long, ByRef B As Long)
  r = col Mod 256
  G = ((col And &HFF00&) \ 256&) Mod 256&
  B = (col And &HFF0000) \ 65536
End Sub

'-------------------------------------------FILTERS

Private Sub NegativeImage(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      iDATA(1, x, Y) = 255 - iDATA(1, x, Y)
      iDATA(2, x, Y) = 255 - iDATA(2, x, Y)
      iDATA(3, x, Y) = 255 - iDATA(3, x, Y)
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Sharpen(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim mf As Long, dF As Long
  
  mProgress = 0
  mf = 24 + Factor
  dF = 8 + Factor
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, x, Y - 1)) + CLng(iDATA(1, x - 1, Y)) + _
          CLng(iDATA(1, x + 1, Y)) + CLng(iDATA(1, x, Y + 1)) + _
          CLng(iDATA(1, x + 1, Y + 1)) + CLng(iDATA(1, x - 1, Y + 1)) + _
          CLng(iDATA(1, x + 1, Y - 1)) + CLng(iDATA(1, x - 1, Y - 1))
      B = (mf * CLng(iDATA(1, x, Y)) - 2 * B) \ dF
      G = CLng(iDATA(2, x, Y - 1)) + CLng(iDATA(2, x - 1, Y)) + _
          CLng(iDATA(2, x + 1, Y)) + CLng(iDATA(2, x, Y + 1)) + _
          CLng(iDATA(2, x + 1, Y + 1)) + CLng(iDATA(2, x - 1, Y + 1)) + _
          CLng(iDATA(2, x + 1, Y - 1)) + CLng(iDATA(2, x - 1, Y - 1))
      G = (mf * CLng(iDATA(2, x, Y)) - 2 * G) \ dF
      r = CLng(iDATA(3, x, Y - 1)) + CLng(iDATA(3, x - 1, Y)) + _
          CLng(iDATA(3, x + 1, Y)) + CLng(iDATA(3, x, Y + 1)) + _
          CLng(iDATA(3, x + 1, Y + 1)) + CLng(iDATA(3, x - 1, Y + 1)) + _
          CLng(iDATA(3, x + 1, Y - 1)) + CLng(iDATA(3, x - 1, Y - 1))
      r = (mf * CLng(iDATA(3, x, Y)) - 2 * r) \ dF
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Engrave(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight - 1
    For x = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, x + 1, Y + 1)) - CLng(iDATA(1, x, Y)) + cB)
      G = Abs(CLng(iDATA(2, x + 1, Y + 1)) - CLng(iDATA(2, x, Y)) + cG)
      r = Abs(CLng(iDATA(3, x + 1, Y + 1)) - CLng(iDATA(3, x, Y)) + cR)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub EngraveMore(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(bDATA(1, x + 1, Y - 1)) - CLng(bDATA(1, x - 1, Y - 1)) + _
          CLng(bDATA(1, x + 1, Y)) - CLng(bDATA(1, x - 1, Y)) + _
          CLng(bDATA(1, x + 1, Y + 1)) - CLng(bDATA(1, x - 1, Y + 1)) + cB
      G = CLng(bDATA(2, x + 1, Y - 1)) - CLng(bDATA(2, x - 1, Y - 1)) + _
          CLng(bDATA(2, x + 1, Y)) - CLng(bDATA(2, x - 1, Y)) + _
          CLng(bDATA(2, x + 1, Y + 1)) - CLng(bDATA(2, x - 1, Y + 1)) + cG
      r = CLng(bDATA(3, x + 1, Y - 1)) - CLng(bDATA(3, x - 1, Y - 1)) + _
          CLng(bDATA(3, x + 1, Y)) - CLng(bDATA(3, x - 1, Y)) + _
          CLng(bDATA(3, x + 1, Y + 1)) - CLng(bDATA(3, x - 1, Y + 1)) + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Emboss(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 1 To PicInfo.bmHeight - 1
    For x = 1 To PicInfo.bmWidth - 1
      B = Abs(CLng(iDATA(1, x, Y)) - CLng(iDATA(1, x + 1, Y + 1)) + cB)
      G = Abs(CLng(iDATA(2, x, Y)) - CLng(iDATA(2, x + 1, Y + 1)) + cG)
      r = Abs(CLng(iDATA(3, x, Y)) - CLng(iDATA(3, x + 1, Y + 1)) + cR)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub EmbossMore(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(bDATA(1, x - 1, Y - 1)) - CLng(bDATA(1, x + 1, Y - 1)) + _
          CLng(bDATA(1, x - 1, Y)) - CLng(bDATA(1, x + 1, Y)) + _
          CLng(bDATA(1, x - 1, Y + 1)) - CLng(bDATA(1, x + 1, Y + 1)) + cB
      G = CLng(bDATA(2, x - 1, Y - 1)) - CLng(bDATA(2, x + 1, Y - 1)) + _
          CLng(bDATA(2, x - 1, Y)) - CLng(bDATA(2, x + 1, Y)) + _
          CLng(bDATA(2, x - 1, Y + 1)) - CLng(bDATA(2, x + 1, Y + 1)) + cG
      r = CLng(bDATA(3, x - 1, Y - 1)) - CLng(bDATA(3, x + 1, Y - 1)) + _
          CLng(bDATA(3, x - 1, Y)) - CLng(bDATA(3, x + 1, Y)) + _
          CLng(bDATA(3, x - 1, Y + 1)) - CLng(bDATA(3, x + 1, Y + 1)) + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Diffuse(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim aX As Long, aY As Long
  Dim r As Long, G As Long, B As Long
  Dim hF As Long

  mProgress = 0
  hF = Factor / 2
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      aX = Rnd * Factor - hF
      aY = Rnd * Factor - hF
      If x + aX < 1 Then aX = 0
      If x + aX > PicInfo.bmWidth Then aX = 0
      If Y + aY < 1 Then aY = 0
      If Y + aY > PicInfo.bmHeight Then aY = 0
      iDATA(1, x, Y) = iDATA(1, x + aX, Y + aY)
      iDATA(2, x, Y) = iDATA(2, x + aX, Y + aY)
      iDATA(3, x, Y) = iDATA(3, x + aX, Y + aY)
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Smooth(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long

  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, x, Y)) + _
        CLng(iDATA(1, x - 1, Y)) + CLng(iDATA(1, x, Y - 1)) + _
        CLng(iDATA(1, x, Y + 1)) + CLng(iDATA(1, x + 1, Y))
      B = B \ 5
      G = CLng(iDATA(2, x, Y)) + _
        CLng(iDATA(2, x - 1, Y)) + CLng(iDATA(2, x, Y - 1)) + _
        CLng(iDATA(2, x, Y + 1)) + CLng(iDATA(2, x + 1, Y))
      G = G \ 5
      r = CLng(iDATA(3, x, Y)) + _
        CLng(iDATA(3, x - 1, Y)) + CLng(iDATA(3, x, Y - 1)) + _
        CLng(iDATA(3, x, Y + 1)) + CLng(iDATA(3, x + 1, Y))
      r = r \ 5
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Blurs(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long

  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, x - 1, Y - 1)) + CLng(iDATA(1, x - 1, Y)) + _
        CLng(iDATA(1, x - 1, Y + 1)) + CLng(iDATA(1, x, Y - 1)) + _
        CLng(iDATA(1, x, Y + 1)) + CLng(iDATA(1, x + 1, Y - 1)) + _
        CLng(iDATA(1, x + 1, Y)) + CLng(iDATA(1, x + 1, Y + 1))
      B = B \ 8
      G = CLng(iDATA(2, x - 1, Y - 1)) + CLng(iDATA(2, x - 1, Y)) + _
        CLng(iDATA(2, x - 1, Y + 1)) + CLng(iDATA(2, x, Y - 1)) + _
        CLng(iDATA(2, x, Y + 1)) + CLng(iDATA(2, x + 1, Y - 1)) + _
        CLng(iDATA(2, x + 1, Y)) + CLng(iDATA(2, x + 1, Y + 1))
      G = G \ 8
      r = CLng(iDATA(3, x - 1, Y - 1)) + CLng(iDATA(3, x - 1, Y)) + _
        CLng(iDATA(3, x - 1, Y + 1)) + CLng(iDATA(3, x, Y - 1)) + _
        CLng(iDATA(3, x, Y + 1)) + CLng(iDATA(3, x + 1, Y - 1)) + _
        CLng(iDATA(3, x + 1, Y)) + CLng(iDATA(3, x + 1, Y + 1))
      r = r \ 8
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub EdgeEnhance(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim mf As Long, dF As Long

  mProgress = 0
  mf = 9 + Factor
  dF = 1 + Factor
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(iDATA(1, x - 1, Y - 1)) + CLng(iDATA(1, x - 1, Y)) + _
        CLng(iDATA(1, x - 1, Y + 1)) + CLng(iDATA(1, x, Y - 1)) + _
        CLng(iDATA(1, x, Y + 1)) + CLng(iDATA(1, x + 1, Y - 1)) + _
        CLng(iDATA(1, x + 1, Y)) + CLng(iDATA(1, x + 1, Y + 1))
      B = (mf * CLng(iDATA(1, x, Y)) - B) \ dF
      G = CLng(iDATA(2, x - 1, Y - 1)) + CLng(iDATA(2, x - 1, Y)) + _
        CLng(iDATA(2, x - 1, Y + 1)) + CLng(iDATA(2, x, Y - 1)) + _
        CLng(iDATA(2, x, Y + 1)) + CLng(iDATA(2, x + 1, Y - 1)) + _
        CLng(iDATA(2, x + 1, Y)) + CLng(iDATA(2, x + 1, Y + 1))
      G = (mf * CLng(iDATA(2, x, Y)) - G) \ dF
      r = CLng(iDATA(3, x - 1, Y - 1)) + CLng(iDATA(3, x - 1, Y)) + _
        CLng(iDATA(3, x - 1, Y + 1)) + CLng(iDATA(3, x, Y - 1)) + _
        CLng(iDATA(3, x, Y + 1)) + CLng(iDATA(3, x + 1, Y - 1)) + _
        CLng(iDATA(3, x + 1, Y)) + CLng(iDATA(3, x + 1, Y + 1))
      r = (mf * CLng(iDATA(3, x, Y)) - r) \ dF
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Contour(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long
  
  mProgress = 0
  Call GetRGB(BackCol, cR, cG, cB)
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = CLng(bDATA(1, x - 1, Y - 1)) + CLng(bDATA(1, x - 1, Y)) + _
          CLng(bDATA(1, x - 1, Y + 1)) + CLng(bDATA(1, x, Y - 1)) + _
          CLng(bDATA(1, x, Y + 1)) + CLng(bDATA(1, x + 1, Y - 1)) + _
          CLng(bDATA(1, x + 1, Y)) + CLng(bDATA(1, x + 1, Y + 1))
      G = CLng(bDATA(2, x - 1, Y - 1)) + CLng(bDATA(2, x - 1, Y)) + _
          CLng(bDATA(2, x - 1, Y + 1)) + CLng(bDATA(2, x, Y - 1)) + _
          CLng(bDATA(2, x, Y + 1)) + CLng(bDATA(2, x + 1, Y - 1)) + _
          CLng(bDATA(2, x + 1, Y)) + CLng(bDATA(2, x + 1, Y + 1))
      r = CLng(bDATA(3, x - 1, Y - 1)) + CLng(bDATA(3, x - 1, Y)) + _
          CLng(bDATA(3, x - 1, Y + 1)) + CLng(bDATA(3, x, Y - 1)) + _
          CLng(bDATA(3, x, Y + 1)) + CLng(bDATA(3, x + 1, Y - 1)) + _
          CLng(bDATA(3, x + 1, Y)) + CLng(bDATA(3, x + 1, Y + 1))
      B = 8 * CLng(bDATA(1, x, Y)) - B + cB
      G = 8 * CLng(bDATA(2, x, Y)) - G + cG
      r = 8 * CLng(bDATA(3, x, Y)) - r + cR
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub GreyScale(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  
  mProgress = 0
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = iDATA(1, x, Y)
      G = iDATA(2, x, Y)
      r = iDATA(3, x, Y)
      B = Speed(r + G + B)
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = B
      iDATA(3, x, Y) = B
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Relief(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      B = 2 * CLng(bDATA(1, x - 1, Y - 1)) + CLng(bDATA(1, x - 1, Y)) + _
          CLng(bDATA(1, x, Y - 1)) - CLng(bDATA(1, x, Y + 1)) - _
          CLng(bDATA(1, x + 1, Y)) - 2 * CLng(bDATA(1, x + 1, Y + 1))
      G = 2 * CLng(bDATA(2, x - 1, Y - 1)) + CLng(bDATA(2, x - 1, Y)) + _
          CLng(bDATA(2, x, Y - 1)) - CLng(bDATA(2, x, Y + 1)) - _
          CLng(bDATA(2, x + 1, Y)) - 2 * CLng(bDATA(2, x + 1, Y + 1))
      r = 2 * CLng(bDATA(3, x - 1, Y - 1)) + CLng(bDATA(3, x - 1, Y)) + _
          CLng(bDATA(3, x, Y - 1)) - CLng(bDATA(3, x, Y + 1)) - _
          CLng(bDATA(3, x + 1, Y)) - 2 * CLng(bDATA(3, x + 1, Y + 1))
      B = (CLng(bDATA(1, x, Y)) + B) \ 2 + 50
      G = (CLng(bDATA(2, x, Y)) + G) \ 2 + 50
      r = (CLng(bDATA(3, x, Y)) + r) \ 2 + 50
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Brightness(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim sF As Single
  
  mProgress = 0
  sF = (Factor + 100) / 100
  For x = 0 To 255
    Speed(x) = x * sF
    If Speed(x) > 255 Then Speed(x) = 255
    If Speed(x) < 0 Then Speed(x) = 0
  Next x
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      iDATA(1, x, Y) = Speed(bDATA(1, x, Y))
      iDATA(2, x, Y) = Speed(bDATA(2, x, Y))
      iDATA(3, x, Y) = Speed(bDATA(3, x, Y))
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Contrast(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim sF As Single
  Dim mCol As Long, nCol As Long

  mProgress = 0
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  mCol = 0
  nCol = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      mCol = mCol + Speed(r + G + B)
      nCol = nCol + 1
    Next x
  Next Y
  mCol = mCol \ nCol
  sF = (Factor + 100) / 100
  For x = 0 To 255
    Speed(x) = (x - mCol) * sF + mCol
  Next x
  pProgress = 5
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = Speed(CLng(bDATA(1, x, Y)))
      G = Speed(CLng(bDATA(2, x, Y)))
      r = Speed(CLng(bDATA(3, x, Y)))
      Do While (B < 0) Or (B > 255) Or (G < 0) Or (G > 255) Or (r < 0) Or (r > 255)
        If (B <= 0) And (G <= 0) And (r <= 0) Then
          B = 0
          G = 0
          r = 0
        End If
        If (B >= 255) And (G >= 255) And (r >= 255) Then
          B = 255
          G = 255
          r = 255
        End If
        If B < 0 Then
          G = G + B \ 2
          r = r + B \ 2
          B = 0
        End If
        If B > 255 Then
          G = G + (B - 255) \ 2
          r = r + (B - 255) \ 2
          B = 255
        End If
        If G < 0 Then
          B = B + G \ 2
          r = r + G \ 2
          G = 0
        End If
        If G > 255 Then
          B = B + (G - 255) \ 2
          r = r + (G - 255) \ 2
          G = 255
        End If
        If r < 0 Then
          G = G + r \ 2
          B = B + r \ 2
          r = 0
        End If
        If r > 255 Then
          G = G + (r - 255) \ 2
          B = B + (r - 255) \ 2
          r = 255
        End If
      Loop
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = 5 + (Y * 95) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Pixelize(ByRef pProgress As Long, ByVal PixSize As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim pX As Long, pY As Long
  Dim sX As Long, sY As Long
  Dim mC As Long
  
  mProgress = 0
  B = 0: G = 0: r = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      If ((x - 1) Mod PixSize) = 0 Then
        sX = ((x - 1) \ PixSize) * PixSize + 1
        sY = ((Y - 1) \ PixSize) * PixSize + 1
        B = 0: G = 0: r = 0: mC = 0
        For pX = sX To sX + PixSize - 1
          For pY = sY To sY + PixSize - 1
            If (pX <= PicInfo.bmWidth) And (pY <= PicInfo.bmHeight) Then
              B = B + CLng(bDATA(1, pX, pY))
              G = G + CLng(bDATA(2, pX, pY))
              r = r + CLng(bDATA(3, pX, pY))
              mC = mC + 1
            End If
          Next pY
        Next pX
        If mC > 0 Then
          B = B \ mC
          G = G \ mC
          r = r \ mC
        End If
      End If
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub SwapBank(ByRef pProgress As Long, ByVal Modo As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long

  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      Select Case Modo
        Case 1: 'RGB -> BRG
          iDATA(1, x, Y) = G
          iDATA(2, x, Y) = r
          iDATA(3, x, Y) = B
        Case 2: 'RGB -> GBR
          iDATA(1, x, Y) = r
          iDATA(2, x, Y) = B
          iDATA(3, x, Y) = G
        Case 3: 'RGB -> RBG
          iDATA(1, x, Y) = G
          iDATA(2, x, Y) = B
          iDATA(3, x, Y) = r
        Case 4: 'RGB -> BGR
          iDATA(1, x, Y) = r
          iDATA(2, x, Y) = G
          iDATA(3, x, Y) = B
        Case 5: 'RGB -> GRB
          iDATA(1, x, Y) = B
          iDATA(2, x, Y) = r
          iDATA(3, x, Y) = G
      End Select
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub NearestColorBW(ByRef pProgress As Long, ByVal BackCol As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim cB As Long, cG As Long, cR As Long

  Call GetRGB(BackCol, cR, cG, cB)
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      If (r < cR) And (G < cG) And (B < cB) Then
        iDATA(1, x, Y) = 0
        iDATA(2, x, Y) = 0
        iDATA(3, x, Y) = 0
      Else
        iDATA(1, x, Y) = 255
        iDATA(2, x, Y) = 255
        iDATA(3, x, Y) = 255
      End If
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub EnhancedDiffusionBW(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim erro As Long, nCol As Long
  Dim mCol As Long

  mProgress = 0
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  mCol = 0
  nCol = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      mCol = mCol + B
      nCol = nCol + 1
    Next x
  Next Y
  pProgress = 5
  DoEvents
  mCol = mCol \ nCol
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      If (x > 1) And (Y > 1) And (x < PicInfo.bmWidth) And (Y < PicInfo.bmHeight) Then
        B = CLng(bDATA(1, x - 1, Y - 1)) + CLng(bDATA(1, x - 1, Y)) + _
          CLng(bDATA(1, x - 1, Y + 1)) + CLng(bDATA(1, x, Y - 1)) + _
          CLng(bDATA(1, x, Y + 1)) + CLng(bDATA(1, x + 1, Y - 1)) + _
          CLng(bDATA(1, x + 1, Y)) + CLng(bDATA(1, x + 1, Y + 1))
        G = CLng(bDATA(2, x - 1, Y - 1)) + CLng(bDATA(2, x - 1, Y)) + _
          CLng(bDATA(2, x - 1, Y + 1)) + CLng(bDATA(2, x, Y - 1)) + _
          CLng(bDATA(2, x, Y + 1)) + CLng(bDATA(2, x + 1, Y - 1)) + _
          CLng(bDATA(2, x + 1, Y)) + CLng(bDATA(2, x + 1, Y + 1))
        r = CLng(bDATA(3, x - 1, Y - 1)) + CLng(bDATA(3, x - 1, Y)) + _
          CLng(bDATA(3, x - 1, Y + 1)) + CLng(bDATA(3, x, Y - 1)) + _
          CLng(bDATA(3, x, Y + 1)) + CLng(bDATA(3, x + 1, Y - 1)) + _
          CLng(bDATA(3, x + 1, Y)) + CLng(bDATA(3, x + 1, Y + 1))
        B = (10 * CLng(bDATA(1, x, Y)) - B) \ 2
        G = (10 * CLng(bDATA(2, x, Y)) - G) \ 2
        r = (10 * CLng(bDATA(3, x, Y)) - r) \ 2
        If r > 255 Then r = 255
        If r < 0 Then r = 0
        If G > 255 Then G = 255
        If G < 0 Then G = 0
        If B > 255 Then B = 255
        If B < 0 Then B = 0
      Else
        B = CLng(bDATA(1, x, Y))
        G = CLng(bDATA(2, x, Y))
        r = CLng(bDATA(3, x, Y))
      End If
      B = Speed(r + G + B)
      B = B + erro
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If B < mCol Then nCol = 0 Else nCol = 255
      erro = (B - nCol) \ 4
      iDATA(1, x, Y) = nCol
      iDATA(2, x, Y) = nCol
      iDATA(3, x, Y) = nCol
    Next x
    mProgress = 5 + (Y * 95) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub OrderedDitherBW(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim VecDither(1 To 4, 1 To 4) As Byte
  Dim cx As Long, cy As Long

  VecDither(1, 1) = 1:    VecDither(1, 2) = 9
  VecDither(1, 3) = 3:    VecDither(1, 4) = 11
  VecDither(2, 1) = 13:   VecDither(2, 2) = 5
  VecDither(2, 3) = 15:   VecDither(2, 4) = 7
  VecDither(3, 1) = 4:    VecDither(3, 2) = 12
  VecDither(3, 3) = 2:    VecDither(3, 4) = 10
  VecDither(4, 1) = 16:   VecDither(4, 2) = 8
  VecDither(4, 3) = 14:   VecDither(4, 4) = 6
  mProgress = 0
  For x = 0 To 765
    Speed(x) = 1 + (x \ 3) \ 16
  Next x
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      cx = 1 + ((x - 1) Mod 4)
      cy = 1 + ((Y - 1) Mod 4)
      If B < VecDither(cx, cy) Then
        iDATA(1, x, Y) = 0
        iDATA(2, x, Y) = 0
        iDATA(3, x, Y) = 0
      Else
        iDATA(1, x, Y) = 255
        iDATA(2, x, Y) = 255
        iDATA(3, x, Y) = 255
      End If
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub FloydSteinbergBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 4, -255 To 255) As Long
  
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  For x = -255 To 255
    PartErr(1, x) = (7 * x) \ 16
    PartErr(2, x) = (3 * x) \ 16
    PartErr(3, x) = (5 * x) \ 16
    PartErr(4, x) = (1 * x) \ 16
  Next x
  erro = 0
  ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
  For x = 1 To PicInfo.bmWidth
    VecErro(1, x) = 0
    VecErro(2, x) = 0
  Next x
  pProgress = 2
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      mCol = mCol + B
      nCol = nCol + 1
    Next x
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      B = B + (VecErro(1, x) * 10) \ PalWeight
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If B < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, x, Y) = nCol
      iDATA(2, x, Y) = nCol
      iDATA(3, x, Y) = nCol
      erro = B - nCol
      If x < PicInfo.bmWidth Then VecErro(1, x + 1) = VecErro(1, x + 1) + PartErr(1, erro)
      If Y < PicInfo.bmHeight Then
        If x > 1 Then VecErro(2, x - 1) = VecErro(2, x - 1) + PartErr(2, erro)
        VecErro(2, x) = VecErro(2, x) + PartErr(3, erro)
        If x < PicInfo.bmWidth Then VecErro(2, x + 1) = VecErro(2, x + 1) + PartErr(4, erro)
      End If
    Next x
    For x = 1 To PicInfo.bmWidth
      VecErro(1, x) = VecErro(2, x)
      VecErro(2, x) = 0
    Next x
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub BurkeBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 7, -255 To 255) As Long
  
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  For x = -255 To 255
    PartErr(1, x) = (8 * x) \ 32
    PartErr(2, x) = (4 * x) \ 32
    PartErr(3, x) = (2 * x) \ 32
    PartErr(4, x) = (4 * x) \ 32
    PartErr(5, x) = (8 * x) \ 32
    PartErr(6, x) = (4 * x) \ 32
    PartErr(7, x) = (2 * x) \ 32
  Next x
  erro = 0
  ReDim VecErro(1 To 2, 1 To PicInfo.bmWidth) As Long
  For x = 1 To PicInfo.bmWidth
    VecErro(1, x) = 0
    VecErro(2, x) = 0
  Next x
  pProgress = 3
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      mCol = mCol + B
      nCol = nCol + 1
    Next x
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      B = B + (VecErro(1, x) * 10) \ PalWeight
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If B < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, x, Y) = nCol
      iDATA(2, x, Y) = nCol
      iDATA(3, x, Y) = nCol
      erro = (B - nCol)
      If x < PicInfo.bmWidth Then VecErro(1, x + 1) = VecErro(1, x + 1) + PartErr(1, erro)
      If x < PicInfo.bmWidth - 1 Then VecErro(1, x + 2) = VecErro(1, x + 2) + PartErr(2, erro)
      If Y < PicInfo.bmHeight Then
        If x > 2 Then VecErro(2, x - 2) = VecErro(2, x - 2) + PartErr(3, erro)
        If x > 1 Then VecErro(2, x - 1) = VecErro(2, x - 1) + PartErr(4, erro)
        VecErro(2, x) = VecErro(2, x) + PartErr(5, erro)
        If x < PicInfo.bmWidth Then VecErro(2, x + 1) = VecErro(2, x + 1) + PartErr(6, erro)
        If x < PicInfo.bmWidth - 1 Then VecErro(2, x + 2) = VecErro(2, x + 2) + PartErr(7, erro)
      End If
    Next x
    For x = 1 To PicInfo.bmWidth
      VecErro(1, x) = VecErro(2, x)
      VecErro(2, x) = 0
    Next x
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub StuckiBW(ByRef pProgress As Long, ByVal PalWeight As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim erro As Long
  Dim VecErro() As Long
  Dim nCol As Long, mCol As Long
  Dim PartErr(1 To 12, -255 To 255) As Long
  
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  For x = -255 To 255
    PartErr(1, x) = (8 * x) \ 42
    PartErr(2, x) = (4 * x) \ 42
    PartErr(3, x) = (2 * x) \ 42
    PartErr(4, x) = (4 * x) \ 42
    PartErr(5, x) = (8 * x) \ 42
    PartErr(6, x) = (4 * x) \ 42
    PartErr(7, x) = (2 * x) \ 42
    PartErr(8, x) = (1 * x) \ 42
    PartErr(9, x) = (2 * x) \ 42
    PartErr(10, x) = (4 * x) \ 42
    PartErr(11, x) = (2 * x) \ 42
    PartErr(12, x) = (1 * x) \ 42
  Next x
  erro = 0
  ReDim VecErro(1 To 3, 1 To PicInfo.bmWidth) As Long
  For x = 1 To PicInfo.bmWidth
    VecErro(1, x) = 0
    VecErro(2, x) = 0
    VecErro(3, x) = 0
  Next x
  pProgress = 3
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      mCol = mCol + B
      nCol = nCol + 1
    Next x
  Next Y
  mCol = mCol \ nCol
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = Speed(r + G + B)
      B = B + (VecErro(1, x) * 10) \ PalWeight
      If B < 0 Then B = 0
      If B > 255 Then B = 255
      If B < mCol Then nCol = 0 Else nCol = 255
      iDATA(1, x, Y) = nCol
      iDATA(2, x, Y) = nCol
      iDATA(3, x, Y) = nCol
      erro = (B - nCol)
      If x < PicInfo.bmWidth Then VecErro(1, x + 1) = VecErro(1, x + 1) + PartErr(1, erro)
      If x < PicInfo.bmWidth - 1 Then VecErro(1, x + 2) = VecErro(1, x + 2) + PartErr(2, erro)
      If Y < PicInfo.bmHeight Then
        If x > 2 Then VecErro(2, x - 2) = VecErro(2, x - 2) + PartErr(3, erro)
        If x > 1 Then VecErro(2, x - 1) = VecErro(2, x - 1) + PartErr(4, erro)
        VecErro(2, x) = VecErro(2, x) + PartErr(5, erro)
        If x < PicInfo.bmWidth Then VecErro(2, x + 1) = VecErro(2, x + 1) + PartErr(6, erro)
        If x < PicInfo.bmWidth - 1 Then VecErro(2, x + 2) = VecErro(2, x + 2) + PartErr(7, erro)
      End If
      If Y < PicInfo.bmHeight - 1 Then
        If x > 2 Then VecErro(3, x - 2) = VecErro(3, x - 2) + PartErr(8, erro)
        If x > 1 Then VecErro(3, x - 1) = VecErro(3, x - 1) + PartErr(9, erro)
        VecErro(3, x) = VecErro(3, x) + PartErr(10, erro)
        If x < PicInfo.bmWidth Then VecErro(3, x + 1) = VecErro(3, x + 1) + PartErr(11, erro)
        If x < PicInfo.bmWidth - 1 Then VecErro(3, x + 2) = VecErro(3, x + 2) + PartErr(12, erro)
      End If
    Next x
    For x = 1 To PicInfo.bmWidth
      VecErro(1, x) = VecErro(2, x)
      VecErro(2, x) = VecErro(3, x)
      VecErro(3, x) = 0
    Next x
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Aqua(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim Med(1 To 4) As Long
  Dim Dev(1 To 4) As Long
  Dim i As Long, j As Long
  Dim sDev As Long, vDev As Long
  
  mProgress = 0
  For Y = 3 To PicInfo.bmHeight - 2
    For x = 3 To PicInfo.bmWidth - 2
      For i = 1 To 3
        Med(1) = CLng(bDATA(i, x - 2, Y - 2)) + CLng(bDATA(i, x - 1, Y - 2)) + CLng(bDATA(i, x, Y - 2)) + _
                 CLng(bDATA(i, x - 2, Y - 1)) + CLng(bDATA(i, x - 1, Y - 1)) + CLng(bDATA(i, x, Y - 1)) + _
                 CLng(bDATA(i, x - 2, Y)) + CLng(bDATA(i, x - 1, Y)) + CLng(bDATA(i, x, Y))
        Med(2) = CLng(bDATA(i, x + 2, Y - 2)) + CLng(bDATA(i, x + 1, Y - 2)) + CLng(bDATA(i, x, Y - 2)) + _
                 CLng(bDATA(i, x + 2, Y - 1)) + CLng(bDATA(i, x + 1, Y - 1)) + CLng(bDATA(i, x, Y - 1)) + _
                 CLng(bDATA(i, x + 2, Y)) + CLng(bDATA(i, x + 1, Y)) + CLng(bDATA(i, x, Y))
        Med(3) = CLng(bDATA(i, x - 2, Y + 2)) + CLng(bDATA(i, x - 1, Y + 2)) + CLng(bDATA(i, x, Y + 2)) + _
                 CLng(bDATA(i, x - 2, Y + 1)) + CLng(bDATA(i, x - 1, Y + 1)) + CLng(bDATA(i, x, Y + 1)) + _
                 CLng(bDATA(i, x - 2, Y)) + CLng(bDATA(i, x - 1, Y)) + CLng(bDATA(i, x, Y))
        Med(4) = CLng(bDATA(i, x + 2, Y + 2)) + CLng(bDATA(i, x + 1, Y + 2)) + CLng(bDATA(i, x, Y + 2)) + _
                 CLng(bDATA(i, x + 2, Y + 1)) + CLng(bDATA(i, x + 1, Y + 1)) + CLng(bDATA(i, x, Y + 1)) + _
                 CLng(bDATA(i, x + 2, Y)) + CLng(bDATA(i, x + 1, Y)) + CLng(bDATA(i, x, Y))
        Med(1) = Med(1) \ 9
        Med(2) = Med(2) \ 9
        Med(3) = Med(3) \ 9
        Med(4) = Med(4) \ 9
        Dev(1) = Abs(CLng(bDATA(i, x - 2, Y - 2)) - Med(1)) + Abs(CLng(bDATA(i, x - 1, Y - 2)) - Med(1)) + Abs(CLng(bDATA(i, x, Y - 2)) - Med(1)) + _
                 Abs(CLng(bDATA(i, x - 2, Y - 1)) - Med(1)) + Abs(CLng(bDATA(i, x - 1, Y - 1)) - Med(1)) + Abs(CLng(bDATA(i, x, Y - 1)) - Med(1)) + _
                 Abs(CLng(bDATA(i, x - 2, Y)) - Med(1)) + Abs(CLng(bDATA(i, x - 1, Y)) - Med(1)) + Abs(CLng(bDATA(i, x, Y)) - Med(1))
        Dev(2) = Abs(CLng(bDATA(i, x + 2, Y - 2)) - Med(2)) + Abs(CLng(bDATA(i, x + 1, Y - 2)) - Med(2)) + Abs(CLng(bDATA(i, x, Y - 2)) - Med(2)) + _
                 Abs(CLng(bDATA(i, x + 2, Y - 1)) - Med(2)) + Abs(CLng(bDATA(i, x + 1, Y - 1)) - Med(2)) + Abs(CLng(bDATA(i, x, Y - 1)) - Med(2)) + _
                 Abs(CLng(bDATA(i, x + 2, Y)) - Med(2)) + Abs(CLng(bDATA(i, x + 1, Y)) - Med(2)) + Abs(CLng(bDATA(i, x, Y)) - Med(2))
        Dev(3) = Abs(CLng(bDATA(i, x - 2, Y + 2)) - Med(3)) + Abs(CLng(bDATA(i, x - 1, Y + 2)) - Med(3)) + Abs(CLng(bDATA(i, x, Y + 2)) - Med(3)) + _
                 Abs(CLng(bDATA(i, x - 2, Y + 1)) - Med(3)) + Abs(CLng(bDATA(i, x - 1, Y + 1)) - Med(3)) + Abs(CLng(bDATA(i, x, Y + 1)) - Med(3)) + _
                 Abs(CLng(bDATA(i, x - 2, Y)) - Med(3)) + Abs(CLng(bDATA(i, x - 1, Y)) - Med(3)) + Abs(CLng(bDATA(i, x, Y)) - Med(3))
        Dev(4) = Abs(CLng(bDATA(i, x + 2, Y + 2)) - Med(4)) + Abs(CLng(bDATA(i, x + 1, Y + 2)) - Med(4)) + Abs(CLng(bDATA(i, x, Y + 2)) - Med(4)) + _
                 Abs(CLng(bDATA(i, x + 2, Y + 1)) - Med(4)) + Abs(CLng(bDATA(i, x + 1, Y + 1)) - Med(4)) + Abs(CLng(bDATA(i, x, Y + 1)) - Med(4)) + _
                 Abs(CLng(bDATA(i, x + 2, Y)) - Med(4)) + Abs(CLng(bDATA(i, x + 1, Y)) - Med(4)) + Abs(CLng(bDATA(i, x, Y)) - Med(4))
        vDev = 99999
        sDev = 0
        For j = 1 To 4
          If Dev(j) < vDev Then
            vDev = Dev(j)
            sDev = j
          End If
        Next j
        iDATA(i, x, Y) = Med(sDev)
      Next i
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Dilate(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim V As Long
  Dim i As Long
  Dim vMax As Long
  
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      For i = 1 To 3
        vMax = 0
        V = CLng(bDATA(i, x - 1, Y - 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x, Y - 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x + 1, Y - 1))
        If V > vMax Then vMax = V
        
        V = CLng(bDATA(i, x - 1, Y))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x, Y))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x + 1, Y))
        If V > vMax Then vMax = V
        
        V = CLng(bDATA(i, x - 1, Y + 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x, Y + 1))
        If V > vMax Then vMax = V
        V = CLng(bDATA(i, x + 1, Y + 1))
        If V > vMax Then vMax = V
        
        iDATA(i, x, Y) = vMax
      Next i
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Erode(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim V As Long
  Dim i As Long
  Dim vMin As Long
  
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      For i = 1 To 3
        vMin = 255
        V = CLng(bDATA(i, x - 1, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y - 1))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(i, x - 1, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(i, x - 1, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y + 1))
        If V < vMin Then vMin = V
        
        iDATA(i, x, Y) = vMin
      Next i
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub ConnectedContour(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim V As Long
  Dim i As Long
  Dim vMin As Long
  
  mProgress = 0
  For Y = 2 To PicInfo.bmHeight - 1
    For x = 2 To PicInfo.bmWidth - 1
      For i = 1 To 3
        vMin = 255
        V = CLng(bDATA(i, x - 1, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y - 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y - 1))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(i, x - 1, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y))
        If V < vMin Then vMin = V
        
        V = CLng(bDATA(i, x - 1, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x, Y + 1))
        If V < vMin Then vMin = V
        V = CLng(bDATA(i, x + 1, Y + 1))
        If V < vMin Then vMin = V
        
        iDATA(i, x, Y) = CLng(iDATA(i, x, Y)) - vMin
      Next i
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub StretchHistogram(ByRef pProgress As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim bMin As Long, bMax As Long
  Dim gMin As Long, gMax As Long
  Dim rMin As Long, rMax As Long
  
  mProgress = 0
  bMin = 255: bMax = 0
  gMin = 255: gMax = 0
  rMin = 255: rMax = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      If B < bMin Then bMin = B
      If B > bMax Then bMax = B
      If G < gMin Then gMin = G
      If G > gMax Then gMax = G
      If r < rMin Then rMin = r
      If r > rMax Then rMax = r
    Next x
    mProgress = (Y * 10) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 10
  DoEvents
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      B = 255 * (B - bMin) / (bMax - bMin)
      G = 255 * (G - gMin) / (gMax - gMin)
      r = 255 * (r - rMin) / (rMax - rMin)
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = 10 + (Y * 90) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub AddNoise(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim V As Long
    
  mProgress = 0
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      G = CLng(bDATA(2, x, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      r = CLng(bDATA(3, x, Y)) + ((Factor * 2 + 1) * Rnd - Factor)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub Saturation(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim V As Long
  Dim sF As Single
    
  mProgress = 0
  For x = 0 To 765
    Speed(x) = x \ 3
  Next x
  sF = Factor / 100
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      V = Speed(B + G + r)
      B = B + sF * (B - V)
      G = G + sF * (G - V)
      r = r + sF * (r - V)
      If r > 255 Then r = 255
      If r < 0 Then r = 0
      If G > 255 Then G = 255
      If G < 0 Then G = 0
      If B > 255 Then B = 255
      If B < 0 Then B = 0
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub

Private Sub GammaCorrection(ByRef pProgress As Long, ByVal Factor As Long)
  Dim x As Long, Y As Long
  Dim r As Long, G As Long, B As Long
  Dim dB As Double, dG As Double, dR As Double
  Dim sF As Single
  Dim Max As Double, Min As Double, MM As Double
  Dim h As Double, s As Double, i As Double
  Dim cB As Double, cG As Double, cR As Double
  Dim Flo As Long
    
  mProgress = 0
  sF = Factor / 100
  For Y = 1 To PicInfo.bmHeight
    For x = 1 To PicInfo.bmWidth
      'get data
      B = CLng(bDATA(1, x, Y))
      G = CLng(bDATA(2, x, Y))
      r = CLng(bDATA(3, x, Y))
      dB = B / 255
      dG = G / 255
      dR = r / 255
      'correct gamma
      dB = dB ^ (1 / sF)
      dG = dG ^ (1 / sF)
      dR = dR ^ (1 / sF)
      'set data
      B = dB * 255
      G = dG * 255
      r = dR * 255
      iDATA(1, x, Y) = B
      iDATA(2, x, Y) = G
      iDATA(3, x, Y) = r
    Next x
    mProgress = (Y * 100) \ PicInfo.bmHeight
    pProgress = mProgress
    DoEvents
  Next Y
  pProgress = 100
  DoEvents
End Sub


