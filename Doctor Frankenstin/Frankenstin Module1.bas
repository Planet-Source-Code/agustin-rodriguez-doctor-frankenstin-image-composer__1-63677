Attribute VB_Name = "Frankenstin"
Option Explicit
Public Objeto_foi_usado(1000) As Integer

Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Declare Function GetSystemMetrics Lib "User32" _
(ByVal nIndex As Long) As Long

Public Const SM_CXSCREEN = 0  ' Screen Width
Public Const SM_CYSCREEN = 1  ' Screen Height
Public Const SM_CYCAPTION = 4 ' Height of window caption
Public Const SM_CYMENU = 15   ' Height of menu
Public Const SM_CXDLGFRAME = 7   ' Width of borders X & Y same + 1 for sizable
Public Const SM_CYSMCAPTION = 51 ' Height of small caption (Tool Windows)


Public Escolha As Integer

Public tel(1000) As Integer
Public EscalaX(1000) As Single, EscalaY(1000) As Single

Public EscalaZX(1000) As Single
Public EscalaZY(1000) As Single

Public Arquivo As String
Public free As Integer

Public Ordem As String

Public Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Type POINTAPI
    x As Long
    Y As Long
End Type
Public Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Pt As POINTAPI
Public xx As Long
Public yy As Long
Public capture As Integer


Public Arquivo_usado_pelo_objeto(1000) As String
Public Quant_Objetos As Integer
Public Pastas(1000) As String

Public Ponteiro_de_arquivo(1000) As Integer
Public use_PASTAS(1000)


Public Qt_pastas As Integer
Public My_path As String

Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1
Public Const SWP_NOZORDER As Long = &H4
Public Const SWP_FRAMECHANGED As Long = &H20
Public Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW As Long = &H80
Public Const SWP_NOACTIVATE As Long = &H10
Public Const SWP_NOCOPYBITS As Long = &H100
Public Const SWP_NOOWNERZORDER As Long = &H200
Public Const SWP_NOREDRAW As Long = &H8
Public Const SWP_NOREMenu_options As Long = SWP_NOOWNERZORDER
Public Const SWP_SHOWWINDOW As Long = &H40

Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const GWL_HWNDPARENT As Long = (-8)

Public F() As New Objeto

Public Selected As Integer


Public Declare Function StretchBlt Lib "GDI32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetStretchBltMode Lib "GDI32" (ByVal hDC As Long, ByVal hStretchMode As Long) As Long

Public Const STRETCHMODE As Long = vbPaletteModeNone


Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Const GWL_EXSTYLE As Long = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const WS_EX_TRANSPARENT As Long = &H20&
Public Const LWA_ALPHA As Long = &H2&
Public Const LWA_COLORKEY As Integer = &H1

Public Const MK_LBUTTON As Long = &H1
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP As Long = &H202
    
Public Const MK_RBUTTON As Long = &H2
Public Const WM_RBUTTONDOWN As Long = &H204
Public Const WM_RBUTTONUP As Long = &H205

Public Const WM_LBUTTONDBLCLK As Long = &H203
    
Public Const HTCAPTION As Long = 2
Public Const WM_NCLBUTTONDOWN As Long = &HA1

Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


  
Public Declare Function PlgBlt Lib "GDI32" (ByVal hdcDest As Long, _
                        lpPoint As POINTS2D, _
                        ByVal hdcSrc As Long, _
                        ByVal nXSrc As Long, _
                        ByVal nYSrc As Long, _
                        ByVal nWidth As Long, _
                        ByVal nHeight As Long, _
                        ByVal hbmMask As Long, _
                        ByVal xMask As Long, _
                        ByVal yMask As Long) As Long

Public Const NotPI = 3.14159265238 / 180

'--------------------------------------------------------------------------------
Public Type POINTS2D
    x As Long
    Y As Long
End Type

Public Sub Rotate(ByRef picDestHdc As Long, xPos As Long, yPos As Long, _
                  ByVal Angle As Long, _
                  ByRef picSrcHdc As Long, srcXoffset As Long, srcYoffset As Long, _
                  ByVal srcWidth As Long, ByVal srcHeight As Long)

  '## Rotate - Rotates an image.
  '##
  '## PicDestHdc      = the hDc of the target picturebox (ie. Picture2.hdc )
  '## xPos            = the target coordinates (note that the image will be centered around these
  '## yPos              coordinates).
  '## Angle           = Rotate Angle (0-360)
  '## PicSrcHdc       = The source image to rotate (ie. Picture1.hdc )
  '## srcXoffset      = The offset coordinates within the Source Image to grab.
  '## srcYoffset
  '## srcWidth        = The width/height of the source image to grab.
  '## srcHeight
  '##
  '## Returns: Nothing.

  Dim Points(3) As POINTS2D
  Dim DefPoints(3) As POINTS2D
  Dim ThetS As Single, ThetC As Single
  Dim Ret As Long
    
    'SET LOCAL AXIS / ALIGNMENT
    Points(0).x = -srcWidth * 0.5
    Points(0).Y = -srcHeight * 0.5
    
    Points(1).x = Points(0).x + srcWidth
    Points(1).Y = Points(0).Y
    
    Points(2).x = Points(0).x
    Points(2).Y = Points(0).Y + srcHeight
    
    'ROTATE AROUND Z-AXIS
    ThetS = Sin(Angle * NotPI)
    ThetC = Cos(Angle * NotPI)
    
    DefPoints(0).x = (Points(0).x * ThetC - Points(0).Y * ThetS) + xPos
    DefPoints(0).Y = (Points(0).x * ThetS + Points(0).Y * ThetC) + yPos
    
    DefPoints(1).x = (Points(1).x * ThetC - Points(1).Y * ThetS) + xPos
    DefPoints(1).Y = (Points(1).x * ThetS + Points(1).Y * ThetC) + yPos
    
    DefPoints(2).x = (Points(2).x * ThetC - Points(2).Y * ThetS) + xPos
    DefPoints(2).Y = (Points(2).x * ThetS + Points(2).Y * ThetC) + yPos
    
    PlgBlt picDestHdc, DefPoints(0), picSrcHdc, srcXoffset, srcYoffset, srcWidth, srcHeight, 0, 0, 0
    
End Sub


Public Function Ordene()
Dim i As Integer

    For i = Len(Ordem) To 2 Step -1
        F(Asc(Mid(Ordem, (i), 1)) - 1).Show , F(Asc(Mid(Ordem, (i - 1), 1)) - 1)
    Next
    F(Asc(Mid(Ordem, (1), 1)) - 1).Show
    
End Function

Public Function Libere_Ordem()
Dim i As Integer
    For i = Len(Ordem) To 1 Step -1
        F(Asc(Mid(Ordem, (i), 1)) - 1).Show , Splash_Form
    Next
End Function

Private Sub Main()
Splash_Form.Show 1

Load Main_form
Form1.Move Screen.Width - Form1.Width, 0
Form1.Show

End Sub

