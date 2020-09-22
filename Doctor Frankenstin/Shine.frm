VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   840
   ClientLeft      =   -1005
   ClientTop       =   0
   ClientWidth     =   960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Shine.frx":0000
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   64
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   3
      Left            =   1785
      Picture         =   "Shine.frx":2F82
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   3
      Top             =   3105
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Index           =   2
      Left            =   225
      Picture         =   "Shine.frx":5F04
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   2
      Top             =   3105
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   1
      Left            =   1530
      Picture         =   "Shine.frx":B446
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   1
      Top             =   1695
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Index           =   0
      Left            =   0
      Picture         =   "Shine.frx":E3C8
      ScaleHeight     =   63
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   0
      Top             =   -60
      Width           =   945
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1305
      Top             =   960
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private capture As Integer
Private xx As Single
Private yy As Single

Private Type POINTAPI
    x As Long
    Y As Long
End Type
Private Pt As POINTAPI

Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "User32" () As Long

Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const LWA_ALPHA As Long = &H2&
Private Const LWA_COLORKEY As Integer = &H1

Private Sub Form_Load()

  Dim Ret As Long
  Dim Normalwindowstyle As Long

    Normalwindowstyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hwnd, vbWhite, 155, LWA_COLORKEY Or LWA_ALPHA
    Width = Picture1(0).Width * Screen.TwipsPerPixelX
    Height = Picture1(0).Height * Screen.TwipsPerPixelY

End Sub

Public Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Timer2.Enabled = True

End Sub

Public Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Static vez As Integer

    If capture Then
        GetCursorPos Pt
        
        Move Pt.x * Screen.TwipsPerPixelX - xx, Pt.Y * Screen.TwipsPerPixelY - yy
    End If

End Sub

Public Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Timer2.Enabled = False
    capture = False

End Sub

Public Sub Capture_move()

    GetCursorPos Pt

    xx = Pt.x * Screen.TwipsPerPixelX - left
    yy = Pt.Y * Screen.TwipsPerPixelY - top
        
    capture = True
    ReleaseCapture
    SetCapture Me.hwnd

End Sub

Private Sub Picture1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Timer2.Enabled = True

End Sub

Private Sub Picture1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim texto(20) As String

    texto(0) = " Less Vertical Stretch to the object " & Pastas(use_PASTAS(Selected)) + "            [ Key Z ]"
    texto(1) = " More Horizontal Stretch to the object " & Pastas(use_PASTAS(Selected)) + "            [ Key S ]"
    texto(2) = " Right Rotation to object " & Pastas(use_PASTAS(Selected)) + "            [ RIGHT Cursor Key or CTRL + Mouse WELL UP]"
    texto(3) = " Next Image in the Library " & Pastas(use_PASTAS(Selected)) + "            [ Up Cursor Key or SHIFT + Mouse WELL UP]"
    texto(4) = " Previous Image in the Library " & Pastas(use_PASTAS(Selected)) + "            [ Down Cursor Key or SHIFT + Mouse WELL DOWN]"
    texto(5) = " Less Scale to the object " & Pastas(use_PASTAS(Selected)) + "            [ - Key or Mouse WELL DOWN]"
    texto(6) = " More Scale to the object " & Pastas(use_PASTAS(Selected)) + "            [ + Key or Mouse WELL UP]"
    texto(7) = " Left Rotation to object " & Pastas(use_PASTAS(Selected)) + "            [ LEFT Cursor Key or CTRL + Mouse WELL DOWN]"
    texto(8) = " Less Horizontal Stretch to the object " & Pastas(use_PASTAS(Selected)) + "            [ Key A ]"
    texto(9) = " More Vertical Stretch to the object " & Pastas(use_PASTAS(Selected)) + "            [ Key W ]"
    texto(10) = " New Object from Library"
    texto(11) = " New Object from graphic file"
    texto(12) = " New Project"
    texto(13) = " Delete active objet (" & Pastas(use_PASTAS(Selected)) & ")"
    texto(14) = " Nothing"
    texto(15) = " Show Transparent Mask"
    texto(16) = " Show Objects as Stretched Picture ( NOTE: The fit is relative to last selected object (" & Pastas(use_PASTAS(Selected)) & ")"
    texto(17) = " Open Project"
    texto(18) = " Save Project"
    texto(19) = " Clone the active object (" & Pastas(use_PASTAS(Selected)) & ")"""

    Picture1(0).ToolTipText = texto(Escolha)

End Sub

Private Sub Picture1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Timer2.Enabled = False

End Sub

Private Sub Timer1_Timer()

  Static x As Integer, passo As Integer

    If passo = 0 Then
        passo = 20: x = 100
    End If
    If (x + passo > 200) Or (x + passo < 40) Then
        passo = -passo
    End If
    x = x + passo

    SetLayeredWindowAttributes Me.hwnd, vbWhite, x, LWA_COLORKEY Or LWA_ALPHA

End Sub

Private Sub Timer2_Timer()

    Select Case Escolha
      Case 0
        F(Selected).Form_KeyDown Asc("Z"), 0
      Case 9
        F(Selected).Form_KeyDown Asc("W"), 0
      Case 8
        F(Selected).Form_KeyDown Asc("A"), 0
      Case 1
        F(Selected).Form_KeyDown Asc("S"), 0
      Case 2
        F(Selected).Form_KeyDown 39, 0
      Case 7
        F(Selected).Form_KeyDown 37, 0
      Case 5
        F(Selected).Form_KeyDown 109, 0
      Case 6
        F(Selected).Form_KeyDown 107, 0
      Case 3
        F(Selected).Form_KeyDown 38, 0
        Timer2.Enabled = False
      Case 4
        F(Selected).Form_KeyDown 40, 0
        Timer2.Enabled = False

      Case 19
        Timer2.Enabled = False
        Main_form.Menu_options_Click (0)

      Case 18
        Timer2.Enabled = False
        Main_form.Menu_options_Click (4)

      Case 17
        Timer2.Enabled = False
        Main_form.Menu_options_Click (3)

      Case 16
        Timer2.Enabled = False
        F(Selected).SetFocus
        Main_form.Menu_options_Click (6)

      Case 12
        Timer2.Enabled = False
        Main_form.Menu_options_Click (2)

      Case 11
        Timer2.Enabled = False
        Main_form.From_files_Click

      Case 10
        Timer2.Enabled = False
        PopupMenu Main_form.From_Library(0)
        Main_form.From_Library(0).Visible = True

      Case 13
        Timer2.Enabled = False
        Main_form.Menu_options_Click (5)

      Case 14

      Case 15
        Timer2.Enabled = False
        Main_form.Menu_options_Click (7)

      Case 20
        Main_form.Menu_options_Click (8)    'END
    End Select

End Sub

