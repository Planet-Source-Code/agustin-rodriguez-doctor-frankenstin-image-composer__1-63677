VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5085
   FillColor       =   &H000000FF&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main Menu.frx":0000
   ScaleHeight     =   323
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Mascara 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4890
      Left            =   5415
      Picture         =   "Main Menu.frx":51842
      ScaleHeight     =   326
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   0
      Top             =   30
      Visible         =   0   'False
      Width           =   5115
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   2430
      Top             =   2235
      Width           =   405
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private coord(20) As POINTAPI
Private Cor(20) As Long

Private capture As Integer
Private xx As Single
Private yy As Single

Private Radiano As Single

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
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
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
    SetLayeredWindowAttributes Me.hwnd, 0, 255, LWA_COLORKEY Or LWA_ALPHA
    Form2.Show , Me
    
    Radiano = 3.14156 / 180
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim rx As Single
  Dim ry As Single

    GetCursorPos Pt
    rx = Pt.x
    ry = Pt.Y
    If Button = 1 Then
        xx = x * Screen.TwipsPerPixelX: yy = Y * Screen.TwipsPerPixelY
        capture = True
        'Call Form2.Form_MouseDown(Button, Shift, rx, ry)
        Call Form2.Capture_move
        ReleaseCapture
        SetCapture Me.hwnd
        
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Static Ultima_cor As Long
  Dim i As Integer

    If capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - xx, Pt.Y * Screen.TwipsPerPixelY - yy
        Call Form2.Form_MouseMove(Button, Shift, x, Y)
    End If

    If Mascara.Point(x, Y) <> Ultima_cor Then
        'Debug.Print Mascara.Point(x, Y);
      
        For i = 0 To 19
            If Mascara.Point(x, Y) = Cor(i) Then
                Escolha = i
                Form2.Timer1.Enabled = True
                'Debug.Print i;
                Ultima_cor = Cor(i)
                If i < 10 Then
                    Form2.Picture = Form2.Picture1(0).Picture
                    Form2.Width = Form2.Picture1(0).Width * Screen.TwipsPerPixelX
                    Form2.Height = Form2.Picture1(0).Height * Screen.TwipsPerPixelY
                  Else 'NOT I...
                    Form2.Picture = Form2.Picture1(1).Picture
                    Form2.Width = Form2.Picture1(1).Width * Screen.TwipsPerPixelX
                    Form2.Height = Form2.Picture1(1).Height * Screen.TwipsPerPixelY
                End If
                
                Form2.Move left + coord(i).x * Screen.TwipsPerPixelX - Form2.Width / 2, top + coord(i).Y * Screen.TwipsPerPixelY - Form2.Height / 2
            End If
        Next i
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False
    Call Form2.Form_MouseUp(Button, Shift, 0, 0)

End Sub

Private Sub Form_Resize()

  Dim Ang As Integer, ag As Single, xt As Single, yt As Single, raio As Integer, i As Integer

    'raio = 130
    raio = 65
    For Ang = 18 To 360 Step 36
    
        ag = Radiano * Ang
        xt = 172 + (Sin(ag)) * 130 'raio
        yt = 162 + -Cos(ag) * 130 'raio
        ''''''''''''''''''''''''''''''''''''''Circle (xt, yt), 5
        coord(i).x = xt
        coord(i).Y = yt
        Cor(i) = Mascara.Point(xt, yt)

        xt = 172 + (Sin(ag)) * 65 'raio
        yt = 162 + -Cos(ag) * 65 '' raio
        ''''''''''''''''''''''''''''''''''''''Circle (xt, yt), 5
        coord(i + 10).x = xt
        coord(i + 10).Y = yt
        Cor(i + 10) = Mascara.Point(xt, yt)

        i = i + 1

    Next Ang

    coord(20).x = Image1.left + 10
    coord(20).Y = Image1.top + 10

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Escolha = 20
    Form2.Move left + coord(20).x * Screen.TwipsPerPixelX - Form2.Width / 2, top + coord(20).Y * Screen.TwipsPerPixelY - Form2.Height / 2

End Sub


