VERSION 5.00
Begin VB.Form Mask 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Mask Maker"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7050
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   470
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   5325
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   3
      Top             =   3750
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      Left            =   5460
      Max             =   255
      Min             =   10
      TabIndex        =   2
      Top             =   5280
      Value           =   10
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   375
      Top             =   4305
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   255
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3510
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Trabalho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   3705
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   104
      TabIndex        =   0
      Top             =   3750
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   6675
      MouseIcon       =   "Mask.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   30
      Width           =   225
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000001FF&
      BackStyle       =   1  'Opaque
      FillColor       =   &H000000FF&
      Height          =   495
      Left            =   6525
      Shape           =   3  'Circle
      Top             =   15
      Width           =   525
   End
   Begin VB.Image Cor 
      Height          =   180
      Index           =   2
      Left            =   4275
      Picture         =   "Mask.frx":030A
      Top             =   3300
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Cor 
      Height          =   180
      Index           =   1
      Left            =   3765
      Picture         =   "Mask.frx":0408
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Cor 
      Height          =   180
      Index           =   0
      Left            =   3450
      Picture         =   "Mask.frx":0506
      Top             =   3345
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image Ponto 
      Height          =   120
      Index           =   0
      Left            =   1980
      Picture         =   "Mask.frx":0603
      Top             =   1245
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu x 
         Caption         =   "- Use  UP or DOWN Arrows to adjust the Transparence Level"
         Index           =   0
      End
      Begin VB.Menu x 
         Caption         =   "- Double Click or press SPACE key to  Mark the Points"
         Index           =   1
      End
      Begin VB.Menu x 
         Caption         =   "- Use LEFT or RIGHT Arrows to Navigate on the line and to see the points"
         Index           =   2
      End
      Begin VB.Menu x 
         Caption         =   "- Drag the Points to adjust the Bezier Curve"
         Index           =   3
      End
      Begin VB.Menu x 
         Caption         =   "- After to close the area, click with SHIFT key pressed inside of it to see the black selection "
         Index           =   4
      End
      Begin VB.Menu x 
         Caption         =   "- Clck inside of  the Black selection with CTRL key pressed to make the Final cut"
         Index           =   5
      End
      Begin VB.Menu x 
         Caption         =   "- Close the Window and the Area will  to be transfer"
         Index           =   6
      End
   End
End
Attribute VB_Name = "Mask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private P_Up As Long
Private P_Down As Long
Private P_Left As Long
Private P_Right As Long

Private Criou_mascara As Integer

Private Grossura_da_linha As Integer
Private PosX As Single
Private PosY As Single

Private FatorX As Single
Private FatorY As Single
Private BaseWidth As Single
Private BaseHeight As Single

Private Const Quant_UNDO As Integer = 127

Private Setor As Integer
Private Escolhidos As String * 16
Private Backup(256, 127) As Long
Private Backup_34ou44(127) As Integer

Private Nivel_para_Undo As Integer

Private Const MK_LBUTTON As Long = &H1
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
    
Private Const MK_RBUTTON As Long = &H2
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205

Private Const WM_LBUTTONDBLCLK As Long = &H203
    
Private Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Declare Function PolyBezier Lib "GDI32" (ByVal hDC As Long, lppt As POINTAPI, ByVal cPoints As Long) As Long

Private Type coor
    x As Long
    Y As Long
End Type
Private xx As Long
Private yy As Long

Private Nr_pontos As Integer
Private Ponto_atual As Integer
Private Sequencia() As POINTAPI

Private Sub Command2_Click()

  Dim Desktop As Long
  Dim ww As Long
  Dim hh As Long
  Dim xoff As Long
  Dim yoff As Long
  Dim i As Integer
  Dim x As Long

    x = Trabalho.ForeColor

    Trabalho.ForeColor = Trabalho.BackColor
    x = (Nr_pontos \ 3) * 3 + 1
    Trabalho.DrawWidth = Grossura_da_linha

    Call PolyBezier(Trabalho.hDC, Sequencia(0), x)

    Me.PaintPicture Trabalho.Image, 0, 0, Width, Height, 0, 0, Width, Height, vbSrcCopy
    Trabalho.ForeColor = x

    '-------------------------------------------------
    
    For i = 0 To Nr_pontos - 1
        Ponto(i).Visible = False
    Next i
 
    Hide
    Desktop = GetDC(GetDesktopWindow)
    
    ww = 1600
    hh = Screen.Height / Screen.TwipsPerPixelY
    xoff = (ScaleX(Width, vbTwips, vbPixels) - ScaleWidth)
    yoff = ScaleY(Height, vbTwips, vbPixels) - ScaleHeight
    
    BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, Desktop, left / Screen.TwipsPerPixelX - 4 + xoff, top / Screen.TwipsPerPixelY - 4 + yoff, &HCC0020
    BitBlt Trabalho.hDC, 0, 0, ScaleWidth, ScaleHeight, Trabalho.hDC, 0, 0, vbDstInvert
    BitBlt hDC, 0, 0, ScaleWidth, ScaleHeight, Trabalho.hDC, 0, 0, vbSrcAnd
    Refresh
        
    DoEvents
 
    Picture1.Width = (P_Right - P_Left)
    Picture1.Height = (P_Down - P_Up)
    Picture1.Cls
    SetLayeredWindowAttributes Me.hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
  
    Refresh
    TransparentBlt Picture1.hDC, 0, 0, P_Right - P_Left, P_Down - P_Up, hDC, P_Left, P_Up, P_Right - P_Left, P_Down - P_Up, vbBlack
        
    Picture1.Refresh
        
    Picture1.Picture = Picture1.Image
    SavePicture Picture1.Image, App.Path & "\temp.bmp"
    Picture1.Picture = LoadPicture(App.Path & "\temp.bmp")
    Picture1.Refresh
  
    DoEvents
    
    Image2Clipboard Mask.Picture1.hDC, P_Right - P_Left, P_Down - P_Up
    
    Refresh
    Show
    
    Criou_mascara = True
    
End Sub

Private Sub Form_DblClick()

  Dim i As Integer

    If Nr_pontos > 0 Then
        Load Ponto(Nr_pontos)
    
    End If

    If xx < P_Left Then
        P_Left = xx
    End If
    If xx > P_Right Then
        P_Right = xx
    End If
    If yy < P_Up Then
        P_Up = yy
    End If
    If yy > P_Down Then
        P_Down = yy
    End If

    Ponto(Nr_pontos).Visible = True

    ReDim Preserve Sequencia(Nr_pontos + 3)

    Ponto_atual = Nr_pontos

    For i = 0 To 3
        Sequencia(Ponto_atual + i).x = xx
        Sequencia(Ponto_atual + i).Y = yy
    Next i

    Ponto_atual = Nr_pontos
    Ponto(Ponto_atual).Move xx, yy
    Ligue_os_pontos
    Preencher

    Nr_pontos = Nr_pontos + 1

End Sub

Private Sub Form_DragOver(Source As Control, x As Single, Y As Single, State As Integer)
  
    Source.Move x, Y
    Sequencia(Source.Index).x = x
    Sequencia(Source.Index).Y = Y

    Ligue_os_pontos

    Preencher

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim i As Integer

    If KeyCode = 39 Or KeyCode = 37 Then
        Ponto_atual = Ponto_atual - (KeyCode = 39) + (KeyCode = 37)
        If Ponto_atual > Nr_pontos Then
            Ponto_atual = 0
        End If
        If Ponto_atual < 0 Then
            Ponto_atual = Nr_pontos
        End If

        For i = 0 To Nr_pontos - 1
            If i >= Ponto_atual - 2 And i <= Ponto_atual + 2 Then
                Ponto(i).Visible = True
                Ponto(i).Picture = Cor(0).Picture
              Else
                Ponto(i).Visible = False
            End If
        
        Next i
        Exit Sub
    End If

    If KeyCode = 38 Then

        HScroll1.Value = HScroll1.Value - 5 * (HScroll1.Value < 250)
        Exit Sub
    End If
    If KeyCode = 40 Then
        HScroll1.Value = HScroll1.Value + 5 * (HScroll1.Value > 10)
        Exit Sub
    End If

    If KeyCode = 17 And Shift = 2 Then

        Trabalho.PaintPicture Image, 0, 0, Trabalho.Width, Trabalho.Height, 0, 0, Trabalho.Width, Trabalho.Height, vbSrcCopy

        Clipboard.Clear

        Clipboard.SetData Trabalho.Image

    End If

    If KeyCode = 32 Then
        Form_DblClick
    End If

End Sub

Private Sub Form_Load()

  Dim Normalwindowstyle As Long
  Dim Ret As Long
  Dim col As Long
  Dim i As Integer
  
    Nr_pontos = 0
      
    Normalwindowstyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Normalwindowstyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes Me.hwnd, 0, 50, LWA_ALPHA

    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    HScroll1.Value = 160
    Form_MouseDown 0, 0, 0, 0
    BaseWidth = Width
    BaseHeight = Height
    ReDim Sequencia(0)

    Sequencia(0).x = Ponto(0).left
    Sequencia(0).Y = Ponto(0).top

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim Ret As Long, mbrush As Long, i As Integer

    If Criou_mascara Then
        Criou_mascara = False
        Timer1.Enabled = True
        HScroll1_Change
        Exit Sub
    End If
    
    xx = x
    yy = Y

    If Button = 1 Then
        If Shift = 1 Then
            Timer1.Enabled = Timer1.Enabled Xor -1

            PosX = x: PosY = Y
            mbrush = CreateSolidBrush(&H1)
            SelectObject Trabalho.hDC, mbrush
            Trabalho.ScaleMode = vbPixels
            ExtFloodFill Trabalho.hDC, x, Y, GetPixel(Trabalho.hDC, x, Y), FLOODFILLSURFACE
            DeleteObject mbrush
            Trabalho.Refresh
            Me.PaintPicture Trabalho.Image, 0, 0, Width, Height, 0, 0, Width, Height, vbSrcCopy
        
            P_Up = Screen.Height
            P_Down = 0
            P_Left = Screen.Width
            P_Right = 0
        
            For i = 0 To Nr_pontos
                If Sequencia(i).x < P_Left Then
                    P_Left = Sequencia(i).x
                End If
                If Sequencia(i).x > P_Right Then
                    P_Right = Sequencia(i).x
                End If
                If Sequencia(i).Y < P_Up Then
                    P_Up = Sequencia(i).Y
                End If
                If Sequencia(i).Y > P_Down Then
                    P_Down = Sequencia(i).Y
                End If
            Next i
        
            Do While P_Right > 0
                For i = P_Up To P_Down
                    If Point(P_Right, i) <> &HFFFFFF Then
                        Exit Do
                    End If
                Next i
                P_Right = P_Right - 1
              
            Loop
            Do While P_Left < ScaleWidth
                For i = P_Up To P_Down
                    If Point(P_Left, i) <> &HFFFFFF Then
                        Exit Do
                    End If
                Next i
                P_Left = P_Left + 1
              
            Loop
              
            Do While P_Up < ScaleHeight
                For i = P_Left To P_Right
                    If Point(i, P_Up) <> &HFFFFFF Then
                        Exit Do
                    End If
                Next i
                P_Up = P_Up + 1
              
            Loop
             
            Do While P_Down > 0
                For i = P_Left To P_Right
                    If Point(i, P_Down) <> &HFFFFFF Then
                        Exit Do
                    End If
                Next i
                P_Down = P_Down - 1
              
            Loop
            
            Exit Sub
        End If
            
        If Shift = 2 Then
            Command2_Click
            Exit Sub
        End If
    
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    xx = x
    yy = Y

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    Preencher

End Sub

Private Sub Form_Resize()

    Shape1.Move ScaleWidth - Shape1.Width
    Label1.Move ScaleWidth - Shape1.Width + 10

    FatorX = Width / BaseWidth
    FatorY = Height / BaseHeight
    BaseWidth = Width
    BaseHeight = Height

    Trabalho.Width = ScaleWidth
    Trabalho.Height = ScaleHeight

    If Nr_pontos Then
        Escala
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim i As Integer

    For i = 1 To Len(Ordem)
        F(Asc(Mid$(Ordem, i, 1)) - 1).Visible = True
    Next i
    Main_form.Trabalho.Cls
    Main_form.Trabalho.Picture = LoadPicture(App.Path & "\temp.bmp")
    Main_form.Trabalho.Refresh
    Main_form.Show
    Main_form.Width = (Main_form.Trabalho.ScaleWidth + GetSystemMetrics(SM_CXDLGFRAME)) * Screen.TwipsPerPixelX
    Main_form.Height = (Main_form.Trabalho.ScaleHeight + GetSystemMetrics(SM_CYCAPTION) / 2) * Screen.TwipsPerPixelY
    Main_form.Form_Resize

End Sub

Private Sub HScroll1_Change()

    SetLayeredWindowAttributes Me.hwnd, 0, HScroll1.Value, LWA_ALPHA

End Sub

Private Sub Label1_Click()

    PopupMenu Menu

End Sub

Private Sub Ponto_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    Ponto_atual = Index
    Ponto(Index).Drag

End Sub

Private Sub Ligue_os_pontos()

  Dim i As Integer
  Dim x As Integer

    Trabalho.Cls

    x = (Nr_pontos \ 3) * 3 + 1

    Call PolyBezier(Trabalho.hDC, Sequencia(0), x)

    Me.PaintPicture Trabalho.Image, 0, 0, Width, Height, 0, 0, Width, Height, vbSrcCopy

End Sub

Public Function MakeDWord(LoWord As Single, HiWord As Single) As Long

    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)

End Function

Private Sub Preencher()

  Dim Ret As Long, mbrush As Long, i As Integer
            
Exit Sub
         
    mbrush = CreateSolidBrush(&HC0FFFF)    '6528
    SelectObject Trabalho.hDC, mbrush
    ScaleMode = vbPixels
    ExtFloodFill Trabalho.hDC, 0, 0, GetPixel(Trabalho.hDC, 0, 0), FLOODFILLSURFACE
    DeleteObject mbrush
    Trabalho.Refresh
          
    mbrush = CreateSolidBrush(&HFFFFFE)
    SelectObject Trabalho.hDC, mbrush
    ScaleMode = vbPixels
    ExtFloodFill Trabalho.hDC, ScaleWidth / 2, ScaleHeight / 2, GetPixel(Trabalho.hDC, ScaleWidth / 2, ScaleHeight / 2), FLOODFILLSURFACE
    DeleteObject mbrush
    Trabalho.Refresh
            
    Me.PaintPicture Trabalho.Image, 0, 0, Width, Height, 0, 0, Width, Height, vbSrcCopy
            
End Sub

Private Sub Escala()

  Static i As Integer

    For i = 0 To Nr_pontos - 1
        Sequencia(i).x = Sequencia(i).x * FatorX
        Sequencia(i).Y = Sequencia(i).Y * FatorY
        Ponto(i).Move Sequencia(i).x, Sequencia(i).Y
    Next i
    Ligue_os_pontos

End Sub

Private Sub Timer1_Timer()

  Static passo As Integer
  Static i As Integer
  Static Y As Integer

    If passo = 0 Then
        passo = 1
        Grossura_da_linha = 1
    End If
    Grossura_da_linha = Grossura_da_linha + passo
    If Grossura_da_linha > 2 Or Grossura_da_linha < 2 Then
        passo = -passo
    End If
    Trabalho.DrawWidth = Grossura_da_linha

    Ligue_os_pontos

    Y = Y + 1
    If Y = 3 Then
        Y = 0
    End If

    For i = 0 To Nr_pontos - 1
        If i >= Ponto_atual - 2 And i <= Ponto_atual + 2 Then
            Ponto(i).Visible = True
            Ponto(i).Picture = Cor(Y).Picture
          Else
            Ponto(i).Visible = False
        End If
        
    Next i

End Sub


