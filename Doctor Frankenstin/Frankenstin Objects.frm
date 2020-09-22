VERSION 5.00
Begin VB.Form Objeto 
   Appearance      =   0  'Flat
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFE&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   75
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicCol 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   315
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   2580
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "Objeto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Button_press As Integer
Private Brilho As Integer
Private Saturação As Integer
Private Gamma As Integer

Private Intensidade_de_Brilho As Long
Private Intensidade_de_Saturação As Integer
Private Correção_de_Gamma As Integer

Private Condição_SHIFT As Integer
Private Translucencia As Integer
Private Não_processe As Integer
Private Const WM_MOUSEWHEEL       As Long = &H20A
Private sc          As cSuperClass
Implements iSuperClass

Private BaseWidth As Long
Private BaseHeight As Long

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()

    Main_form.File1.Path = App.Path & "\library\" & Pastas(use_PASTAS(Tag))

End Sub

Private Sub Form_DblClick()

  Dim s As Integer
  Dim ss As Integer
  Dim x As String

    If Button_press <> 1 Then
        Exit Sub
    End If

    Select Case Condição_SHIFT
      Case 0
        s = InStr(Ordem, Chr$(Tag + 1))
        If s = Len(Ordem) Then 'MANDE PRA TRAZ
            Ordem = Chr$(Tag + 1) + Mid$(Ordem, 1, s - 1)
          Else
            Ordem = Mid$(Ordem, 1, s - 1) + Mid$(Ordem, s + 1) + Chr$(Tag + 1)
        End If

      Case 1
        s = InStr(Ordem, Chr$(Tag + 1))
        ss = s - 1
        If ss < 1 Then
            Exit Sub
        End If
        x = Mid$(Ordem, s, 1)
        Mid$(Ordem, s, 1) = Mid$(Ordem, ss, 1)
        Mid$(Ordem, ss, 1) = x
      Case 2
        s = InStr(Ordem, Chr$(Tag + 1))
        ss = s + 1
        If s = Len(Ordem) Then
            Exit Sub
        End If
        x = Mid$(Ordem, s, 1)
        Mid$(Ordem, s, 1) = Mid$(Ordem, ss, 1)
        Mid$(Ordem, ss, 1) = x
        
    End Select

    Libere_Ordem
    Ordene
    Condição_SHIFT = 0

End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim Arquivo As String
  Dim r As Integer, x As Long, Y As Long, i As Integer
  Dim Quadrado_Original As Long
  
  Dim Quadrado_Atual As Long
  
    On Error GoTo erro
    
    Condição_SHIFT = Shift
    Select Case KeyCode
      Case 66 And Shift = 4 'Alt B
        GetCursorPos Pt
        Display.Move Pt.x * Screen.TwipsPerPixelX, Pt.Y * Screen.TwipsPerPixelY
        If Brilho = False Then
            Display.Label1.Caption = "Bright"
            Display.Show , Me: Me.SetFocus
        End If
        Display.Label2.Caption = Intensidade_de_Brilho
        Brilho = True
        Exit Sub
            
      Case 71 And Shift = 4 'Alt G
        Display.Move Pt.x * Screen.TwipsPerPixelX, Pt.Y * Screen.TwipsPerPixelY
        If Gamma = False Then
            Display.Label1.Caption = "Gamma Correction"
            Display.Show , Me: Me.SetFocus
        End If
        Display.Label2.Caption = 100 - Correção_de_Gamma
        Gamma = True
        Exit Sub
            
      Case 84 And Shift = 4 'Alt T
        Display.Move Pt.x * Screen.TwipsPerPixelX, Pt.Y * Screen.TwipsPerPixelY
        If Saturação = False Then
            Display.Label1.Caption = "Saturation"
            Display.Show , Me: Me.SetFocus
        End If
        Display.Label2.Caption = Intensidade_de_Saturação
        Saturação = True
        Exit Sub
        
      Case 65, 83, 87, 90, 107, 109, 255
        If KeyCode <> 255 Then
            Objeto_foi_usado(Tag) = True
        End If
        
        If KeyCode = 107 Then
xxx:
            EscalaZX(Tag) = EscalaZX(Tag) + 0.1
            EscalaZY(Tag) = EscalaZY(Tag) + 0.1
        End If
    
        If KeyCode = 109 Then
            EscalaZX(Tag) = EscalaZX(Tag) - 0.1
            EscalaZY(Tag) = EscalaZY(Tag) - 0.1
        End If
    
        AutoRedraw = True
        x = left: Y = top
        If KeyCode = 255 Then
            GoTo pula
        End If
        
        EscalaZX(Tag) = EscalaZX(Tag) - 0.1 * (KeyCode = 83) + 0.1 * (KeyCode = 65)
        EscalaZY(Tag) = EscalaZY(Tag) - 0.1 * (KeyCode = 87) + 0.1 * (KeyCode = 90)
        If EscalaZX(Tag) <= 0.1 Then
            EscalaZX(Tag) = 0.1
        End If
        If EscalaZY(Tag) <= 0.1 Then
            EscalaZY(Tag) = 0.1
        End If
pula:
        
        Main_form.Original.Picture = LoadPicture(My_path + Pastas(use_PASTAS(Tag)) + "\" + Arquivo_usado_pelo_objeto(Tag))
        
        Main_form.Trabalho.Width = Main_form.Original.Width * EscalaZX(Tag)
        Main_form.Trabalho.Height = Main_form.Original.Height * EscalaZY(Tag)
        Main_form.Trabalho.Cls
        Call StretchBlt(Main_form.Trabalho.hDC, 0, 0, Main_form.Trabalho.Width, Main_form.Trabalho.Height, Main_form.Original.hDC, 0, 0, Main_form.Original.Width, Main_form.Original.Height, vbSrcCopy)
        
        Main_form.Original.Picture = LoadPicture("")
        
        Não_processe = True
        Quadrado_Original = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
        PicCol.Width = Main_form.Trabalho.Width
        PicCol.Height = Main_form.Trabalho.Height
        Quadrado_Atual = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
        Não_processe = False
        
        PicCol_Resize
        PicCol.Cls
        AutoRedraw = True
        Call StretchBlt(PicCol.hDC, 0, 0, Main_form.Trabalho.Width, Main_form.Trabalho.Height, Main_form.Trabalho.hDC, 0, 0, Main_form.Trabalho.Width, Main_form.Trabalho.Height, vbSrcCopy)
        Pintar
        AutoRedraw = False
        Refresh
        Move left - ((Quadrado_Atual - Quadrado_Original) / 2) * Screen.TwipsPerPixelX, top - ((Quadrado_Atual - Quadrado_Original) / 2) * Screen.TwipsPerPixelY
        Exit Sub
        
      Case 38, 40
        Objeto_foi_usado(Tag) = True
        If Shift Then
            Translucencia = (Translucencia - 5 * (KeyCode = 38) + 5 * (KeyCode = 40))
            If Translucencia < 50 Then
                Translucencia = 50
            End If
            If Translucencia > 255 Then
                Translucencia = 255
            End If
            SetLayeredWindowAttributes Me.hwnd, 255, Translucencia, LWA_COLORKEY Or LWA_ALPHA
            Exit Sub
        End If

        Ponteiro_de_arquivo(Tag) = Ponteiro_de_arquivo(Tag) - (KeyCode = 38) + (KeyCode = 40)
        If Ponteiro_de_arquivo(Tag) < 0 Then
            Ponteiro_de_arquivo(Tag) = Main_form.File1.ListCount - 1
        End If
        If Ponteiro_de_arquivo(Tag) > Main_form.File1.ListCount - 1 Then
            Ponteiro_de_arquivo(Tag) = 0
        End If

        Arquivo_usado_pelo_objeto(Tag) = Main_form.File1.List(Ponteiro_de_arquivo(Tag))
        
        Arquivo = Arquivo_usado_pelo_objeto(Tag)
        
        EscalaZX(Tag) = EscalaZX(Tag) - 0.1
        EscalaZY(Tag) = EscalaZY(Tag) - 0.1
        GoTo xxx
    
      Case 39
        Objeto_foi_usado(Tag) = True
        tel(Tag) = tel(Tag) + 1
        If tel(Tag) = 360 Then
            tel(Tag) = 0
        End If
        Pintar
        
      Case 37
        Objeto_foi_usado(Tag) = True
        tel(Tag) = tel(Tag) - 1
        If tel(Tag) = 0 Then
            tel(Tag) = 360
        End If
        Pintar

      Case 27
        End
  
    End Select
sair:

Exit Sub

erro:
        
    r = MsgBox(Error, vbCritical)
    Resume sair

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Condição_SHIFT = 0

    Brilho = False
    Saturação = False
    Gamma = False
    Display.Hide

End Sub

Private Sub Form_Load()

  Dim Normalwindowstyle As Long
  Dim Ret As Long, xx As Integer, n As String
  Dim col As Long
  Dim i As Integer
  
    Set sc = New cSuperClass
  
    With sc
        Call .AddMsg(WM_MOUSEWHEEL)
        Call .Subclass(hwnd, Me)
    End With
  
    Translucencia = 255
    
    Normalwindowstyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    col = 255 'RGB(255, 255, 255)
    SetLayeredWindowAttributes Me.hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
       
    Call SetStretchBltMode(hDC, STRETCHMODE)
    
    Correção_de_Gamma = 100
    
End Sub

Public Function MakeDWord(LoWord As Single, HiWord As Single) As Long

  'MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim returnval As Long
    
    Button_press = Button
    Selected = Tag
    
    Main_form.Trabalho.Width = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    Main_form.Trabalho.Height = Main_form.Trabalho.Width
      
    If Button = 1 Then
        
        xx = x * Screen.TwipsPerPixelX: yy = Y * Screen.TwipsPerPixelY
        capture = True
        ReleaseCapture
        SetCapture Me.hwnd
    End If

    If Button = 2 Then
        PopupMenu Main_form.Menu
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If capture Then
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - xx, Pt.Y * Screen.TwipsPerPixelY - yy
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False

End Sub

Private Sub Form_Paint()

  Static vez As Integer

    If vez = 0 Then
        vez = 1
        Cls
        Rotate Me.hDC, ScaleWidth / 2, ScaleHeight / 2, tel(Tag), PicCol.hDC, 0, 0, PicCol.Width, PicCol.Height
        AutoRedraw = True
    
        Rotate Me.hDC, ScaleWidth / 2, ScaleHeight / 2, tel(Tag), PicCol.hDC, 0, 0, PicCol.Width, PicCol.Height
        AutoRedraw = False
    
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set sc = Nothing

End Sub

Private Sub PicCol_Resize()

    If Não_processe Then
        Exit Sub
    End If

    Main_form.Trabalho.Width = Sqr(PicCol.Width ^ 2 + PicCol.Height ^ 2)
    Main_form.Trabalho.Height = Main_form.Trabalho.Width
    Width = Main_form.Trabalho.ScaleWidth * Screen.TwipsPerPixelX
    Height = Main_form.Trabalho.ScaleHeight * Screen.TwipsPerPixelY
 
    BaseWidth = Width
    BaseHeight = Height

End Sub

Private Sub Pintar()
    
    Main_form.Trabalho.Cls
    Rotate Main_form.Trabalho.hDC, Main_form.Trabalho.ScaleWidth / 2, Main_form.Trabalho.ScaleHeight / 2, tel(Tag), PicCol.hDC, 0, 0, PicCol.Width, PicCol.Height
    Call StretchBlt(hDC, 0, 0, Int(Main_form.Trabalho.Width * EscalaX(Tag)), Int(Main_form.Trabalho.Height * EscalaY(Tag)), Main_form.Trabalho.hDC, 0, 0, Main_form.Trabalho.Width, Main_form.Trabalho.Height, vbSrcCopy)
              
End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
  
  'Case WM_MOUSEWHEEL
  
    Select Case wParam
      Case -7864308
        'Debug.Print "UP+Shift+Control";
        Form_KeyDown 40, 1
      Case 7864332
        'Debug.Print "DOWN+Shift+Control";
        Form_KeyDown 38, 1
      Case 7864324
        'Debug.Print "UP+Shift";
        Form_KeyDown 40, 0
      Case -7864316
        'Debug.Print "DOWN+Shift";
        Form_KeyDown 38, 0
      Case -7864320
        If Brilho Then
            Intensidade_de_Brilho = Intensidade_de_Brilho + 5
            If Intensidade_de_Brilho > 100 Then
                Intensidade_de_Brilho = 100
            End If
            Fazer_Brilho
            Exit Sub
        End If
        
        If Saturação Then
            Intensidade_de_Saturação = Intensidade_de_Saturação + 5
            If Intensidade_de_Saturação > 100 Then
                Intensidade_de_Saturação = 100
            End If
            Fazer_Brilho
            Exit Sub
        End If
        
        If Gamma Then
            Correção_de_Gamma = Correção_de_Gamma - 5
            If Correção_de_Gamma < 5 Then
                Correção_de_Gamma = 5
            End If
            Fazer_Brilho
            Exit Sub
        End If
                
        Form_KeyDown 107, 0
        'Debug.Print "DOWN";
        
      Case 7864320
        If Brilho Then
            Intensidade_de_Brilho = Intensidade_de_Brilho - 5
            If Intensidade_de_Brilho < 0 Then
                Intensidade_de_Brilho = 0
            End If
            Fazer_Brilho
            Exit Sub
        End If
       
        If Saturação Then
            Intensidade_de_Saturação = Intensidade_de_Saturação - 5
            If Intensidade_de_Saturação < 0 Then
                Intensidade_de_Saturação = 0
            End If
            Fazer_Brilho 'Saturação
            Exit Sub
        End If
        
        If Gamma Then
            Correção_de_Gamma = Correção_de_Gamma + 5
            If Correção_de_Gamma > 100 Then
                Correção_de_Gamma = 100
            End If
            Fazer_Brilho
            Exit Sub
        End If
        
        Form_KeyDown 109, 0
        'Debug.Print "UP";
      Case -7864312
        Form_KeyDown 39, 0
        'Debug.Print "DOWN + Control";
      Case 7864328
        Form_KeyDown 37, 0
        'Debug.Print "UP + Control";
    End Select
        
    'Debug.Print wParam;

End Sub

Private Sub Fazer_Brilho()

    frmFilters.Picture1.Width = ScaleWidth
    frmFilters.Picture1.Height = ScaleHeight
    frmFilters.Picture1.PaintPicture Image, 0, 0, ScaleWidth, ScaleHeight, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy
    Call FilterG(iBRIGHTNESS, frmFilters.Picture1.Image, Intensidade_de_Brilho, mProgress)
    Call FilterG(iSATURATION, frmFilters.Picture1.Image, Intensidade_de_Saturação, mProgress)
    Call FilterG(iGAMMA, frmFilters.Picture1.Image, Correção_de_Gamma, mProgress)
    
    Call BitBlt(hDC, 0, 0, frmFilters.Picture1.ScaleWidth, frmFilters.Picture1.ScaleHeight, frmFilters.Picture1.hDC, 0, 0, SRCCOPY)

End Sub

Private Sub Fazer_Saturação()
  
End Sub


