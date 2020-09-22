VERSION 5.00
Begin VB.Form Main_form 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3915
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   4755
   ControlBox      =   0   'False
   Icon            =   "Frankenstin Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   261
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   317
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   2535
      Pattern         =   "*.bmp;*.gif;*jpg"
      TabIndex        =   4
      Top             =   1575
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1785
      Left            =   660
      Picture         =   "Frankenstin Main.frx":164A
      ScaleHeight     =   119
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   119
      TabIndex        =   3
      Top             =   1545
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   75
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1485
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox Original 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   465
      Left            =   1845
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   110
      TabIndex        =   1
      Top             =   2895
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Trabalho 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   0
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   38
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   3  'Dot
      Height          =   15
      Left            =   3645
      Top             =   0
      Width           =   15
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   4  'Dash-Dot
      Height          =   15
      Left            =   3630
      Top             =   0
      Width           =   15
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu_options 
         Caption         =   "Clone"
         Index           =   0
      End
      Begin VB.Menu menu_options 
         Caption         =   "New from"
         Index           =   1
         Begin VB.Menu From_Library 
            Caption         =   "Library"
            Index           =   0
            Begin VB.Menu Itens 
               Caption         =   "Itens"
               Index           =   0
            End
         End
         Begin VB.Menu From_Library 
            Caption         =   "Files"
            Index           =   1
         End
      End
      Begin VB.Menu menu_options 
         Caption         =   "New Project"
         Index           =   2
      End
      Begin VB.Menu menu_options 
         Caption         =   "Open"
         Index           =   3
      End
      Begin VB.Menu menu_options 
         Caption         =   "Save"
         Index           =   4
      End
      Begin VB.Menu menu_options 
         Caption         =   "Delete"
         Index           =   5
      End
      Begin VB.Menu menu_options 
         Caption         =   "Show as Picture"
         Index           =   6
      End
      Begin VB.Menu menu_options 
         Caption         =   "Mask Maker"
         Index           =   7
      End
      Begin VB.Menu menu_options 
         Caption         =   "End"
         Index           =   8
      End
   End
   Begin VB.Menu Menu_principal 
      Caption         =   "Menu Principal"
      Visible         =   0   'False
      Begin VB.Menu Files 
         Caption         =   "File"
         Begin VB.Menu Open 
            Caption         =   "Open"
         End
         Begin VB.Menu Save 
            Caption         =   "Save"
         End
         Begin VB.Menu Put_on_library 
            Caption         =   "Put on the Lybrary"
            Begin VB.Menu Itens_to_put 
               Caption         =   "Itens"
               Index           =   0
            End
         End
      End
      Begin VB.Menu Edit 
         Caption         =   "Edit"
         Begin VB.Menu Copy 
            Caption         =   "Copy"
            Shortcut        =   ^C
         End
         Begin VB.Menu Paste 
            Caption         =   "Paste"
            Shortcut        =   ^V
         End
         Begin VB.Menu Crop 
            Caption         =   "Crop"
         End
      End
      Begin VB.Menu Retouch 
         Caption         =   "Retouch Window"
      End
      Begin VB.Menu Maskmaker 
         Caption         =   "Mask Maker"
      End
      Begin VB.Menu Preferences 
         Caption         =   "Preferences"
         Begin VB.Menu Standard 
            Caption         =   "Resize to Standard when put on the Library"
            Checked         =   -1  'True
         End
         Begin VB.Menu Set_standard 
            Caption         =   "Set the Standard of the library"
            Begin VB.Menu Set_library 
               Caption         =   "Itens"
               Index           =   0
            End
         End
      End
   End
End
Attribute VB_Name = "Main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private capture As Integer
Private xxxx As Single
Private yyyy As Single

Private Vx1 As Single
Private vy1 As Single
Private Type arq
    Pastas As String
    Nome_do_arq As String
    left As Single
    top As Single
    EscalaZX As Single
    EscalaZY As Single
    EscalaX As Single
    EscalaY As Single
    tel As Single
End Type

Private Sub Command1_Click()

End Sub

Private Sub Copy_Click()

  Dim retval As Long

    Trabalho.Width = ScaleWidth
    Trabalho.Height = ScaleHeight
    Trabalho.Cls

    Call SetStretchBltMode(hDC, STRETCHMODE)
    retval = StretchBlt(Trabalho.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)
    
    Clipboard.Clear

    Clipboard.SetData Trabalho.Image

End Sub

Private Sub Crop_Click()

  Dim xoff As Single
  Dim yoff As Single
  
    Trabalho.Cls
    Trabalho.Width = Shape1.Width
    Trabalho.Height = Shape1.Height
    Call SetStretchBltMode(Trabalho.hDC, STRETCHMODE)
    StretchBlt Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, hDC, Shape1.left, Shape1.top, Trabalho.ScaleWidth, Trabalho.ScaleHeight, vbSrcCopy
    Trabalho.Refresh

    StretchBlt hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, vbSrcCopy

    xoff = (ScaleX(Width, vbTwips, vbPixels) - ScaleWidth)
    yoff = ScaleY(Height, vbTwips, vbPixels) - ScaleHeight
                
    Width = (Trabalho.Width + xoff) * Screen.TwipsPerPixelX
    Height = (Trabalho.Height + yoff) * Screen.TwipsPerPixelY
    Shape1.Move 0, 0, 1, 1
    Shape2.Move 0, 0, 1, 1

End Sub

Private Sub Form_Load()

  Dim i As Integer

    Prepare_pastas
    For i = 0 To Qt_pastas
        If i > 0 Then
            Load Itens(i)
            Load Itens_to_put(i)
            Load Set_library(i)
        End If
        Itens(i).Caption = Pastas(i)
        Itens_to_put(i).Caption = Pastas(i)
        Set_library(i).Caption = Pastas(i)
    Next i
 
    free = FreeFile
    If Dir$(App.Path & "\Projects\Last file.Frankenstin") <> "" Then
        Open App.Path & "\Projects\Last file.Frankenstin" For Input As free
        Input #free, Arquivo
        Close free
    End If
    
    If Arquivo <> "" Then
        Abrir_arquivo (Arquivo)
      Else
        Novo_Projeto
    End If
 
End Sub

Private Sub Form_LostFocus()

    Shape1.Width = 1
    Shape1.Height = 1
    Shape2.Width = 1
    Shape2.Height = 1

    Cls
    Trabalho.Picture = LoadPicture("")
    Hide

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 1 And Shift = 1 Then
    
        Vx1 = x: vy1 = Y
        Shape1.Move x, Y, 0, 0
        Shape2.Move x, Y, 0, 0
        Exit Sub
    End If
    
    If Button = 2 Then
        PopupMenu Menu_principal
        Exit Sub
    End If

    xxxx = x * Screen.TwipsPerPixelX: yyyy = Y * Screen.TwipsPerPixelY
    capture = True
    ReleaseCapture
    SetCapture Me.hwnd

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

  Dim w As Single, h As Single
  Dim xx As Single, yy As Single

    If x < Vx1 Then
        xx = x - Vx1
    End If

    If Y < vy1 Then
        yy = Y - vy1
    End If

    w = x - Vx1
    h = Y - vy1

    If Button = 1 And Shift Then
        Shape1.Move Vx1 + xx, vy1 + yy, Abs(w), Abs(h)
        Shape2.Move Vx1 + xx, vy1 + yy, Abs(w), Abs(h)
        Exit Sub
    End If

    If capture Then
        DoEvents
        GetCursorPos Pt
        Move Pt.x * Screen.TwipsPerPixelX - xxxx, Pt.Y * Screen.TwipsPerPixelY - yyyy
        Exit Sub
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    capture = False

End Sub

Public Sub Form_Resize()

  Dim retval As Long

    Call SetStretchBltMode(hDC, STRETCHMODE)
    retval = StretchBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, vbSrcCopy)
    Refresh

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim i As Integer

    For i = 0 To Quant_Objetos - 1
        Unload F(i)
    Next i
    Unload Splash_Form
    Unload Form1
    Unload Form2
    Unload Display
    Unload Mask
    Unload frmFilters
    End

End Sub

Private Sub Mask_maker_Click()

    Mask.Show

End Sub

Public Sub From_files_Click()

  Dim i As Integer
  Dim Nome_Arquivo As String

    Arquivo = OpenDialog(Me, "All types|*.bmp;*.jpg;*.gif|Bitmap|*.bmp|Compuserv|*.gif|JPeg|*.jpg", "Doctor Frankenstin Open", "")
    If Arquivo <> "" Then
        If Dir$(App.Path & "\Library\Files", vbDirectory) = "" Then
            MkDir App.Path & "\Library\Files"
        End If
        For i = Len(Arquivo) To 1 Step -1
            If Mid$(Arquivo, i, 1) = "\" Then
                Exit For
            End If
        Next i
        
        Nome_Arquivo = Mid$(Arquivo, i + 1)
    
        FileCopy Arquivo, App.Path & "\Library\Files\" & Mid$(Arquivo, i + 1)
    
        Main_form.File1.Path = App.Path & "\Library\Files"
        For i = 0 To Qt_pastas
            If Pastas(i) = "Files" Then
                GoTo oK
            End If
        Next i
        Pastas(i) = "Files"
        Load Itens(i)
        Itens(i).Caption = "Files"
oK:
        Itens_Click (i)
        Arquivo_usado_pelo_objeto(Quant_Objetos - 1) = Nome_Arquivo
        For i = 0 To Main_form.File1.ListCount - 1
            If Main_form.File1.List(i) = Nome_Arquivo Then
                Exit For
            End If
        Next i
        Ponteiro_de_arquivo(Quant_Objetos - 1) = i - 1
        F(Quant_Objetos - 1).Form_KeyDown 38, 0
    
    End If

End Sub

Private Sub From_Library_Click(Index As Integer)

    If Index = 1 Then
        From_files_Click
    End If

End Sub

Private Sub Itens_Click(Index As Integer)

  Dim x As Integer

    Crie_nova_instancia
    x = Quant_Objetos - 1
            
    use_PASTAS(x) = Index
    Arquivo_usado_pelo_objeto(x) = Dir$(My_path & Pastas(Index) & "\*.*")
    Ponteiro_de_arquivo(x) = 0
    EscalaX(x) = 1
    EscalaY(x) = 1
    EscalaZX(x) = 1
    EscalaZY(x) = 1
    tel(x) = 0
        
    F(x).PicCol.Picture = LoadPicture(App.Path & "\Library\" & Pastas(Index) & "\" & Arquivo_usado_pelo_objeto(x))
    F(x).Tag = x
    F(x).Visible = True
    F(x).Form_KeyDown 255, 0
    F(x).Move (Screen.Width - F(x).Width) / 2, (Screen.Height - F(x).Height) / 2
        
    Libere_Ordem
    Ordene
    Objeto_foi_usado(x) = True

End Sub

Private Sub Itens_to_put_Click(Index As Integer)

  Dim x As Integer

    If Standard.Checked Then
        Arquivo = App.Path & "\Library\" & Itens(Index).Caption & "\Standard.txt"
        If Dir$(Arquivo) = "" Then
            Definir_valor_Padrão (Index)
        End If
        free = FreeFile
        Open Arquivo For Input As free
        Input #free, x
        x = x + (GetSystemMetrics(SM_CXDLGFRAME) + 1) * 2
        Close free
        If x <= 0 Then
            GoTo continua
        End If
        Height = (Height * (x * Screen.TwipsPerPixelX)) / Width
        Width = x * Screen.TwipsPerPixelX
    End If

continua:

    Trabalho.Cls
    Trabalho.Width = ScaleWidth
    Trabalho.Height = ScaleHeight
    Call SetStretchBltMode(Trabalho.hDC, STRETCHMODE)
    StretchBlt Trabalho.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy
    Trabalho.Refresh
    SavePicture Trabalho.Image, App.Path & "\library\" & Itens(Index).Caption & "\" & Format$(Now, "mmddyy-hhmmss") & ".bmp"

End Sub

Private Sub Maskmaker_Click()

    Mask.Show

End Sub

Private Sub Menu_Click()

    menu_options(0).Caption = "Clone " & Pastas(use_PASTAS(Selected))

End Sub

Private Sub Open_Click()

  Static Ultimo_dir As String

    Arquivo = OpenDialog(Me, "*.bmp|*jpg|*.gif", "Open Picture", Ultimo_dir)
    If Arquivo = "" Then
        Exit Sub
    End If
    Ultimo_dir = CurDir

    Trabalho.Picture = LoadPicture(Arquivo)
    Trabalho.Refresh
    Width = Trabalho.ScaleWidth * Screen.TwipsPerPixelX
    Height = Trabalho.ScaleHeight * Screen.TwipsPerPixelY

End Sub

Private Sub Paste_Click()

  Dim xoff As Single
  Dim yoff As Single
                
    Trabalho.Cls

    Trabalho.Picture = Clipboard.GetData
    If WindowState <> 0 Then
        WindowState = 0
    End If

    xoff = (ScaleX(Width, vbTwips, vbPixels) - ScaleWidth)
    yoff = ScaleY(Height, vbTwips, vbPixels) - ScaleHeight
                
    Width = (Trabalho.Width + xoff) * Screen.TwipsPerPixelX
    Height = (Trabalho.Height + yoff) * Screen.TwipsPerPixelY
                
End Sub

Public Sub Menu_options_Click(Index As Integer)

  Dim s As Integer, x As Long, i As Integer, r As Integer
  Dim retval As Long
  Dim k As Integer
  Dim xoff As Single
  Dim yoff As Single
  Dim h As String * 255
  
    ' On Error GoTo erro
   
    Select Case Index
        
      Case 0 'CLONE
        
        Crie_nova_instancia
        x = Quant_Objetos - 1
            
        use_PASTAS(x) = use_PASTAS(Selected)
        Arquivo_usado_pelo_objeto(x) = Arquivo_usado_pelo_objeto(Selected)
        Ponteiro_de_arquivo(x) = Ponteiro_de_arquivo(Selected)
        EscalaX(x) = EscalaX(Selected)
        EscalaY(x) = EscalaY(Selected)
        EscalaZX(x) = EscalaZX(Selected)
        EscalaZY(x) = EscalaZY(Selected)
        tel(x) = tel(Selected)
        
        F(x).PicCol.Picture = F(Selected).PicCol.Picture
        F(x).Tag = x
        F(x).Visible = True
        F(x).Form_KeyDown 255, 0
        F(x).Move F(Selected).left + F(Selected).Width / 4, F(Selected).top + F(Selected).Height / 4
        
        Libere_Ordem
        Ordene
        
      Case 1 'NEW FROM
        
      Case 2 'NEW PROJECT
        r = MsgBox("Do you are sure?", vbYesNo Or vbQuestion)
        If r = 7 Then
            Exit Sub
        End If
        
        Novo_Projeto
        
      Case 3 'OPEN
        Arquivo = OpenDialog(Me, "Frankenstin files|*.Frankenstin", "Doctor Frankenstin Open", App.Path & "\Projects")
        If Arquivo <> "" Then
            Abrir_arquivo (Arquivo)
        End If
        
      Case 4 'SAVE
        Arquivo = SaveDialog(Me, "Frankenstin files|*.Frankenstin", ".Frankenstin", "Doctor Frankenstin Save", App.Path & "\Projects")
        If Arquivo <> "" Then
            Salvar (Arquivo)
        End If
        
      Case 5 'DELETE
        s = InStr(Ordem, Chr$(Selected + 1))
        If s = 0 Then
            Exit Sub
        End If
        If Len(Ordem) > 1 Then
            Libere_Ordem
            Unload F(Selected)
            Ordem = Mid$(Ordem, 1, s - 1) + Mid$(Ordem, s + 1)
            Ordene
        End If
      Case 6 'SHOW
        Trabalho.Width = F(Selected).ScaleWidth
        Trabalho.Height = F(Selected).ScaleHeight
        Trabalho.Cls
  
        For k = 1 To Len(Ordem)
            i = Asc(Mid$(Ordem, k, 1)) - 1
            DoEvents
            TransparentBlt Trabalho.hDC, _
                           (F(i).left - F(Selected).left) / Screen.TwipsPerPixelX, _
                           (F(i).top - F(Selected).top) / Screen.TwipsPerPixelY, _
                           F(i).ScaleWidth, _
                           F(i).ScaleHeight, _
                           F(i).hDC, 0, 0, _
                           F(i).ScaleWidth, _
                           F(i).ScaleHeight, _
                           vbRed
                
            Trabalho.Refresh
            Call SetStretchBltMode(hDC, STRETCHMODE)
            retval = StretchBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, Trabalho.hDC, 0, 0, Trabalho.ScaleWidth, Trabalho.ScaleHeight, vbSrcCopy)
            Refresh
        Next k
                
        xoff = (ScaleX(Width, vbTwips, vbPixels) - ScaleWidth)
        yoff = ScaleY(Height, vbTwips, vbPixels) - ScaleHeight
                
        Width = (Trabalho.Width + xoff) * Screen.TwipsPerPixelX
        Height = (Trabalho.Height + yoff) * Screen.TwipsPerPixelY
                      
        Me.Show
      
      Case 7
        For i = 1 To Len(Ordem)
            F(Asc(Mid$(Ordem, i, 1)) - 1).Visible = False
        Next i
              
        Mask_maker_Click
      Case 8 'EXIT
        Unload Me
            
    End Select

sair:

Exit Sub

erro:

    Resume sair

End Sub

Public Sub Abrir_arquivo(Arquivo)

  Dim i As Integer
  Dim k As Integer
  Dim m As Integer
  Dim A As arq
  Dim qt_obj As Integer
  Dim novo_Quant_obj As Integer
  Dim j As Integer
  Dim r As Integer
  Dim h As String * 255
  
    For i = 0 To Quant_Objetos - 1
        Unload F(i)
        DoEvents
    Next i

    ReDim F(0)
  
    free = FreeFile
    Open Arquivo For Binary As free
    Get #free, 1, h
    Ordem = Mid$(h, 1, InStr(h, Chr$(255)) - 1)
    Quant_Objetos = Len(Ordem)
    For i = 0 To Len(Ordem) - 1
        Get #free, , A
    'Debug.Print A.Pastas; " "; A.Nome_do_arq
    
        'Verifique se o Arquivo existe
        For k = 0 To Qt_pastas
            If Pastas(k) = A.Pastas Then
                File1.Path = App.Path & "\Library\" & A.Pastas
                If File1.ListCount = 0 Then 'Pastas existe mas está vazia
                    GoTo salte
                End If
                For m = 0 To File1.ListCount - 1
                    If A.Nome_do_arq = File1.List(m) Then
                        GoTo continue
                    End If
                Next m
                'Pastas existe, não está vazia, mas o arquivo não existe
                GoTo siga
            End If
        Next k
        MkDir App.Path & "\Library\" & A.Pastas 'Não Existe a Pasta nem o arquivo
salte:
        FileCopy App.Path & "\Error.bmp", App.Path & "\Library\" & A.Pastas & "\Error.bmp"
        
siga:
        r = MsgBox("The file: " & UCase$(A.Nome_do_arq) & vbCrLf & "form Library: " & UCase$(A.Pastas) & vbCrLf & "does not exist." & vbCrLf & "It was changed to: " & vbCrLf & UCase$(File1.List(0)) & vbCrLf & "from Library: " & UCase$(A.Pastas), vbCritical)
                
        A.Nome_do_arq = File1.List(0)
        m = 0
continue:
        
        ReDim Preserve F(i)
        Ponteiro_de_arquivo(i) = m
        F(i).Tag = i
        use_PASTAS(i) = k
        EscalaZX(i) = A.EscalaZX
        EscalaZY(i) = A.EscalaZY
        EscalaX(i) = A.EscalaX
        EscalaY(i) = A.EscalaY
        SetLayeredWindowAttributes F(i).hwnd, 255, 0, LWA_COLORKEY Or LWA_ALPHA
        Arquivo_usado_pelo_objeto(i) = A.Nome_do_arq
        F(i).PicCol.Picture = LoadPicture(App.Path & "\Library\" & A.Pastas & "\" & A.Nome_do_arq)
        tel(i) = A.tel
        
        F(i).Show
        F(i).Form_KeyDown 255, 0
        F(i).Move A.left, A.top
        F(i).Refresh
        Objeto_foi_usado(i) = True
proximo:
    Next i

    Libere_Ordem
    Ordene

    For i = 0 To Quant_Objetos - 1
        SetLayeredWindowAttributes F(i).hwnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
    Next i

End Sub

Public Sub Salvar(Arquivo)

  Dim x As Single
  Dim i As Integer
  Dim A As arq
  Dim k As Integer
  Dim nova_ordem As String
  Dim h As String * 255
  Dim Bkp_ordem As String
  
    Bkp_ordem = Ordem
  
novamente:
    For i = 1 To Len(Ordem)
        If Objeto_foi_usado(Asc(Mid$(Ordem, i, 1)) - 1) = False Then
            Ordem = Mid$(Ordem, 1, i - 1) + Mid$(Ordem, i + 1)
            GoTo novamente
        End If
    Next i
    If Ordem = "" Then
        i = MsgBox("Not any object was used. Please use one and save again.", vbCritical)
        Ordem = Bkp_ordem
        Exit Sub
    End If
  
    List1.Clear
    For i = 1 To Len(Ordem)
        List1.AddItem Mid$(Ordem, i, 1)
        List1.ItemData(List1.NewIndex) = i
    Next i

    For i = 0 To Len(Ordem) - 1
        nova_ordem = nova_ordem + Chr$(List1.ItemData(i))
    Next i
    
    free = FreeFile
    Open Arquivo For Binary As free
      
    h = nova_ordem + Chr$(255)
    
    Put #free, , h
    For i = 0 To Len(Ordem) - 1
        k = Asc(Mid$(Ordem, i + 1, 1)) - 1
        With A
            .Pastas = Pastas(use_PASTAS(k))
            .Nome_do_arq = Arquivo_usado_pelo_objeto(k)
            .left = F(k).left
            .top = F(k).top
            .EscalaZX = EscalaZX(k)
            .EscalaZY = EscalaZY(k)
            .EscalaX = EscalaX(k)
            .EscalaY = EscalaY(k)
            .tel = tel(k)
        End With
        Put #free, , A
    Next i
    Close free
    
    free = FreeFile
    Open App.Path & "\Projects\Last file.Frankenstin" For Output As free
    Write #free, Arquivo
    Close free
    Ordem = Bkp_ordem

End Sub

Private Sub Retouch_Click()

  Dim retval As Long

    Trabalho.Width = Width / Screen.TwipsPerPixelX
    Trabalho.Height = Height / Screen.TwipsPerPixelY
    Trabalho.Cls

    Call SetStretchBltMode(hDC, STRETCHMODE)
    frmFilters.Picture1.Width = ScaleWidth
    frmFilters.Picture1.Height = ScaleHeight
    retval = StretchBlt(frmFilters.Picture1.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)

    frmFilters.Picture3.Width = ScaleWidth
    frmFilters.Picture3.Height = ScaleHeight
    retval = StretchBlt(frmFilters.Picture3.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)

    frmFilters.Prepare
    frmFilters.Move left, top, Width, Height
    frmFilters.Show
    frmFilters.SetFocus
    
End Sub

Private Sub Novo_Projeto()

  Dim i As Integer
  Dim Arquivo As String
  Dim Tipo_Anterior As String
  Dim tw As Long, th As Long
  Dim PosX As Long, PosY As Long
  Dim cx As Single
  Dim cy As Single
  
    Erase Objeto_foi_usado
  
    For i = 0 To Quant_Objetos - 1
        Unload F(i)
 
    Next i
    Erase F
  
    Quant_Objetos = 0
    Qt_pastas = 0
    Ordem = ""
    
    Erase Ponteiro_de_arquivo
  
    PosX = 100 * Screen.TwipsPerPixelX
  
    Prepare_pastas

    For i = 0 To Qt_pastas
        File1.Path = App.Path & "\Library\" & Pastas(i)
        Arquivo = File1.List(0)
        If Arquivo <> "" Then

            Arquivo_usado_pelo_objeto(Quant_Objetos) = Arquivo
            Crie_nova_instancia
       
            Original.Picture = LoadPicture(My_path & Pastas(i) & "\" & Arquivo)
            If Original.Width > Original.Height Then
                tw = 80
                th = Original.Height * 80 / Original.Width
              Else
                th = 80
                tw = Original.Width * 80 / Original.Height
            End If
        
            Trabalho.Width = Trabalho.TextWidth(UCase$(Pastas(i)))
            If Trabalho.Width < Picture1.Width Then
                Trabalho.Width = Picture1.Width
            End If
        
            Trabalho.Height = Picture1.Height + Trabalho.TextHeight(UCase$(Pastas(i)))
            Trabalho.Cls
            Picture1.Cls
            GdiTransparentBlt Picture1.hDC, (Picture1.ScaleWidth - tw) / 2, (Picture1.ScaleHeight - th) / 2, tw, th, Original.hDC, 0, 0, Original.ScaleWidth, Original.ScaleHeight, vbRed
            Picture1.Refresh
            GdiTransparentBlt Trabalho.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, vbRed
            Trabalho.CurrentX = (Trabalho.ScaleWidth - Trabalho.TextWidth(UCase$(Pastas(i)))) / 2
            Trabalho.CurrentY = Picture1.ScaleHeight
            cx = Trabalho.CurrentX
            cy = Trabalho.CurrentY
            Trabalho.ForeColor = vbWhite
            Trabalho.Print UCase$(Pastas(i))
            Trabalho.CurrentX = cx + 1
            Trabalho.CurrentY = cy + 1
            Trabalho.ForeColor = &H80000008
            Trabalho.Print UCase$(Pastas(i))
            
            SavePicture Trabalho.Image, App.Path & "\temp.bmp"
            F(i).PicCol.Picture = LoadPicture(App.Path & "\temp.bmp")
            F(i).Tag = i
            use_PASTAS(i) = i
        
            EscalaX(i) = 1
            EscalaY(i) = 1
            EscalaZX(i) = 1
            EscalaZY(i) = 1
            tel(i) = 0
        
            F(i).Move PosX - (F(i).Width) / 2, PosY
            PosY = PosY + (Picture1.ScaleHeight + Trabalho.TextHeight(UCase$("A"))) * Screen.TwipsPerPixelY
            If PosY > Screen.Height - 1000 Then
                PosY = 0
                PosX = PosX + (Trabalho.ScaleWidth + 20) * Screen.TwipsPerPixelX
                If PosX > Screen.Width Then
                    PosX = 0
                End If
            End If
                       
            Ponteiro_de_arquivo(i) = 0
        End If
    Next i
   
    Ordene
    Libere_Ordem
    
    Trabalho.Cls
    Original.Picture = LoadPicture("")
    Hide

End Sub

Private Sub Crie_nova_instancia()

    ReDim Preserve F(Quant_Objetos)
    Quant_Objetos = Quant_Objetos + 1
    Ordem = Ordem + Chr$(Quant_Objetos)

End Sub

Private Sub Save_Click()

    Arquivo = SaveDialog(Me, "*.bmp", ".bmp", "Save Picture", "")

    Trabalho.Cls
    Trabalho.Width = ScaleWidth
    Trabalho.Height = ScaleHeight
    Call SetStretchBltMode(Trabalho.hDC, STRETCHMODE)
    StretchBlt Trabalho.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy
    Trabalho.Refresh
    SavePicture Trabalho.Image, Arquivo

End Sub

Private Sub Definir_valor_Padrão(Index)

  Dim x As Integer
  Dim free As Integer

    Arquivo = App.Path & "\Library\" & Itens(Index).Caption & "\Standard.txt"
    x = InputBox("Please, enter the Standard Width for this Library" & vbCrLf & "Use the value ZERO to ignore the Resize preference in this library", "Standard Width Definition")
    free = FreeFile
    Open Arquivo For Output As free
    Print #free, x
    Close free

End Sub

Private Sub Set_library_Click(Index As Integer)

    Definir_valor_Padrão (Index)

End Sub

Private Sub Standard_Click()

    Standard.Checked = -Standard.Checked

End Sub

Private Sub Prepare_pastas()

    My_path = App.Path + "\Library\"
    Arquivo = Dir$(My_path & "*.*", vbDirectory)
    Do While Arquivo <> ""
        If Arquivo <> "." And Arquivo <> ".." Then
            If (GetAttr(My_path & Arquivo) And vbDirectory) = vbDirectory Then
                Pastas(Qt_pastas) = Arquivo
                Qt_pastas = Qt_pastas + 1
            End If
        End If
        Arquivo = Dir
        DoEvents
    Loop
    Qt_pastas = Qt_pastas - 1

End Sub


