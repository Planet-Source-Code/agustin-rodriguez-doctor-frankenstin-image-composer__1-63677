VERSION 5.00
Begin VB.Form frmFilters 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   ClientHeight    =   4380
   ClientLeft      =   4425
   ClientTop       =   1560
   ClientWidth     =   5850
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   292
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   5055
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   0
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   60
      Top             =   60
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   4230
      Left            =   30
      ScaleHeight     =   282
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1.50000e5
      Width           =   5055
   End
   Begin VB.Menu Menu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu File 
         Caption         =   "&File"
         Begin VB.Menu mnuOpenFile 
            Caption         =   "Load"
         End
         Begin VB.Menu Save 
            Caption         =   "Save"
         End
         Begin VB.Menu mnuSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuExit 
            Caption         =   "Exit"
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit"
         Begin VB.Menu mnuUndo 
            Caption         =   "&Fix"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuRedo 
            Caption         =   "Restore"
            Shortcut        =   ^Z
         End
         Begin VB.Menu mnuSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopy 
            Caption         =   "Copy"
            Shortcut        =   ^C
         End
      End
      Begin VB.Menu mnuRetouch 
         Caption         =   "&Filters"
         Begin VB.Menu mnuDefinition 
            Caption         =   "Definition"
            Begin VB.Menu mnuBlur 
               Caption         =   "Smooth"
            End
            Begin VB.Menu mnuBlurMore 
               Caption         =   "Blur"
            End
            Begin VB.Menu mnuSharpen 
               Caption         =   "Sharpen"
            End
            Begin VB.Menu mnuSharpenMore 
               Caption         =   "Sharpen More"
            End
            Begin VB.Menu mnuDiffuse 
               Caption         =   "Diffuse"
            End
            Begin VB.Menu mnuDiffuseMore 
               Caption         =   "Diffuse More"
            End
            Begin VB.Menu mnuPixelize 
               Caption         =   "Pixelize"
            End
         End
         Begin VB.Menu mnuEdges 
            Caption         =   "Edges"
            Begin VB.Menu mnuEmboss 
               Caption         =   "Emboss"
            End
            Begin VB.Menu mnuEmbossMore 
               Caption         =   "Emboss More"
            End
            Begin VB.Menu mnuEngrave 
               Caption         =   "Engrave"
            End
            Begin VB.Menu mnuEngraveMore 
               Caption         =   "Engrave More"
            End
            Begin VB.Menu mnuRelief 
               Caption         =   "Relief"
            End
            Begin VB.Menu mnuEdge 
               Caption         =   "Edge Enhance"
            End
            Begin VB.Menu mnuContour 
               Caption         =   "Contour"
            End
            Begin VB.Menu mnuConnection 
               Caption         =   "Connected Contour"
            End
         End
         Begin VB.Menu mnuColors 
            Caption         =   "Colors"
            Begin VB.Menu mnuGreyScale 
               Caption         =   "GreyScale"
            End
            Begin VB.Menu mnuBlackWhite 
               Caption         =   "Black && White"
               Begin VB.Menu mnuBW1 
                  Caption         =   "Nearest Color"
               End
               Begin VB.Menu mnuBW2 
                  Caption         =   "Enhanced Diffusion"
               End
               Begin VB.Menu mnuBW3 
                  Caption         =   "Ordered Dither"
               End
               Begin VB.Menu mnuBW4 
                  Caption         =   "Floyd-Steinberg"
               End
               Begin VB.Menu mnuBW5 
                  Caption         =   "Burke"
               End
               Begin VB.Menu mnuBW6 
                  Caption         =   "Stucki"
               End
            End
            Begin VB.Menu mnuNegative 
               Caption         =   "Negative"
            End
            Begin VB.Menu mnuSwapColors 
               Caption         =   "Swap Colors"
               Begin VB.Menu mnuSwapBank 
                  Caption         =   "RGB -> BRG"
                  Index           =   1
               End
               Begin VB.Menu mnuSwapBank 
                  Caption         =   "RGB -> GBR"
                  Index           =   2
               End
               Begin VB.Menu mnuSwapBank 
                  Caption         =   "RGB -> RBG"
                  Index           =   3
               End
               Begin VB.Menu mnuSwapBank 
                  Caption         =   "RGB -> BGR"
                  Index           =   4
               End
               Begin VB.Menu mnuSwapBank 
                  Caption         =   "RGB -> GRB"
                  Index           =   5
               End
            End
            Begin VB.Menu mnuAqua 
               Caption         =   "Aqua"
            End
            Begin VB.Menu mnuAddNoise 
               Caption         =   "Add Noise"
            End
            Begin VB.Menu mnuGamma 
               Caption         =   "Gamma Correction"
            End
         End
         Begin VB.Menu mnuIntensity 
            Caption         =   "Intensity"
            Begin VB.Menu mnuBrighter 
               Caption         =   "Brighter"
               Shortcut        =   {F3}
            End
            Begin VB.Menu mnuDarker 
               Caption         =   "Darker"
               Shortcut        =   {F4}
            End
            Begin VB.Menu mnuContrast1 
               Caption         =   "Increase Contrast"
               Shortcut        =   {F5}
            End
            Begin VB.Menu mnuContrast2 
               Caption         =   "Decrease Contrast"
               Shortcut        =   {F6}
            End
            Begin VB.Menu mnuDilate 
               Caption         =   "Dilate"
            End
            Begin VB.Menu mnuErode 
               Caption         =   "Erode"
            End
            Begin VB.Menu mnuStretch 
               Caption         =   "Contrast Stretch"
            End
            Begin VB.Menu mnuSaturationI 
               Caption         =   "Increase Saturation"
               Shortcut        =   {F7}
            End
            Begin VB.Menu mnuSaturationD 
               Caption         =   "Decrease Saturation"
               Shortcut        =   {F8}
            End
         End
      End
      Begin VB.Menu Close_retouch 
         Caption         =   "Close Retouch Window"
      End
   End
End
Attribute VB_Name = "frmFilters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private capture As Integer
Private xxxx As Single
Private yyyy As Single

Private Não_processe As Integer
Private mProgress As Long

Private Sub Close_retouch_Click()

    Unload Me

End Sub

Private Sub Form_Activate()

    PopupMenu Menu, , ScaleWidth, 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
      Case Asc("Z")
        If Shift = 2 Then
            mnuRedo_Click
        End If
      Case Asc("C")
        If Shift = 2 Then
            mnuCopy_Click
        End If
      Case 114 'F3
        mnuBrighter_Click
      Case 115 'F4
        mnuDarker_Click
      Case 116 'F5
        mnuContrast1_Click
      Case 117 'F6
        mnuContrast2_Click
      Case 118
        mnuSaturationI_Click
      Case 119
        mnuSaturationD_Click
    End Select
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = 1 Then
        xxxx = x * Screen.TwipsPerPixelX: yyyy = Y * Screen.TwipsPerPixelY
        capture = True
        ReleaseCapture
        SetCapture Me.hwnd
        Exit Sub
    End If
    
    If Button = 2 Then
        PopupMenu Menu
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call mnuExit_Click

End Sub

Private Sub Form_Resize()

  Dim retval As Long

    If Não_processe Then
        Exit Sub
    End If

    Call SetStretchBltMode(hDC, STRETCHMODE)
    retval = StretchBlt(hDC, 0, 0, ScaleWidth, ScaleHeight, Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, vbSrcCopy)

    Picture1.Width = ScaleWidth
    Picture1.Height = ScaleHeight

    Call SetStretchBltMode(Picture1.hDC, STRETCHMODE)
    retval = StretchBlt(Picture1.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)

    Refresh

End Sub

Private Sub mnuCopy_Click()

  Dim retval As Long

    Main_form.Trabalho.Width = ScaleWidth
    Main_form.Trabalho.Height = ScaleHeight
    Main_form.Trabalho.Cls

    Call SetStretchBltMode(Main_form.Trabalho.hDC, STRETCHMODE)
    retval = StretchBlt(Main_form.Trabalho.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)
    
    Clipboard.Clear

    Clipboard.SetData Main_form.Trabalho.Image

End Sub

Private Sub mnuExit_Click()

    Unload frmFilters
 
End Sub

Private Sub mnuOpenFile_Click()
  
  Dim filename As String
  
    On Error GoTo e_trap
            
    filename = OpenDialog(Me, "All types|*.bmp;*.jpg;*.gif|Bitmap|*.bmp|Compuserv|*.gif|JPeg|*.jpg", "Doctor Frankenstin Open", "")
            
    If filename <> "" Then
            
        Picture1.Picture = LoadPicture(filename)
        Picture3.Picture = LoadPicture(filename)
        Prepare
  
    End If
    
sair:

Exit Sub

e_trap:
    
    Resume sair
  
Exit Sub

    If filename <> "" Then
        Picture1.Picture = LoadPicture(filename)
        Picture2.Width = Picture1.Width
        Picture2.Height = Picture1.Height
        Call BitBlt(Picture2.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
        frmFilters.Width = (Picture1.ScaleWidth + 8) * Screen.TwipsPerPixelX
        frmFilters.Height = (Picture1.ScaleHeight + 46) * Screen.TwipsPerPixelY
        Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
        Me.Refresh
    
    End If

End Sub

Private Sub mnuRedo_Click()
 
    frmFilters.Picture1.Width = Picture3.ScaleWidth
    frmFilters.Picture1.Height = Picture3.ScaleHeight
    Me.Refresh
    Call SetStretchBltMode(Picture1.hDC, STRETCHMODE)
    Call StretchBlt(Picture1.hDC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture3.hDC, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, vbSrcCopy)
    Prepare
    Me.Refresh

End Sub

Private Sub Timer1mnuUndo_Click()

    Call BitBlt(Me.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hDC, 0, 0, SRCCOPY)
    Call BitBlt(Picture2.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Call BitBlt(Picture1.hDC, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuUndo_Click()

    Picture1.Width = ScaleWidth
    Picture1.Height = ScaleHeight
    Picture2.Width = ScaleWidth
    Picture2.Height = ScaleHeight
    Picture3.Width = ScaleWidth
    Picture3.Height = ScaleHeight

    Call SetStretchBltMode(Picture1.hDC, STRETCHMODE)
    Call StretchBlt(Picture1.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)

    Call SetStretchBltMode(Picture3.hDC, STRETCHMODE)
    Call StretchBlt(Picture3.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy)

    Me.Refresh

End Sub

Private Sub Save_Click()
  
    Arquivo = SaveDialog(Me, "Graphic file|*.bmp", ".bmp", "Doctor Frankenstin Save", "")
    If Arquivo <> "" Then
        Picture1.Cls
        Picture1.Width = ScaleWidth
        Picture1.Height = ScaleHeight
        Call SetStretchBltMode(Picture1.hDC, STRETCHMODE)
        StretchBlt Picture1.hDC, 0, 0, ScaleWidth, ScaleHeight, hDC, 0, 0, ScaleWidth, ScaleHeight, vbSrcCopy
        Picture1.Refresh
        SavePicture Picture1.Image, Arquivo
    End If
    
End Sub

Private Sub Timer1_Timer()

    If mProgress = 100 Then
        Timer1.Enabled = False
        Prepare
    End If
 
End Sub

Private Sub mnuSharpenMore_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSHARPEN, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuSharpen_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSHARPEN, Picture1.Image, 2, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuEmboss_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iEMBOSS, Picture1.Image, RGB(0, 120, 120), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuEmbossMore_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iEMBOSSMORE, Picture1.Image, RGB(0, 120, 120), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBlur_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSMOOTH, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBlurMore_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iBLUR, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuDiffuseMore_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iDIFFUSE, Picture1.Image, 12, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuDiffuse_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iDIFFUSE, Picture1.Image, 6, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuEdge_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iEDGE, Picture1.Image, 1, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuEngrave_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iENGRAVE, Picture1.Image, RGB(0, 120, 120), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuEngraveMore_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iENGRAVEMORE, Picture1.Image, RGB(0, 120, 120), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuContour_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCONTOUR, Picture1.Image, RGB(255, 255, 255), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)

    Me.Refresh
  
End Sub

Private Sub mnuGreyScale_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iGREYSCALE, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuRelief_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iRELIEF, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBrighter_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iBRIGHTNESS, Picture1.Image, 10, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuDarker_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iBRIGHTNESS, Picture1.Image, -10, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuPixelize_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iPIXELIZE, Picture1.Image, 3, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuNegative_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iNEGATIVE, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuSwapBank_Click(Index As Integer)

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSWAPBANK, Picture1.Image, Index, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuContrast1_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCONTRAST, Picture1.Image, 10, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuContrast2_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCONTRAST, Picture1.Image, -10, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW1_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH1, Picture1.Image, RGB(180, 180, 180), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW2_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH2, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW3_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH3, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW4_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH4, Picture1.Image, 15, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW5_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH5, Picture1.Image, 15, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuBW6_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCOLDEPTH6, Picture1.Image, 15, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuAqua_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iAQUA, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuDilate_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iDILATE, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuErode_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iERODE, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuConnection_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iCONNECTION, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuStretch_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSTRETCH, Picture1.Image, 0, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuAddNoise_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iADDNOISE, Picture1.Image, 50, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuSaturationD_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSATURATION, Picture1.Image, -20, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuSaturationI_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iSATURATION, Picture1.Image, 15, mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Private Sub mnuGamma_Click()

    Timer1.Enabled = True
    Call BitBlt(Picture2.hDC, 0, 0, Me.ScaleWidth, Me.ScaleHeight, Me.hDC, 0, 0, SRCCOPY)
    Call FilterG(iGAMMA, Picture1.Image, InputBox("Gamma Factor * 100"), mProgress)
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh
  
End Sub

Public Sub Prepare()

    Picture2.Width = Picture1.Width
    Picture2.Height = Picture1.Height
    Call BitBlt(Picture2.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    frmFilters.Width = Picture1.ScaleWidth * Screen.TwipsPerPixelX + Width - ScaleWidth * Screen.TwipsPerPixelX
    frmFilters.Height = Picture1.ScaleHeight * Screen.TwipsPerPixelY + Height - ScaleHeight * Screen.TwipsPerPixelY
    Call BitBlt(Me.hDC, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture1.hDC, 0, 0, SRCCOPY)
    Me.Refresh

End Sub


