VERSION 5.00
Begin VB.Form Splash_Form 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1470
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   2775
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMoveNext 
      Enabled         =   0   'False
      Left            =   0
      Top             =   3240
   End
   Begin VB.PictureBox picView 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   0
      Width           =   675
   End
End
Attribute VB_Name = "Splash_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sldFrames_Value As Long
Private sldFrames_Max As Long

Private Const HWND_TOPMOST As Integer = -1
Private Const HWND_NOTOPMOST As Integer = -2
Private Const SWP_NOMOVE As Integer = &H2
Private Const SWP_NOSIZE As Integer = &H1

Private Const LWA_COLORKEY As Integer = &H1
Private Const LWA_ALPHA As Integer = &H2
Private Const GWL_EXSTYLE As Integer = (-20)
Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function apiSetWindowPos Lib "User32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private GIF As cGifReader

Private Const MODULE_NAME As String = "frmAnimation"

Private Declare Function UpdateWindow Lib "User32" (ByVal hwnd As Long) As Long

Private m_oRenderer             As cBmpRenderer
Private WithEvents m_oReader    As cGifReader
Attribute m_oReader.VB_VarHelpID = -1
Private m_lFrameCount           As Long
Private m_aFrames()             As UcsFrameInfo

Private Type UcsFrameInfo
    oPic        As StdPicture
    nDelay      As Long
End Type

Public Function Init(oRdr As cGifReader) As Boolean

  Const FUNC_NAME     As String = "Init"
    
    On Error GoTo EH
    
    
    Set m_oReader = oRdr
    If m_oRenderer.Init(oRdr) Then
        Set picView.Picture = Nothing
        If oRdr.MoveLast() Then
            m_lFrameCount = oRdr.FrameIndex + 1
            If m_lFrameCount > 1 Then
                sldFrames_Max = m_lFrameCount
            End If
        End If
    End If

Exit Function

EH:
    Resume Next

End Function

Private Sub Form_Activate()

  Const FUNC_NAME     As String = "Form_Activate"
  Dim lIdx            As Long
  Dim sInfo           As String
 
    On Error GoTo EH
    If UBound(m_aFrames) < 0 And m_lFrameCount > 0 Then
        ReDim m_aFrames(1 To m_lFrameCount)
        If m_oRenderer.MoveFirst() Then
            lIdx = 0
            Do While True
                If Not m_oRenderer.MoveNext Then
                    Exit Do
                End If
                lIdx = lIdx + 1
                With m_aFrames(lIdx)
                    Set .oPic = m_oRenderer.Image
                    .nDelay = m_oRenderer.Reader.DelayTime
                    sldFrames_Value = lIdx
                    If lIdx = 1 Then
                        sldFrames_Change 'MOSTRA SÓ O PRIMEIRO FRAME
                    End If
                    DoEvents
                End With
            Loop
        End If
    End If
    sldFrames_Value = 0 'ENTRE AQUI QUAL O FRAME INICIAL
    tmrMoveNext_Timer
    tmrMoveNext.Interval = 1
    tmrMoveNext.Enabled = True
    
Exit Sub

EH:
    Resume Next

End Sub

Private Sub Form_Initialize()

    Set m_oRenderer = New cBmpRenderer
    ReDim m_aFrames(-1 To -1)
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Terminar
    End If

End Sub

Private Sub Form_Load()

  Dim Ret As Long
  Dim Arquivo As String
  Dim filenumber As Integer

  Dim sFilename   As String * 260
  Dim lRetval     As Long
  Dim Path_wallpaper As String

    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    
    SetLayeredWindowAttributes Me.hwnd, 10, 255, LWA_COLORKEY Or LWA_ALPHA
    
    Set GIF = New cGifReader
    
    Arquivo = App.Path & "\Doctor Frank1.gif"
    
    If GIF.Init(Arquivo) Then
        
        Width = GIF.ScreenWidth * Screen.TwipsPerPixelX
        Height = GIF.ScreenHeight * Screen.TwipsPerPixelY
           
        GIF.MoveFirst
    End If
    ReDim m_aFrames(-1 To -1)
    Init GIF
      
    Width = GIF.ScreenWidth * Screen.TwipsPerPixelX
    Height = GIF.ScreenHeight * Screen.TwipsPerPixelY
    
    picView.Width = Width
    picView.Height = Height
    picView.BackColor = GIF.BackgroundColor
    
End Sub

Private Sub Form_Resize()

    Width = GIF.ScreenWidth * Screen.TwipsPerPixelX
    Height = GIF.ScreenHeight * Screen.TwipsPerPixelY

End Sub

Private Sub m_oReader_Progress(ByVal CurrentLine As Long)

    UpdateWindow Me.hwnd

End Sub

Private Sub Open_gif_Click()

  Const FUNC_NAME     As String = "cmdOpen_Click"
  Dim filenumber As Integer
  Dim filename As String
  Static Ultimo_dir As String
  
    On Error GoTo EHCancel
    
    filename = OpenDialog(Me, "Compuserve|*.gif", "Doctor Frankenstin Open", Ultimo_dir)
    If filename = "" Then
        Exit Sub
    End If
    Ultimo_dir = CurDir
    On Error GoTo EH
    If GIF.Init(filename) Then
        Width = GIF.ScreenWidth * Screen.TwipsPerPixelX
        Height = GIF.ScreenHeight * Screen.TwipsPerPixelY
        GIF.MoveFirst
    End If
    
    ReDim m_aFrames(-1 To -1)
    
    Init GIF
    Form_Activate
    
    Width = GIF.ScreenWidth * Screen.TwipsPerPixelX
    Height = GIF.ScreenHeight * Screen.TwipsPerPixelY
    picView.Width = Width
    picView.Height = Height
            
    picView.BackColor = GIF.BackgroundColor
        
EHCancel:

Exit Sub

EH:
    
    Resume Next

End Sub

Private Sub sldFrames_Change()

  Const FUNC_NAME     As String = "sldFrames_Change"
  Dim lDelay          As Long
    
    On Error GoTo EH
    With m_aFrames(sldFrames_Value)
        lDelay = IIf(.nDelay < 8, 80, .nDelay * 10)
        Set picView.Picture = .oPic
        If tmrMoveNext.Enabled Then
            tmrMoveNext.Interval = lDelay
            tmrMoveNext.Enabled = False
            tmrMoveNext.Enabled = True
        End If
    End With
        
Exit Sub

EH:
    Resume Next

End Sub

Private Sub sldFrames_Scroll()
    
    sldFrames_Change

End Sub

Private Sub picView_Click()

    Terminar

End Sub

Private Sub tmrMoveNext_Timer()

  Static passo As Integer
  Static vezes As Integer

    If passo = 0 Then
        passo = 1
    End If
    
    If (sldFrames_Value + passo) < 1 Or (sldFrames_Value + passo) > sldFrames_Max Then
        vezes = vezes + 1
        If vezes < 100 Then
            Exit Sub
        End If
        Terminar
        Exit Sub
        
        passo = -passo
    End If
    sldFrames_Value = sldFrames_Value + passo
    sldFrames_Change
    DoEvents

End Sub

Private Sub Terminar()

    tmrMoveNext.Enabled = False
    Me.Hide

End Sub

