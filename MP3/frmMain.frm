VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   ClientHeight    =   2160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   2160
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox cmdExit 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   3360
      Picture         =   "frmMain.frx":4FD50
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   6
      Top             =   0
      Width           =   135
   End
   Begin MSComctlLib.Slider SliderVol 
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "Volume"
      Top             =   960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   9
      Max             =   2500
      SelStart        =   2500
      TickStyle       =   3
      Value           =   2500
   End
   Begin MSComctlLib.Slider SliderPos 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   9
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider SliderBal 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      ToolTipText     =   "Balance"
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   9
      Min             =   -5000
      Max             =   5000
      TickStyle       =   3
   End
   Begin VB.Timer tmrPos 
      Interval        =   100
      Left            =   3720
      Top             =   0
   End
   Begin VB.Label cmdPlayList 
      BackStyle       =   0  'Transparent
      Caption         =   "PL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1870
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3120
      Picture         =   "frmMain.frx":4FE72
      Top             =   1800
      Width           =   405
   End
   Begin VB.Image ImgOrin 
      Height          =   105
      Left            =   720
      Picture         =   "frmMain.frx":50544
      Top             =   1920
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Image imgMin 
      Height          =   105
      Left            =   3120
      Picture         =   "frmMain.frx":50666
      Top             =   0
      Width           =   150
   End
   Begin VB.Image imgMouse 
      Height          =   105
      Left            =   480
      Picture         =   "frmMain.frx":50788
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgPlay 
      Height          =   375
      Left            =   720
      Top             =   960
      Width           =   255
   End
   Begin VB.Image imgPause 
      Height          =   375
      Left            =   960
      Top             =   960
      Width           =   375
   End
   Begin VB.Image imgFF 
      Height          =   375
      Left            =   1680
      Top             =   960
      Width           =   375
   End
   Begin VB.Image imgBack 
      Height          =   375
      Left            =   360
      Top             =   960
      Width           =   255
   End
   Begin VB.Image imgStop 
      Height          =   375
      Left            =   1320
      Top             =   960
      Width           =   375
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "***Gold MP3 by Henk le Roux ***"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Gold MP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPlaylist_Click()
    If frmPlaylist.Visible = True Then
        frmPlaylist.Visible = False
    Else
        frmPlaylist.Visible = True
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "GoldMP3"
    Me.Show
    DoEvents
    frmPlaylist.Show
    Load frmLoad
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMin.Picture = ImgOrin.Picture
    frmPlaylist.Left = Me.Left
    frmPlaylist.Top = Me.Top + Me.Height
    frmLoad.Left = Me.Left + Me.Width
    frmLoad.Top = Me.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmPlaylist
    Unload frmLoad
End Sub


Private Sub imgBack_Click()
    If frmPlaylist.lstPL.ListCount = 0 Then
        Exit Sub
    Else
        If frmPlaylist.lstPL.ListIndex - 1 > -1 Then
            frmPlaylist.lstPL.ListIndex = frmPlaylist.lstPL.ListIndex - 1
            MediaPlayer1.Filename = frmPlaylist.lstPL.Text
        Else
            frmPlaylist.lstPL.ListIndex = frmPlaylist.lstPL.ListCount - 1
            MediaPlayer1.Filename = frmPlaylist.lstPL.Text
        End If
    End If
End Sub


Private Sub imgFF_Click()
    If frmPlaylist.lstPL.ListCount = 0 Then
        Exit Sub
    Else
        If frmPlaylist.lstPL.ListIndex + 1 < frmPlaylist.lstPL.ListCount Then
            frmPlaylist.lstPL.ListIndex = frmPlaylist.lstPL.ListIndex + 1
            MediaPlayer1.Filename = frmPlaylist.lstPL.Text
        Else
            frmPlaylist.lstPL.ListIndex = 0
            MediaPlayer1.Filename = frmPlaylist.lstPL.Text
        End If
    End If
End Sub


Private Sub imgMin_Click()
    frmPlaylist.Hide
    frmLoad.Hide
    Me.WindowState = vbMinimized
End Sub


Private Sub imgMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgMin.Picture = imgMouse.Picture
End Sub


Private Sub imgPause_Click()
    If MediaPlayer1.CurrentPosition <> -1 Then MediaPlayer1.Pause
End Sub


Private Sub imgPlay_Click()
    If MediaPlayer1.CurrentPosition <> -1 Then MediaPlayer1.Play
End Sub


Private Sub imgStop_Click()
    If MediaPlayer1.CurrentPosition <> -1 Then
        MediaPlayer1.Stop
        MediaPlayer1.CurrentPosition = 0
    End If
End Sub


Private Sub SliderBal_Click()
On Error Resume Next
    MediaPlayer1.Balance = SliderBal.Value
End Sub


Private Sub SliderPos_Click()
On Error Resume Next
    MediaPlayer1.CurrentPosition = SliderPos.Value
End Sub


Private Sub SliderVol_Click()
Dim plVol As Long
On Error Resume Next
    plVol = SliderVol.Value - 2500
    MediaPlayer1.Volume = plVol
End Sub


Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
Dim msValue As String
    Randomize Timer
    If frmPlaylist.lstPL.ListIndex + 1 = frmPlaylist.lstPL.ListCount Then
        msValue = 0
    Else
        msValue = frmPlaylist.lstPL.ListIndex + 1
    End If
    
    frmPlaylist.lstPL.ListIndex = msValue
    MediaPlayer1.Filename = frmPlaylist.lstPL.Text
End Sub


Private Sub tmrPos_Timer()
    SliderPos.Value = MediaPlayer1.CurrentPosition
    tinseconden = MediaPlayer1.CurrentPosition
    Dim min As Integer
    Dim sec As Integer
    min = tinseconden \ 60
    sec = tinseconden - (min * 60)
    If sec = "-1" Then sec = "0"
    lblStatus.Caption = "*** MP3 Player by Henk le Roux " & min & ":" & sec & " ***"
End Sub


