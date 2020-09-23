VERSION 5.00
Begin VB.Form frmPlaylist 
   BackColor       =   &H8000000C&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   Picture         =   "frmPlaylist.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstPL 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
   Begin VB.Image cmdRemove 
      Height          =   375
      Left            =   1560
      Top             =   3840
      Width           =   375
   End
   Begin VB.Image cmdAdd 
      Height          =   375
      Left            =   1080
      Top             =   3840
      Width           =   375
   End
End
Attribute VB_Name = "frmPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    frmLoad.Show
End Sub


Private Sub cmdRemove_Click()
    If lstPL.ListIndex <> -1 Then lstPL.RemoveItem (lstPL.ListIndex)
End Sub


Private Sub Form_Load()
    Me.Left = Me.Left
    Me.Top = Me.Top + Me.Height
    Me.Width = frmMain.Width
    gbPlaylist = True
    DoEvents
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
    gbPlaylist = False
End Sub


Private Sub lstPL_DblClick()
   frmMain.MediaPlayer1.Filename = lstPL.Text
    
    If lstPL.Text <> "" Then
        frmMain.MediaPlayer1.Play
        frmMain.SliderPos.Max = frmMain.MediaPlayer1.Duration
        'CmdPause.Enabled = True
    Else
        MsgBox "Error reading file", vbCritical & vbOKOnly, "Error"
    End If
End Sub
