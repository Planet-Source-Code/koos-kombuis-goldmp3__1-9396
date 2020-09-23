VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H80000010&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLoad.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0080FFFF&
      Height          =   5040
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000007&
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label lblNo 
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   230
      Left            =   3200
      TabIndex        =   5
      Top             =   6070
      Width           =   255
   End
   Begin VB.Label lblOK 
      BackStyle       =   0  'Transparent
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2720
      TabIndex        =   4
      Top             =   6070
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2640
      Picture         =   "frmLoad.frx":4F90E
      Top             =   6000
      Width           =   840
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Load File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub


Private Sub Form_Load()
    Me.Left = frmMain.Left + frmMain.Width
    Me.Top = frmMain.Top
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub


Private Sub lblNo_Click()
    Me.Hide
End Sub


Private Sub lblOK_Click()
Dim miCount As Integer

File1.Path = Dir1.Path

    If File1.ListCount <> 0 Then
        For miCount = 1 To File1.ListCount
            File1.ListIndex = miCount - 1
            
            If Len(Dir1.Path) > 3 Then
                frmPlaylist.lstPL.AddItem Dir1.Path & "\" & File1.Filename
            Else
                frmPlaylist.lstPL.AddItem Dir1.Path & File1.Filename
            End If
        Next miCount
                Unload Me
    Else
        MsgBox "No files were found in specific folder", vbOKOnly, "Error"
    End If
End Sub
