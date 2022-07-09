VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dada Skin Changer"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox textBasedir 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2520
      Width           =   3495
   End
   Begin VB.Timer timerUpdate 
      Interval        =   1000
      Left            =   4200
      Top             =   120
   End
   Begin VB.CommandButton btnSetToDefault 
      Caption         =   "&Set to Default Skin"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label labelLastUpdate 
      Caption         =   "labelLastUpdate"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1920
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Last skin update:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   780
      Left            =   4680
      Picture         =   "frmMain.frx":0EDE
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label6 
      Caption         =   "Last file check:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Listen status:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label textInFile 
      Caption         =   "textInFile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label labelLastCheck 
      Caption         =   "labelLastCheck"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1560
      Width           =   3495
   End
   Begin VB.Label labelStatus 
      Caption         =   "OK"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Data base directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Skin listed in file:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":2A22
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oWinamp As New WINAMPCOMLib.Application

Dim CurrentSkin As String
Dim CurrentBasedir As String
Dim DefaultSkin As String

Private Sub btnSetToDefault_Click()
    Call SetWinampSkin(DefaultSkin)
    Call SetSkinToFile(DefaultSkin)
End Sub

Private Sub btnAbout_Click()
    frmAbout.Show
End Sub

Private Sub Form_Load()
    DefaultSkin = "[base-2.91.wsz"
    CurrentBasedir = "M:\Stuff\Data\winampxp\"
    Me.WindowState = 1
    Call FullUpdate
End Sub

Private Function SetToFile(strValue As String, strFile As String, strType As String)
    FileName = CurrentBasedir & strFile & ".txt"
    On Error GoTo ErrorHandler
    Open FileName For Output As #1
        Print #1, strValue
    Close #1
    Exit Function
ErrorHandler:
    labelStatus = "Error (Can't find or open " & strType & " file)"
End Function

Private Function SetPlayingToFile()
    Call SetToFile(oWinamp.Status, "is_playing", "playing")
End Function

Private Function SetSongToFile()
    Call SetToFile(oWinamp.CurrentSongFileName, "song", "song")
End Function

Private Function SetTitleToFile()
    Call SetToFile(oWinamp.CurrentSongTitle, "title", "title")
End Function

Private Function SetSkinToFile(strSkin As String)
    Call SetToFile(strSkin, "skin", "skin")
End Function

Private Function GetSkinFromFile()
    FileName = CurrentBasedir & "skin.txt"
    labelStatus = "OK"
    labelLastCheck = Now
    On Error GoTo ErrorHandler
    Dim FileLine As String
    Open FileName For Input As #1
        Line Input #1, FileLine
    Close #1
    If FileLine <> CurrentSkin Then
        CurrentSkin = FileLine
        Call SetWinampSkin(CurrentSkin)
    End If
    textInFile = FileLine
    textBasedir = CurrentBasedir
    Exit Function
ErrorHandler:
    labelStatus = "Error (Can't find or open skin file)"
End Function

Private Function SetWinampSkin(strSkin As String)
    oWinamp.SkinName = strSkin
    labelLastUpdate = Now
End Function

Private Function FullUpdate()
    Call GetSkinFromFile
    Call SetSongToFile
    Call SetTitleToFile
    Call SetPlayingToFile
End Function

Private Sub textBasedir_Change()
    CurrentBasedir = textBasedir
    Call FullUpdate
End Sub

Public Sub timerUpdate_Timer()
    Call FullUpdate
End Sub

