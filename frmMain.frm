VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dada Skin Changer"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox textFilename 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2400
      Width           =   3495
   End
   Begin VB.Timer timerUpdate 
      Interval        =   1000
      Left            =   3360
      Top             =   360
   End
   Begin VB.CommandButton btnSetToDefault 
      Caption         =   "&Set to Default Skin"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Label labelLastUpdate 
      Caption         =   "labelLastUpdate"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label7 
      Caption         =   "Last skin update:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1800
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
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Listen status:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1080
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
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label labelLastCheck 
      Caption         =   "labelLastCheck"
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
   End
   Begin VB.Label labelStatus 
      Caption         =   "OK"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label4 
      Caption         =   "Filename we're listening to:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Skin listed in file:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "This program will change the Winamp skin based on the contents of a text file."
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private oWinamp As New WINAMPCOMLib.Application

Dim CurrentSkin As String
Dim CurrentFilename As String
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
    CurrentFilename = "M:\Stuff\Data\winampxp_skin.txt"
    Me.WindowState = 1
    Call GetSkinFromFile
End Sub

Private Function SetSkinToFile(strSkin As String)
    On Error GoTo ErrorHandler
    Open CurrentFilename For Output As #1
        Print #1, strSkin
    Close #1
    Call GetSkinFromFile
    Exit Function
ErrorHandler:
    labelStatus = "Error (Can't find or open file)"
End Function

Private Function GetSkinFromFile()
    labelStatus = "OK"
    labelLastCheck = Now
    On Error GoTo ErrorHandler
    Dim FileLine As String
    Open CurrentFilename For Input As #1
        Line Input #1, FileLine
    Close #1
    If FileLine <> CurrentSkin Then
        CurrentSkin = FileLine
        Call SetWinampSkin(CurrentSkin)
    End If
    textInFile = FileLine
    textFilename = CurrentFilename
    Exit Function
ErrorHandler:
    labelStatus = "Error (Can't find or open file)"
End Function

Private Function SetWinampSkin(strSkin As String)
    oWinamp.SkinName = strSkin
    labelLastUpdate = Now
End Function

Private Sub textFilename_Change()
    CurrentFilename = textFilename
    Call GetSkinFromFile
End Sub

Public Sub timerUpdate_Timer()
    Call GetSkinFromFile
End Sub
