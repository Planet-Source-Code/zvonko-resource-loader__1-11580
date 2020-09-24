VERSION 5.00
Begin VB.Form frmExample 
   Caption         =   "Loading resources example"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   Icon            =   "frmExample.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "String"
      Height          =   750
      Left            =   3870
      TabIndex        =   10
      Top             =   15
      Width           =   555
   End
   Begin VB.TextBox txt 
      Height          =   3045
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   855
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.PictureBox pic 
      AutoSize        =   -1  'True
      Height          =   3555
      Left            =   45
      ScaleHeight     =   3495
      ScaleWidth      =   4305
      TabIndex        =   8
      Top             =   810
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.CommandButton Command8 
      Caption         =   "GIF"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   7
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command7 
      Caption         =   "JPEG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      TabIndex        =   6
      Top             =   15
      Width           =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "HTML"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   990
      TabIndex        =   5
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2910
      TabIndex        =   4
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command4 
      Caption         =   "WAVE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   3
      Top             =   390
      Width           =   960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cursor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1950
      TabIndex        =   2
      Top             =   15
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Icon"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   990
      TabIndex        =   1
      Top             =   15
      Width           =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bitmap"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   960
   End
End
Attribute VB_Name = "frmExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public res As LoadRes

Private Sub Command1_Click()
    pic.Move 45, 810, 4365, 3555
    txt.Visible = False
    pic.Visible = True
    pic.Picture = res.LoadPictureFromDLL(103, resBitmap)
End Sub

Private Sub Command2_Click()
    pic.Move 45, 810, 4365, 3555
    txt.Visible = False
    pic.Visible = True
    pic.Picture = res.LoadPictureFromDLL(126, resIcon)
End Sub

Private Sub Command3_Click()
    pic.Move 45, 810, 4365, 3555
    txt.Visible = False
    pic.Visible = True
    pic.Picture = res.LoadPictureFromDLL(132, resCursor)
End Sub

Private Sub Command4_Click()
    res.PlayWaveFromDLL 137
End Sub

Private Sub Command5_Click()
    Set res = Nothing
    Unload Me
End Sub

Private Sub Command6_Click()
    txt.Move 45, 810, Me.Width - 45, Me.Height - txt.Top - 45
    txt.Visible = True
    pic.Visible = False
    ResType = "html"
    txt.Text = res.LoadHtmlFromDLL(136)
End Sub

Private Sub Command7_Click()
    pic.Move 45, 810, 4365, 3555
    txt.Visible = False
    pic.Visible = True
    pic.Picture = res.LoadPictureFromDLL(127, resJPEG)
End Sub

Private Sub Command8_Click()
    pic.Move 45, 810, 4365, 3555
    txt.Visible = False
    pic.Visible = True
    pic.Picture = res.LoadPictureFromDLL(129, resGIF)
End Sub

Private Sub Command9_Click()
    Randomize
    MsgBox res.LoadStringFromDLL(100)
    MsgBox res.LoadStringFromDLL(101)
    MsgBox res.LoadStringFromDLL(102)
End Sub

Private Sub Form_Load()
    Set res = New LoadRes
    res.DllName = App.Path & "\Resources.dll"
End Sub

Private Sub Form_Resize()
    txt.Move 45, 810, Me.Width - 200, Me.Height - txt.Top - 420
End Sub
