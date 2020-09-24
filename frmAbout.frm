VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About ResDllLoad"
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   608
      TabIndex        =   6
      Top             =   2100
      Width           =   1995
   End
   Begin VB.Label Label6 
      Caption         =   "Attention: This program is freeware. You can freely distribute it but without any profit. Comments are welcome!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   5
      Top             =   1365
      Width           =   3135
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      X1              =   0
      X2              =   3220
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   3220
      X2              =   0
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Label Label5 
      Caption         =   "zvonko.bostjancic@siol.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   915
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   945
      Width           =   1965
   End
   Begin VB.Label Label4 
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   2565
      Picture         =   "frmAbout.frx":0152
      Stretch         =   -1  'True
      Top             =   30
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   45
      Picture         =   "frmAbout.frx":0608
      Stretch         =   -1  'True
      Top             =   45
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2000 by Zvonko Boštjanèiè"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   285
      TabIndex        =   2
      Top             =   705
      Width           =   2670
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      TabIndex        =   1
      Top             =   405
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ResDllLoad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   405
      Left            =   570
      TabIndex        =   0
      Top             =   45
      Width           =   1965
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Label2.Caption = "Version " & App.Major & "." & App.Minor
End Sub

Sub OdpriURL(frm As Form, ByVal Komanda As String)
    Dim x
    x = ShellExecute(frm.hwnd, "open", Komanda, "", "", SW_SHOW)
End Sub

Private Sub Label5_Click()
    OdpriURL Me, "mailto:zvonko.bostjancic@siol.net"
End Sub
