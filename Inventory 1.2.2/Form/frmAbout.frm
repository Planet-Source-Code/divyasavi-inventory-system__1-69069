VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   "Contact"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12225
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      X1              =   2280
      X2              =   2280
      Y1              =   5400
      Y2              =   7440
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFC0C0&
      X1              =   2400
      X2              =   2400
      Y1              =   5280
      Y2              =   7320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   2760
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      X1              =   120
      X2              =   2880
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+91 9825958895"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   2685
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "haresh_valani@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   4560
      Width           =   2685
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Divya Soft Tech"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   2685
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Version :-   1.0.1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   2685
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©   2007-08"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   2685
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DressCode Inventory System"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   2685
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Application Licensed to   :"
      Height          =   345
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JALARAM PETROLEUM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   450
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   2925
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":0000
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   2760
      TabIndex        =   1
      Top             =   7200
      Width           =   9375
   End
   Begin VB.Image Image4 
      Height          =   5955
      Left            =   3000
      Picture         =   "frmAbout.frx":010D
      Top             =   1080
      Width           =   9030
   End
   Begin VB.Image Image3 
      Height          =   1350
      Left            =   240
      Picture         =   "frmAbout.frx":121D1
      Top             =   6600
      Width           =   1500
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   435
      Left            =   11760
      TabIndex        =   0
      Top             =   80
      Width           =   225
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   0
      Picture         =   "frmAbout.frx":12638
      Stretch         =   -1  'True
      Top             =   7630
      Width           =   12225
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Picture         =   "frmAbout.frx":6EA7A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12225
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Label1.ForeColor = &H80FF&
    Label1.FontBold = False
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Label1.ForeColor = &H80FF&
    Label1.FontBold = False
End Sub

Private Sub Label1_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Label1.ForeColor = &HFF&
    Label1.FontBold = True
End Sub
