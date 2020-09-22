VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Inventory System"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   13935
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PicMenu 
      BorderStyle     =   0  'None
      FillColor       =   &H80000001&
      Height          =   8175
      Left            =   0
      ScaleHeight     =   8175
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
      Begin VB.PictureBox PicReport 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   0
         ScaleHeight     =   1575
         ScaleWidth      =   3255
         TabIndex        =   11
         Top             =   4800
         Width           =   3255
         Begin VB.Label lblEmpList 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Employee Details"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":4878
            MousePointer    =   99  'Custom
            TabIndex        =   16
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label lblRptEmp 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Emp Wise"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":4B82
            MousePointer    =   99  'Custom
            TabIndex        =   13
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label lblRptItems 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Items Wise"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":4E8C
            MousePointer    =   99  'Custom
            TabIndex        =   12
            Top             =   600
            Width           =   3255
         End
      End
      Begin VB.PictureBox PicTrans 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   0
         ScaleHeight     =   2175
         ScaleWidth      =   3255
         TabIndex        =   5
         Top             =   2040
         Width           =   3255
         Begin VB.Label lblDad 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Dad"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":5196
            MousePointer    =   99  'Custom
            TabIndex        =   9
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label lblReturn 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Return"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":54A0
            MousePointer    =   99  'Custom
            TabIndex        =   8
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label lblIssue 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Issue"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":57AA
            MousePointer    =   99  'Custom
            TabIndex        =   7
            Top             =   600
            Width           =   3255
         End
         Begin VB.Label lblReceive 
            Alignment       =   2  'Center
            BackColor       =   &H80000009&
            Caption         =   "Receive"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            MouseIcon       =   "frmMain.frx":5AB4
            MousePointer    =   99  'Custom
            TabIndex        =   6
            Top             =   120
            Width           =   3255
         End
      End
      Begin VB.Label lblExit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":5DBE
         MousePointer    =   99  'Custom
         TabIndex        =   15
         Top             =   6480
         Width           =   3300
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   1440
         TabIndex        =   14
         Top             =   7920
         Width           =   1065
      End
      Begin VB.Image Image6 
         Height          =   1020
         Left            =   1560
         Picture         =   "frmMain.frx":60C8
         Top             =   7200
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   840
         Picture         =   "frmMain.frx":6BC9
         Top             =   6480
         Width           =   1500
      End
      Begin VB.Label lblReport 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":7030
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   4320
         Width           =   3300
      End
      Begin VB.Label lblCStock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "Current Stock"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":733A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   120
         Width           =   3300
      End
      Begin VB.Label lblEmpMst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Employee Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":7644
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   600
         Width           =   3300
      End
      Begin VB.Label lblItems 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "New Items"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":794E
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1080
         Width           =   3300
      End
      Begin VB.Label lblTrans 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Transaction"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   0
         MouseIcon       =   "frmMain.frx":7C58
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1560
         Width           =   3300
      End
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   0
      Picture         =   "frmMain.frx":7F62
      Stretch         =   -1  'True
      Top             =   9960
      Width           =   15375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
On Error Resume Next
    PicMenu.BackColor = &H80000009
    PicMenu.Height = Image2.Top - PicMenu.Top
    lblCStock.Top = 120
    lblEmpMst.Top = lblCStock.Top + lblCStock.Height + 165
    lblItems.Top = lblEmpMst.Top + lblEmpMst.Height + 165
    lblTrans.Top = lblItems.Top + lblItems.Height + 165
    PicTrans.Visible = False
    lblReport.Top = lblTrans.Top + lblTrans.Height + 165
    PicReport.Visible = False
    lblExit.Top = lblReport.Top + lblReport.Height + 165
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Call MouseMove
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    End
End Sub

Private Sub Image6_Click()
On Error Resume Next
    frmAbout.Show , Me
    frmAbout.Left = PicMenu.Width
    frmAbout.Top = 2050
    frmAbout.Width = Me.Width - frmAbout.Left
    frmAbout.Height = Image2.Top - (frmAbout.Top / 1.15)
    frmAbout.Image2.Top = frmAbout.Height - frmAbout.Image2.Height
End Sub

Private Sub lblCStock_Click()
On Error Resume Next
    frmCStock.Show , Me
    frmCStock.Left = PicMenu.Width
    frmCStock.Top = 2050
    frmCStock.Width = Me.Width - frmCStock.Left
    frmCStock.Height = Image2.Top - frmCStock.Top
   
End Sub

Private Sub lblCStock_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblCStock.BackStyle = 1
    lblCStock.BackColor = &H80FF&
End Sub

Private Sub lblDad_Click()
On Error Resume Next
    frmDad.Show , Me
    frmDad.Left = PicMenu.Width
    frmDad.Top = 2050
    frmDad.Width = Me.Width - frmDad.Left
    frmDad.Height = Image2.Top - frmDad.Top
End Sub

Private Sub lblDad_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblDad.BackStyle = 1
    lblDad.BackColor = &H80FF&
End Sub

Private Sub lblEmpList_Click()
On Error Resume Next
    frmEmpList.Show , Me
    frmEmpList.Left = PicMenu.Width
    frmEmpList.Top = 2050
    frmEmpList.Width = Me.Width - frmEmpList.Left
    frmEmpList.Height = Image2.Top - frmEmpList.Top
End Sub

Private Sub lblEmpList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblEmpList.BackStyle = 1
    lblEmpList.BackColor = &H80FF&
End Sub

Private Sub lblEmpMst_Click()
On Error Resume Next
    frmEmpMst.Show , Me
    frmEmpMst.Left = PicMenu.Width
    frmEmpMst.Top = 2050
    frmEmpMst.Width = Me.Width - frmEmpMst.Left
    frmEmpMst.Height = Image2.Top - frmEmpMst.Top
End Sub

Private Sub lblEmpMst_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblEmpMst.BackStyle = 1
    lblEmpMst.BackColor = &H80FF&
End Sub

Private Sub lblExit_Click()
On Error Resume Next
    End
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblExit.BackStyle = 1
    lblExit.BackColor = &H80FF&
End Sub

Private Sub lblIssue_Click()
On Error Resume Next
    frmIssue.Show , Me
    frmIssue.Left = PicMenu.Width
    frmIssue.Top = 2050
    frmIssue.Width = Me.Width - frmIssue.Left
    frmIssue.Height = Image2.Top - frmIssue.Top
End Sub

Private Sub lblIssue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblIssue.BackStyle = 1
    lblIssue.BackColor = &H80FF&
End Sub

Private Sub lblItems_Click()
On Error Resume Next
    frmItems.Show , Me
    frmItems.Left = PicMenu.Width
    frmItems.Top = 2050
    frmItems.Width = Me.Width - frmItems.Left
    frmItems.Height = Image2.Top - frmItems.Top
End Sub

Private Sub lblItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblItems.BackStyle = 1
    lblItems.BackColor = &H80FF&
End Sub

Private Sub lblReceive_Click()
On Error Resume Next
    frmReceive.Show , Me
    frmReceive.Left = PicMenu.Width
    frmReceive.Top = 2050
    frmReceive.Width = Me.Width - frmReceive.Left
    frmReceive.Height = Image2.Top - frmReceive.Top
End Sub

Private Sub lblReceive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblReceive.BackStyle = 1
    lblReceive.BackColor = &H80FF&
End Sub

Private Sub lblReport_Click()
On Error Resume Next
    If PicTrans.Visible = True Then
        PicTrans.Visible = False
        lblReport.Top = lblTrans.Top + lblTrans.Height + 165
        PicReport.Top = lblReport.Top + lblReport.Height + 165
        lblExit.Top = lblReport.Top + lblReport.Height + 165
    End If
    
    If PicReport.Visible = False Then
        PicReport.Visible = True
        PicReport.Top = lblReport.Top + lblReport.Height + 165
        lblExit.Top = PicReport.Top + PicReport.Height + 165
    ElseIf PicReport.Visible = True Then
        PicReport.Visible = False
        lblExit.Top = lblReport.Top + lblReport.Height + 165
    End If
    
End Sub

Private Sub lblReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblReport.BackStyle = 1
    lblReport.BackColor = &H80FF&
End Sub

Private Sub lblReturn_Click()
On Error Resume Next
   frmReturn.Show , Me
   frmReturn.Left = PicMenu.Width
   frmReturn.Top = 2050
   frmReturn.Width = Me.Width - frmReturn.Left
   frmReturn.Height = Image2.Top - frmReturn.Top

End Sub

Private Sub lblReturn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblReturn.BackStyle = 1
    lblReturn.BackColor = &H80FF&
End Sub

Private Sub lblRptEmp_Click()
On Error Resume Next
    frmRptEW.Show , Me
    frmRptEW.Left = PicMenu.Width
    frmRptEW.Top = 2050
    frmRptEW.Width = Me.Width - frmRptEW.Left
    frmRptEW.Height = Image2.Top - frmRptEW.Top
End Sub

Private Sub lblRptEmp_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblRptEmp.BackStyle = 1
    lblRptEmp.BackColor = &H80FF&
End Sub

Private Sub lblRptItems_Click()
On Error Resume Next
    frmRptIW.Show , Me
    frmRptIW.Left = PicMenu.Width
    frmRptIW.Top = 2050
    frmRptIW.Width = Me.Width - frmRptIW.Left
    frmRptIW.Height = Image2.Top - frmRptIW.Top
End Sub

Private Sub lblRptItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblRptItems.BackStyle = 1
    lblRptItems.BackColor = &H80FF&
End Sub

Private Sub lblTrans_Click()
On Error Resume Next
    If PicReport.Visible = True Then
        PicReport.Visible = False
    End If
    
    If PicTrans.Visible = False Then
        PicTrans.Visible = True
        PicTrans.Top = lblTrans.Top + lblTrans.Height + 165
        lblReport.Top = PicTrans.Top + PicTrans.Height + 165
        lblExit.Top = lblReport.Top + lblReport.Height + 165
    ElseIf PicTrans.Visible = True Then
        PicTrans.Visible = False
        lblReport.Top = lblTrans.Top + lblTrans.Height + 165
        lblExit.Top = lblReport.Top + lblReport.Height + 165
    End If
    'Image1.Top = lblReport.Top + lblReport.Height + 165
    'Image1.Height = Image2.Top - Image1.Top
End Sub

Private Sub lblTrans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    lblTrans.BackStyle = 1
    lblTrans.BackColor = &H80FF&
End Sub

Private Sub PicMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
     Call MouseMove
End Sub

Private Sub PicReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Call MouseMove
End Sub

Private Sub PicTrans_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Call MouseMove
End Sub
Public Function MouseMove()
On Error Resume Next
    lblCStock.BackStyle = 0
    lblCStock.BackColor = &H80000009
    lblEmpMst.BackStyle = 0
    lblEmpMst.BackColor = &H80000009
    lblItems.BackStyle = 0
    lblItems.BackColor = &H80000009
    lblTrans.BackStyle = 0
    lblTrans.BackColor = &H80000009
    lblReceive.BackStyle = 0
    lblReceive.BackColor = &H80000009
    lblIssue.BackStyle = 0
    lblIssue.BackColor = &H80000009
    lblReturn.BackStyle = 0
    lblReturn.BackColor = &H80000009
    lblDad.BackStyle = 0
    lblDad.BackColor = &H80000009
    lblReport.BackStyle = 0
    lblReport.BackColor = &H80000009
    lblRptEmp.BackStyle = 0
    lblRptEmp.BackColor = &H80000009
    lblRptItems.BackStyle = 0
    lblRptItems.BackColor = &H80000009
    lblEmpList.BackStyle = 0
    lblEmpList.BackColor = &H80000009
    lblExit.BackStyle = 0
    lblExit.BackColor = &H80000009
End Function
