VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReceive 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "New Items Receive"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H80000009&
      Caption         =   "&Find"
      Height          =   495
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtSrNo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   15
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000009&
      Caption         =   "&New"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000009&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000009&
      Caption         =   "&Close"
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbRcvIBy 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   3
      Top             =   4440
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker RcvDate 
      Height          =   330
      Left            =   4920
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39274
   End
   Begin VB.TextBox txtRcvIQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   355
      Left            =   4920
      TabIndex        =   2
      Top             =   3960
      Width           =   3855
   End
   Begin VB.ComboBox cmbISize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4920
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin VB.ComboBox cmbRcvIName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmReceive.frx":0000
      Left            =   4920
      List            =   "frmReceive.frx":0002
      TabIndex        =   0
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Image Image7 
      Height          =   1020
      Left            =   2160
      Picture         =   "frmReceive.frx":0004
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIVE STOCK ENTRY"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   600
      Left            =   3240
      TabIndex        =   17
      Top             =   1320
      Width           =   6315
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Trans. No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Date :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   14
      Top             =   4920
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Receive  By :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3120
      TabIndex        =   13
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Size :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   12
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Item Qty :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Item Name :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2640
      TabIndex        =   10
      Top             =   3000
      Width           =   2055
   End
End
Attribute VB_Name = "frmReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbISize_Change()
On Error Resume Next
    cmbISize = UCase(cmbISize)
    SendKeys "{End}"
End Sub

Private Sub cmbISize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbISize, cmbISize.Text)
End Sub

Private Sub cmbISize_LostFocus()
On Error Resume Next
If cmbISize = "" Then
    Exit Sub
Else
    CheckData "Items", "ISize", cmbISize.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item Size , " & cmbISize.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbISize.Text = ""
        cmbISize.SetFocus
    'Else
    '    txtRcvIQty.SetFocus
    End If
End If
End Sub

Private Sub cmbRcvIBy_Change()
On Error Resume Next
    cmbRcvIBy = UCase(cmbRcvIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbRcvIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbRcvIBy, cmbRcvIBy.Text)
End Sub

Private Sub cmbRcvIBy_LostFocus()
On Error Resume Next
If cmbRcvIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbRcvIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Receiver Name , " & cmbRcvIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRcvIBy.Text = ""
        cmbRcvIBy.SetFocus
    'Else
    '    RcvDate.SetFocus
    End If
End If
End Sub

Private Sub cmbRcvIName_Change()
On Error Resume Next
    cmbRcvIName = UCase(cmbRcvIName)
    SendKeys "{End}"
End Sub

Private Sub cmbRcvIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbRcvIName, cmbRcvIName.Text)
End Sub

Private Sub cmbRcvIName_LostFocus()
On Error Resume Next
If cmbRcvIName = "" Then
    Exit Sub
Else
    CheckData "Items", "IName", cmbRcvIName.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item , " & cmbRcvIName.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRcvIName.Text = ""
        cmbRcvIName.SetFocus
    'Else
    '    cmbISize.SetFocus
    End If
End If
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Receive Where Receive.SrNo = " & Val(txtSrNo)
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    'frmSearch.Caption = "Receive"
    frmSearch.Left = frmMain.PicMenu.Width
    frmSearch.Top = 2050
    frmSearch.Width = frmMain.Width - frmSearch.Left
    frmSearch.Height = frmMain.Image2.Top - frmSearch.Top
    frmSearch.Show , frmMain
End Sub

Private Sub cmdNew_Click()
On Error Resume Next
    Call ClearAll
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
If cmbRcvIName = "" Then
    MsgBox "Please Select Item Name ", vbCritical, Me.Caption
    cmbRcvIName.SetFocus
    Exit Sub
End If
If cmbISize = "" Then
    MsgBox "Please Select Item Size ", vbCritical, Me.Caption
    cmbISize.SetFocus
    Exit Sub
End If
If txtRcvIQty = "" Then
    MsgBox "Please Enter Receive Quantity ", vbCritical, Me.Caption
    txtRcvIQty.SetFocus
    Exit Sub
End If
If cmbRcvIBy = "" Then
    MsgBox "Please Select Receive By Name ", vbCritical, Me.Caption
    cmbRcvIBy.SetFocus
    Exit Sub
End If

With rs
    '.Open "Select * from Receive where RcvItems = '" & UCase(cmbRcvIName) & "'", con, adOpenDynamic, adLockOptimistic
    .Open "Select * from Receive where SrNo = '" & Val(txtSrNo) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from Receive ", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !SrNo = GetNewNo("Receive")
        !RcvItems = UCase(cmbRcvIName)
        !RcvSize = UCase(cmbISize)
        !Receive = txtRcvIQty
        !RcvBy = cmbRcvIBy
        !RcvDate = RcvDate
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !SrNo = txtSrNo
        !RcvItems = UCase(cmbRcvIName)
        !RcvSize = UCase(cmbISize)
        !Receive = txtRcvIQty
        !RcvBy = cmbRcvIBy
        !RcvDate = RcvDate
        .Update
        .Close
        MsgBox "Information is Updated", vbInformation, Me.Caption
    End If
    
End With
Set rs = Nothing
Call ClearAll
End Sub

Private Sub Form_Load()
On Error Resume Next
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    con.Open
Call ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set rs = Nothing
    Set con = Nothing
End Sub
Public Function ClearAll()
On Error Resume Next

FeedData "Items", "IName", cmbRcvIName
FeedData "Items", "ISize", cmbISize
FeedData "EmpMaster", "EmpName", cmbRcvIBy
txtSrNo = GetNewNo("Receive")
cmbRcvIName.Text = ""
cmbISize.Text = ""
txtRcvIQty.Text = ""
cmbRcvIBy.Text = ""
RcvDate = Date
cmbRcvIName.SetFocus
End Function

