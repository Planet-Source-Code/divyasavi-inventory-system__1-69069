VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIssue 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Issue Item"
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
      TabIndex        =   9
      Top             =   5880
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
      Left            =   4080
      TabIndex        =   17
      Text            =   "0"
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbIssIBy 
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
      Left            =   4080
      TabIndex        =   3
      Top             =   4080
      Width           =   3855
   End
   Begin VB.ComboBox cmbIssIName 
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
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
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
      Left            =   4080
      TabIndex        =   1
      Top             =   3120
      Width           =   3855
   End
   Begin VB.TextBox txtIssIQty 
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
      Left            =   4080
      TabIndex        =   2
      Top             =   3600
      Width           =   3855
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
      Left            =   4080
      TabIndex        =   4
      Top             =   4560
      Width           =   3855
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000009&
      Caption         =   "&Close"
      Height          =   495
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000009&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "&Save"
      Height          =   495
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000009&
      Caption         =   "&New"
      Height          =   495
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker IssDate 
      Height          =   330
      Left            =   4080
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39274
   End
   Begin VB.Image Image6 
      Height          =   1020
      Left            =   2760
      Picture         =   "frmIssue.frx":0000
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ISSUE STOCK ENTRY"
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
      Left            =   3840
      TabIndex        =   21
      Top             =   1080
      Width           =   5445
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Current Stock"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   20
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblCStock 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8400
      TabIndex        =   19
      Top             =   3240
      Width           =   1575
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
      Left            =   2640
      TabIndex        =   18
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Issue  By :-"
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
      Left            =   2280
      TabIndex        =   16
      Top             =   4080
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Item Name :-"
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
      Left            =   2040
      TabIndex        =   15
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Item Qty :-"
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
      Left            =   1800
      TabIndex        =   14
      Top             =   3600
      Width           =   2055
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
      Left            =   2640
      TabIndex        =   13
      Top             =   3120
      Width           =   1215
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
      Left            =   2280
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date :-"
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
      Left            =   2520
      TabIndex        =   11
      Top             =   5040
      Width           =   1335
   End
End
Attribute VB_Name = "frmIssue"
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
If cmbISize.Text = "" Then
    Exit Sub
Else
    CheckData "Items", "ISize", cmbISize.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item Size , " & cmbISize.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbISize.Text = ""
        cmbISize.SetFocus
        Exit Sub
    'Else
    '    txtIssIQty.SetFocus
    End If
End If

Call CStock
If lblCStock = 0 Then
    MsgBox "You can't Issue " & UCase(cmbIssIName) & ", 0 (Zero) Stock available.", vbCritical, Me.Caption
    Call ClearAll
End If
End Sub

Private Sub cmbIssIBy_Change()
On Error Resume Next
    cmbIssIBy = UCase(cmbIssIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbIssIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbIssIBy, cmbIssIBy.Text)
End Sub

Private Sub cmbIssIBy_LostFocus()
On Error Resume Next
If cmbIssIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbIssIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Issuer , " & cmbIssIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbIssIBy.Text = ""
        cmbIssIBy.SetFocus
    'Else
    '    cmbRcvIBy.SetFocus
    End If
End If
End Sub

Private Sub cmbIssIName_LostFocus()
On Error Resume Next
If cmbIssIName = "" Then
    Exit Sub
Else
    CheckData "Items", "IName", cmbIssIName.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item , " & cmbIssIName.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbIssIName.Text = ""
        cmbIssIName.SetFocus
    Else
        cmbISize.SetFocus
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

Private Sub cmbIssIName_Change()
On Error Resume Next
    cmbIssIName = UCase(cmbIssIName)
    SendKeys "{End}"
End Sub

Private Sub cmbIssIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbIssIName, cmbIssIName.Text)
End Sub

Private Sub cmbRcvIBy_LostFocus()
On Error Resume Next
If cmbRcvIBy = "" Then
    Exit Sub
Else
   CheckData "EmpMaster", "EmpName", cmbRcvIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item Receiver , " & cmbRcvIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbRcvIBy.Text = ""
        cmbRcvIBy.SetFocus
    'Else
    '    IssDate.SetFocus
    End If
End If
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Issue Where Issue.SrNo = " & Val(txtSrNo)
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    'frmSearch.Caption = "Issue"
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
If cmbIssIName = "" Then
    MsgBox "Plese Select Item Name.", vbCritical, Me.Caption
    cmbIssIName.SetFocus
    Exit Sub
End If
If cmbISize = "" Then
    MsgBox "Please Select Item Size ", vbCritical, Me.Caption
    cmbISize.SetFocus
    Exit Sub
End If
If txtIssIQty = "" Then
    MsgBox "Please Enter Issue Quantity ", vbCritical, Me.Caption
    txtIssIQty.SetFocus
    Exit Sub
End If
If cmbIssIBy = "" Then
    MsgBox "Please Select Issue By Name ", vbCritical, Me.Caption
    cmbIssIBy.SetFocus
    Exit Sub
End If
If cmbRcvIBy = "" Then
    MsgBox "Please Select Receive By Name ", vbCritical, Me.Caption
    cmbRcvIBy.SetFocus
    Exit Sub
End If

With rs
    '.Open "Select * from Issue where IssItems = '" & UCase(cmbIssIName) & "'", con, adOpenDynamic, adLockOptimistic
    .Open "Select * from Issue where SrNo = '" & Val(txtSrNo) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from Issue", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !SrNo = GetNewNo("Issue")
        !IssItems = UCase(cmbIssIName)
        !IssSize = UCase(cmbISize)
        !Issue = txtIssIQty
        !Issueby = UCase(cmbIssIBy)
        !IReceiveby = UCase(cmbRcvIBy)
        !IssDate = IssDate
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !SrNo = txtSrNo
        !IssItems = UCase(cmbIssIName)
        !IssSize = cmbISize
        !Issue = txtIssIQty
        !Issueby = UCase(cmbIssIBy)
        !IReceiveby = UCase(cmbRcvIBy)
        !IssDate = IssDate
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
Public Function CStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
Dim rstmp1 As New ADODB.Recordset
Dim rstmp2 As New ADODB.Recordset
Dim rstmp3 As New ADODB.Recordset
Dim rstmp4 As New ADODB.Recordset
Dim a, b, c, d, E As Integer
    
    rstmp.Open "Select OpnStock from Items where Items.IName ='" & UCase(cmbIssIName) & "' and Items.ISize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            a = 0
        Else
            a = rstmp(0)
        End If
    
    rstmp1.Open "Select sum(Receive) from Receive where Receive.RcvItems ='" & UCase(cmbIssIName) & "' and Receive.RcvSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp1(0)) Then
            b = 0
        Else
            b = rstmp1(0)
        End If
    rstmp2.Open "Select sum(Issue) from Issue where Issue.IssItems ='" & UCase(cmbIssIName) & "' and Issue.IssSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp2(0)) Then
            c = 0
        Else
            c = rstmp2(0)
        End If
    rstmp3.Open "Select sum(Return) from Return where Return.RtnItems ='" & UCase(cmbIssIName) & "' and Return.RtnSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp3(0)) Then
            d = 0
        Else
            d = rstmp3(0)
        End If
    rstmp4.Open "Select sum(Dad) from Dad where Dad.DadItems ='" & UCase(cmbIssIName) & "' and Dad.DadSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp4(0)) Then
            E = 0
        Else
            E = rstmp4(0)
        End If
    
    lblCStock = a + b - c + d - E
    rstmp.Close
    rstmp1.Close
    rstmp2.Close
    rstmp3.Close
    rstmp4.Close
    
Set rstmp = Nothing
Set rstmp1 = Nothing
Set rstmp2 = Nothing
Set rstmp3 = Nothing
Set rstmp4 = Nothing

End Function
Public Function ClearAll()
On Error Resume Next

FeedData "Items", "IName", cmbIssIName
FeedData "Items", "ISize", cmbISize
FeedData "EmpMaster", "EmpName", cmbIssIBy
FeedData "EmpMaster", "EmpName", cmbRcvIBy
txtSrNo = GetNewNo("Issue")
cmbIssIName.Text = ""
cmbISize.Text = ""
txtIssIQty.Text = ""
cmbIssIBy.Text = ""
cmbRcvIBy.Text = ""
IssDate = Date

cmbIssIName.SetFocus
End Function
