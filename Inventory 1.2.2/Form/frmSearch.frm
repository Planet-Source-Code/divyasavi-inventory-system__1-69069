VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSearch 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H80000009&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.ComboBox cmbSearch 
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
      Left            =   3360
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin MSComctlLib.ListView LVSearch 
      Height          =   5055
      Left            =   2745
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8916
      View            =   3
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   9240
      Picture         =   "frmSearch.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   540
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image13 
      Height          =   1020
      Left            =   2280
      Picture         =   "frmSearch.frx":0B0C
      Top             =   480
      Width           =   1020
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Contmp As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub cmbSearch_Change()
On Error Resume Next
    cmbSearch = UCase(cmbSearch)
    SendKeys "{End}"
End Sub

Private Sub cmbSearch_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If Me.Caption = "Items" Then
        Call Items_KeyDown
    ElseIf Me.Caption = "Receive" Then
        Call Receive_KeyDown
    ElseIf Me.Caption = "Issue" Then
        Call Issue_KeyDown
    ElseIf Me.Caption = "Return" Then
        Call Return_KeyDown
    ElseIf Me.Caption = "Dad" Then
        Call Dad_KeyDown
    ElseIf Me.Caption = "EmpMaster" Then
        Call Emp_KeyDown
    Else
        MsgBox "Please Enter Valid Data.", vbCritical, "Search"
    End If
End Sub

Private Sub cmbSearch_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbSearch, cmbSearch.Text)
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
    Call LVSearch_Click
End Sub

Private Sub Form_Load()
On Error Resume Next
Contmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
Contmp.Open
    
    'ImgClose.Left = cmdOk.Left + cmdOk.Width + 100
    'ImgClose.Top = cmdOk.Top
    
    If frmItems.cmdFind = True Then
        Me.Caption = "Items"
    ElseIf frmReceive.cmdFind = True Then
        Me.Caption = "Receive"
    ElseIf frmIssue.cmdFind = True Then
        Me.Caption = "Issue"
    ElseIf frmReturn.cmdFind = True Then
        Me.Caption = "Return"
    ElseIf frmDad.cmdFind = True Then
        Me.Caption = "Dad"
    ElseIf frmEmpMst.cmdFind = True Then
        Me.Caption = "EmpMaster"
    Else
        Me.Caption = ""
    End If
    
    
        
    If Me.Caption = "Items" Then
        Call ItemsSearch
    ElseIf Me.Caption = "Receive" Then
        Call RcvSearch
    ElseIf Me.Caption = "Issue" Then
        Call IssSearch
    ElseIf Me.Caption = "Return" Then
        Call RtnSearch
    ElseIf Me.Caption = "Dad" Then
        Call DadSearch
    ElseIf Me.Caption = "EmpMaster" Then
        Call EmpSearch
    Else
        MsgBox "Please Enter valid Data ", vbInformation, "Searcher"
    End If
    cmbSearch.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    'Me.Caption = ""
    Set rs = Nothing
    Set Contmp = Nothing
End Sub

Public Function ItemsSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Item Name", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Size", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Open Stock", 1200, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove 1
    
    FeedData "Items", "IName", cmbSearch
    lblSearch.Caption = "Items"
End Function

Public Function RcvSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Rcv Name", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Size", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Rcv Qty", 1300, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 6, , "Rcv Date", 1100, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 7, , "Rcv By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove (1)
    
    FeedData "Receive", "RcvItems", cmbSearch
    lblSearch.Caption = "Receive Items"
    cmbSearch.SetFocus
End Function

Public Function IssSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Iss Item", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Size", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Iss Qty", 1300, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 6, , "Iss Date", 1100, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 7, , "Iss By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 8, , "Rcv By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove (1)
    
    FeedData "Issue", "IssItems", cmbSearch
    lblSearch.Caption = "Issue Items"
    cmbSearch.SetFocus
End Function

Public Function RtnSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Rtn Item", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Size", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Rtn Qty", 1300, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 6, , "Rtn Date", 1100, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 7, , "Rtn By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 8, , "Rtn Rcv By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove (1)
    
    FeedData "Return", "RtnItems", cmbSearch
    lblSearch.Caption = "Return Items"
    cmbSearch.SetFocus
End Function

Public Function DadSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Dad Item", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Size", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Dad Qty", 1300, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 6, , "Dad Date", 1100, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 7, , "Dad By", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove (1)
    
    FeedData "Return", "RtnItems", cmbSearch
    lblSearch.Caption = "Dad Items"
    cmbSearch.SetFocus
End Function

Public Function EmpSearch()
On Error Resume Next
    LVSearch.ColumnHeaders.Add 1, , "Sr No"
    'LVSearch.ColumnHeaders.Add 2, , "Sr No", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 2, , "Emp Code", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 3, , "Emp No.", 800, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 4, , "Emp Name", 2400, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 5, , "Desig", 1200, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 6, , "Address", 4000, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 7, , "City", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 8, , "Phone", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Add 9, , "Mobile", 1500, lvwColumnCenter
    LVSearch.ColumnHeaders.Remove (1)
    
    FeedData "EmpMaster", "EmpName", cmbSearch
    lblSearch.Caption = "Employee Name"
    cmbSearch.SetFocus
End Function

Public Function Items_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        .Open "Select SrNo, IName, ISize, OpnStock from Items where Items.IName Like '" & cmbSearch & "%' order by Items.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !IName
            LVSearch.ListItems(i).SubItems(2) = !ISize
            LVSearch.ListItems(i).SubItems(3) = !OpnStock
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Public Function Receive_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        .Open "Select SrNo, RcvItems, RcvSize, Receive, RcvDate, RcvBy from Receive where Receive.RcvItems Like '" & cmbSearch & "%' order by Receive.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !RcvItems
            LVSearch.ListItems(i).SubItems(2) = !RcvSize
            LVSearch.ListItems(i).SubItems(3) = !Receive
            LVSearch.ListItems(i).SubItems(4) = !RcvDate
            LVSearch.ListItems(i).SubItems(5) = !RcvBy
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Public Function Issue_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        
        .Open "Select SrNo, IssItems, IssSize, Issue, IssDate, IssueBy, IReceiveBy from Issue where Issue.IssItems Like '" & cmbSearch & "%' order by Issue.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !IssItems
            LVSearch.ListItems(i).SubItems(2) = !IssSize
            LVSearch.ListItems(i).SubItems(3) = !Issue
            LVSearch.ListItems(i).SubItems(4) = !IssDate
            LVSearch.ListItems(i).SubItems(5) = !Issueby
            LVSearch.ListItems(i).SubItems(6) = !IReceiveby
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Public Function Return_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        .Open "Select SrNo, RtnItems, RtnSize, Return, RtnDate, ReturnBy, RReceiveBy from Return where Return.RtnItems Like '" & cmbSearch & "%' order by Return.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !RtnItems
            LVSearch.ListItems(i).SubItems(2) = !RtnSize
            LVSearch.ListItems(i).SubItems(3) = !Return
            LVSearch.ListItems(i).SubItems(4) = !RtnDate
            LVSearch.ListItems(i).SubItems(5) = !Returnby
            LVSearch.ListItems(i).SubItems(6) = !RReceiveby
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Public Function Dad_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        .Open "Select SrNo, DadItems, DadSize, Dad, DadDate, DadBy from Dad where Dad.DadItems Like '" & cmbSearch & "%' order by Dad.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !DadItems
            LVSearch.ListItems(i).SubItems(2) = !DadSize
            LVSearch.ListItems(i).SubItems(3) = !Dad
            LVSearch.ListItems(i).SubItems(4) = !DadDate
            LVSearch.ListItems(i).SubItems(5) = !Dadby
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Public Function Emp_KeyDown()
On Error Resume Next
Dim i As Integer
i = 1
    LVSearch.ListItems.Clear
    With rs
        .Open "Select EmpCode, SrNo, EmpName, Desig, Add, City, Phone, Mobile from EmpMaster where EmpMaster.EmpName Like '" & cmbSearch & "%' order by EmpMaster.SrNo", Contmp, adOpenDynamic, adLockOptimistic
        Do While Not .EOF
            LVSearch.ListItems.Add i, , !SrNo
            LVSearch.ListItems(i).SubItems(1) = !EmpCode
            LVSearch.ListItems(i).SubItems(2) = !EmpName
            LVSearch.ListItems(i).SubItems(3) = !Desig
            LVSearch.ListItems(i).SubItems(4) = !Add
            LVSearch.ListItems(i).SubItems(5) = !City
            LVSearch.ListItems(i).SubItems(6) = !Phone
            LVSearch.ListItems(i).SubItems(7) = !Mobile
            i = i + 1
            .MoveNext
        Loop
        .Close
    End With
End Function

Private Sub ImgClose_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub LVSearch_Click()
On Error Resume Next
    If LVSearch.ListItems.Count <= 0 Then Exit Sub
    
    If Me.Caption = "Items" Then
        Call ItemsClick
    ElseIf Me.Caption = "Receive" Then
        Call ReceiveClick
    ElseIf Me.Caption = "Issue" Then
        Call IssueClick
    ElseIf Me.Caption = "Return" Then
        Call ReturnClick
    ElseIf Me.Caption = "Dad" Then
        Call DadClick
    ElseIf Me.Caption = "EmpMaster" Then
        Call EmpMasterClick
    Else
        MsgBox "Please Enter valid Data ", vbInformation, "Searcher"
    End If
    
End Sub

Private Sub LVSearch_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = 13 Then Call LVSearch_Click
End Sub
Public Function ItemsClick()
On Error Resume Next
    With rs
        .Open "Select * from Items where Items.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmItems.txtSrNo = !SrNo
            frmItems.cmbIName = !IName
            frmItems.cmbISize = !ISize
            frmItems.txtItems = !OpnStock
        .Close
    End With
Me.Caption = ""
Unload Me
End Function

Public Function ReceiveClick()
On Error Resume Next
    With rs
        .Open "Select * from Receive where Receive.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmReceive.txtSrNo = !SrNo
            frmReceive.cmbRcvIName = !RcvItems
            frmReceive.cmbISize = !RcvSize
            frmReceive.txtRcvIQty = !Receive
            frmReceive.RcvDate = !RcvDate
            frmReceive.cmbRcvIBy = !RcvBy
        .Close
    End With
Me.Caption = ""
Unload Me
End Function

Public Function IssueClick()
On Error Resume Next
    With rs
        .Open "Select * from Issue where Issue.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmIssue.txtSrNo = !SrNo
            frmIssue.cmbIssIName = !IssItems
            frmIssue.cmbISize = !IssSize
            frmIssue.txtIssIQty = !Issue
            frmIssue.IssDate = !IssDate
            frmIssue.cmbIssIBy = !Issueby
            frmIssue.cmbRcvIBy = !IReceiveby
        .Close
    End With
Me.Caption = ""
Unload Me
End Function

Public Function ReturnClick()
On Error Resume Next
    With rs
        .Open "Select * from Return where Return.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmReturn.txtSrNo = !SrNo
            frmReturn.cmbRtnIName = !RtnItems
            frmReturn.cmbISize = !RtnSize
            frmReturn.txtRtnIQty = !Return
            frmReturn.RtnDate = !RtnDate
            frmReturn.cmbRtnIBy = !Returnby
            frmReturn.cmbRRcvIBy = !RReceiveby
        .Close
    End With
Me.Caption = ""
Unload Me
End Function

Public Function DadClick()
On Error Resume Next
    With rs
        .Open "Select * from Dad where Dad.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmDad.txtSrNo = !SrNo
            frmDad.cmbDadIName = !DadItems
            frmDad.cmbISize = !DadSize
            frmDad.txtDadIQty = !Dad
            frmDad.DadDate = !DadDate
            frmDad.cmbDadIBy = !Dadby
        .Close
    End With
Me.Caption = ""
Unload Me
End Function

Public Function EmpMasterClick()
On Error Resume Next
    With rs
        .Open "Select * from EmpMaster where EmpMaster.SrNo = " & LVSearch.SelectedItem.Text, Contmp, adOpenDynamic, adLockOptimistic
            frmEmpMst.txtSrNo = !SrNo
            frmEmpMst.txtEmpCode = !EmpCode
            frmEmpMst.cmbEmpName = !EmpName
            frmEmpMst.cmbDesig = !Desig
            frmEmpMst.txtAdd = !Add
            frmEmpMst.cmbCity = !City
            frmEmpMst.txtPhone = !Phone
            frmEmpMst.txtMobile = !Mobile
        .Close
    End With
Me.Caption = ""
Unload Me
End Function
