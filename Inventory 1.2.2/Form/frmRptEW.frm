VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptEW 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Employee Wise Inventory"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView LVEmp 
      Height          =   2970
      Left            =   330
      TabIndex        =   3
      Top             =   1920
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   5239
      SortKey         =   1
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SrNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Items"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Receive"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "RcvDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Issue By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Return"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "RtnDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Return Receive By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Dad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "DadDate"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox cmbEmpName 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   5895
   End
   Begin MSComctlLib.ListView LVTotal 
      Height          =   2610
      Left            =   2325
      TabIndex        =   4
      Top             =   5280
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   4604
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SrNo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Items"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Receive"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Return"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee's Items wise Total  :-"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   4920
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name :-"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   9240
      Picture         =   "frmRptEW.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   540
   End
   Begin VB.Image Image8 
      Height          =   1020
      Left            =   1680
      Picture         =   "frmRptEW.frx":0B0C
      Top             =   360
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Inventory Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   600
      Left            =   3600
      TabIndex        =   0
      Top             =   240
      Width           =   5790
   End
End
Attribute VB_Name = "frmRptEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub cmbEmpName_Change()
On Error Resume Next
    cmbEmpName = UCase(cmbEmpName)
    SendKeys "{End}"
End Sub

Private Sub cmbEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbEmpName, cmbEmpName.Text)
End Sub

Private Sub cmbEmpName_LostFocus()
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
    
Dim i, a, b, c As Integer
LVEmp.ListItems.Clear
LVTotal.ListItems.Clear
i = 1
    rs.Open "Select SrNo, IName, ISize from Items ", con, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            LVEmp.ListItems.Add i, , rs!SrNo
            LVEmp.ListItems(i).SubItems(1) = rs!IName
            LVEmp.ListItems(i).SubItems(2) = rs!ISize
            LVTotal.ListItems.Add i, , rs!SrNo
            LVTotal.ListItems(i).SubItems(1) = rs!IName
            LVTotal.ListItems(i).SubItems(2) = rs!ISize
            a = i
            rs1.Open "Select IssItems,Issue, IssDate, Issueby from Issue where Issue.IReceiveby = '" & UCase(cmbEmpName) & "' and Issue.IssItems = '" & LVEmp.ListItems(i).SubItems(1) & "'", con, adOpenDynamic, adLockOptimistic
                
                Do While Not rs1.EOF
                'If a = 1 Then
                 If a = rs!SrNo Then
                  
                    LVEmp.ListItems(a).SubItems(3) = rs1!Issue
                    LVEmp.ListItems(a).SubItems(4) = rs1!IssDate
                    LVEmp.ListItems(a).SubItems(5) = rs1!Issueby
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Issue
                        'LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(3) - LVTotal.ListItems(i).SubItems(4)
                    End If
                Else
                    'LVEmp.ListItems.Add a, , rs!SrNo & "." & a
                    LVEmp.ListItems.Add a, , rs!SrNo + (a / 1000)
                    LVEmp.ListItems(a).SubItems(1) = rs!IName
                    LVEmp.ListItems(a).SubItems(2) = rs!ISize
                    LVEmp.ListItems(a).SubItems(3) = rs1!Issue
                    LVEmp.ListItems(a).SubItems(4) = rs1!IssDate
                    LVEmp.ListItems(a).SubItems(5) = rs1!Issueby
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Issue
                    End If
                End If
                If LVEmp.ListItems(a).SubItems(3) = "" Then
                    LVEmp.ListItems(a).SubItems(3) = 0
                End If
                If LVEmp.ListItems(a).SubItems(4) = "" Then
                    LVEmp.ListItems(a).SubItems(4) = "--"
                End If
                If LVEmp.ListItems(a).SubItems(5) = "" Then
                    LVEmp.ListItems(a).SubItems(5) = "--"
                End If
                If LVEmp.ListItems(a).SubItems(6) = "" Then
                    LVEmp.ListItems(a).SubItems(6) = 0
                End If
                If LVEmp.ListItems(a).SubItems(7) = "" Then
                    LVEmp.ListItems(a).SubItems(7) = "--"
                End If
                If LVEmp.ListItems(a).SubItems(8) = "" Then
                    LVEmp.ListItems(a).SubItems(8) = "--"
                End If
                If LVEmp.ListItems(a).SubItems(9) = "" Then
                    LVEmp.ListItems(a).SubItems(9) = 0
                End If
                If LVEmp.ListItems(a).SubItems(10) = "" Then
                    LVEmp.ListItems(a).SubItems(10) = "--"
                End If
                a = a + 1
                rs1.MoveNext
                Loop
            b = i
            rs2.Open "Select RtnItems,Return, RtnDate, RReceiveby from Return where Return.Returnby ='" & UCase(cmbEmpName) & "' and Return.RtnItems = '" & LVEmp.ListItems(i).SubItems(1) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs2.EOF
                'If b = 1 Then
                If b = rs!SrNo Then
                    LVEmp.ListItems(b).SubItems(6) = rs2!Return
                    LVEmp.ListItems(b).SubItems(7) = rs2!RtnDate
                    LVEmp.ListItems(b).SubItems(8) = rs2!RReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Return
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Return
                    
                    End If
                Else
                    'LVEmp.ListItems.Add b, , rs!SrNo + (a / 1000)
                    'LVEmp.ListItems(b).SubItems(1) = rs!IName
                    'LVEmp.ListItems(b).SubItems(2) = rs!ISize
                    LVEmp.ListItems(b).SubItems(6) = rs2!Return
                    LVEmp.ListItems(b).SubItems(7) = rs2!RtnDate
                    LVEmp.ListItems(b).SubItems(8) = rs2!RReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Return
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Return
                    
                    End If
                End If
                If LVEmp.ListItems(b).SubItems(3) = "" Then
                    LVEmp.ListItems(b).SubItems(3) = 0
                End If
                If LVEmp.ListItems(b).SubItems(4) = "" Then
                    LVEmp.ListItems(b).SubItems(4) = "--"
                End If
                If LVEmp.ListItems(b).SubItems(5) = "" Then
                    LVEmp.ListItems(b).SubItems(5) = "--"
                End If
                If LVEmp.ListItems(b).SubItems(6) = "" Then
                    LVEmp.ListItems(b).SubItems(6) = 0
                End If
                If LVEmp.ListItems(b).SubItems(7) = "" Then
                    LVEmp.ListItems(b).SubItems(7) = "--"
                End If
                If LVEmp.ListItems(b).SubItems(8) = "" Then
                    LVEmp.ListItems(b).SubItems(8) = "--"
                End If
                If LVEmp.ListItems(b).SubItems(9) = "" Then
                    LVEmp.ListItems(b).SubItems(9) = 0
                End If
                If LVEmp.ListItems(b).SubItems(10) = "" Then
                    LVEmp.ListItems(b).SubItems(10) = "--"
                End If
                b = b + 1
                rs2.MoveNext
                Loop
            
            c = i
            rs3.Open "Select DadItems,Dad, DadDate from Dad where Dad.Dadby ='" & UCase(cmbEmpName) & "' and Dad.DadItems = '" & LVEmp.ListItems(i).SubItems(1) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs3.EOF
                'If c = 1 Then
                If c = rs!SrNo Then
                    LVEmp.ListItems(c).SubItems(9) = rs3!Dad
                    LVEmp.ListItems(c).SubItems(10) = rs3!DadDate
                Else
                    'LVEmp.ListItems.Add c, , rs!SrNo + (a / 1000)
                    'LVEmp.ListItems(c).SubItems(1) = rs!IName
                    'LVEmp.ListItems(c).SubItems(2) = rs!ISize
                    LVEmp.ListItems(c).SubItems(9) = rs3!Dad
                    LVEmp.ListItems(c).SubItems(10) = rs3!DadDate
                End If
                If LVEmp.ListItems(c).SubItems(3) = "" Then
                    LVEmp.ListItems(c).SubItems(3) = 0
                End If
                If LVEmp.ListItems(c).SubItems(4) = "" Then
                    LVEmp.ListItems(c).SubItems(4) = "--"
                End If
                If LVEmp.ListItems(c).SubItems(5) = "" Then
                    LVEmp.ListItems(c).SubItems(5) = "--"
                End If
                If LVEmp.ListItems(c).SubItems(6) = "" Then
                    LVEmp.ListItems(c).SubItems(6) = 0
                End If
                If LVEmp.ListItems(c).SubItems(7) = "" Then
                    LVEmp.ListItems(c).SubItems(7) = "--"
                End If
                If LVEmp.ListItems(c).SubItems(8) = "" Then
                    LVEmp.ListItems(c).SubItems(8) = "--"
                End If
                If LVEmp.ListItems(c).SubItems(9) = "" Then
                    LVEmp.ListItems(c).SubItems(9) = 0
                End If
                If LVEmp.ListItems(c).SubItems(10) = "" Then
                    LVEmp.ListItems(c).SubItems(10) = "--"
                End If
                c = c + 1
                rs3.MoveNext
                Loop
        
                If LVEmp.ListItems(i).SubItems(3) = "" Then
                    LVEmp.ListItems(i).SubItems(3) = 0
                End If
                If LVEmp.ListItems(i).SubItems(4) = "" Then
                    LVEmp.ListItems(i).SubItems(4) = "--"
                End If
                If LVEmp.ListItems(i).SubItems(5) = "" Then
                    LVEmp.ListItems(i).SubItems(5) = "--"
                End If
                If LVEmp.ListItems(i).SubItems(6) = "" Then
                    LVEmp.ListItems(i).SubItems(6) = 0
                End If
                If LVEmp.ListItems(i).SubItems(7) = "" Then
                    LVEmp.ListItems(i).SubItems(7) = "--"
                End If
                If LVEmp.ListItems(i).SubItems(8) = "" Then
                    LVEmp.ListItems(i).SubItems(8) = "--"
                End If
                If LVEmp.ListItems(i).SubItems(9) = "" Then
                    LVEmp.ListItems(i).SubItems(9) = 0
                End If
                If LVEmp.ListItems(i).SubItems(10) = "" Then
                    LVEmp.ListItems(i).SubItems(10) = "--"
                End If
                If LVTotal.ListItems(i).SubItems(3) = "" Then
                    LVTotal.ListItems(i).SubItems(3) = 0
                End If
                If LVTotal.ListItems(i).SubItems(4) = "" Then
                    LVTotal.ListItems(i).SubItems(4) = 0
                End If
            LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(3) - LVTotal.ListItems(i).SubItems(4)
        i = i + 1
        rs.MoveNext
        rs1.Close
    Loop
LVEmp.Sorted = True
LVEmp.SortKey = 0
     
rs.Close
'rs1.Close
rs2.Close
rs3.Close
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
End Sub

Private Sub Form_Load()
On Error Resume Next
    con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    con.Open
Call ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Set con = Nothing
End Sub
Public Function ClearAll()
On Error Resume Next

FeedData "EmpMaster", "EmpName", cmbEmpName
LVEmp.ListItems.Clear
LVTotal.ListItems.Clear
cmbEmpName.SetFocus
End Function

Private Sub ImgClose_Click()
On Error Resume Next
    Unload Me
End Sub
