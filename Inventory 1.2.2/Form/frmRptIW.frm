VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRptIW 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Item wise Inventory"
   ClientHeight    =   8130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbSize 
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
      Left            =   6840
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox cmbIName 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   4935
   End
   Begin MSComctlLib.ListView LVItem 
      Height          =   2970
      Left            =   1305
      TabIndex        =   2
      Top             =   2040
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5239
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
      NumItems        =   16
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SrNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Receive"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "RcvDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Receive By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Issue"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Iss Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Issue By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Iss Rcv By"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Return"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "RtnDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Return By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Return Receive By"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Dad"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "DadDate"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   15
         Text            =   "Dad By"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSComctlLib.ListView LVTotal 
      Height          =   2370
      Left            =   1852
      TabIndex        =   5
      Top             =   5400
      Width           =   8560
      _ExtentX        =   15108
      _ExtentY        =   4180
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SrNo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Opn Stock"
         Object.Width           =   2540
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
         Text            =   "Issue"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Return"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Dad"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Balance"
         Object.Width           =   1940
      EndProperty
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Size :-"
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
      Left            =   6840
      TabIndex        =   7
      Top             =   1200
      Width           =   960
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Total  :-"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Image Image14 
      Height          =   1020
      Left            =   1320
      Picture         =   "frmRptIW.frx":0000
      Top             =   240
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Items Inventory Report"
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
      Left            =   3495
      TabIndex        =   4
      Top             =   360
      Width           =   4860
   End
   Begin VB.Image ImgClose 
      Height          =   540
      Left            =   10320
      Picture         =   "frmRptIW.frx":0A61
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   540
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :-"
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
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   1440
   End
End
Attribute VB_Name = "frmRptIW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub cmbIName_LostFocus()
On Error Resume Next
'If cmbIName = "" Then
'    MsgBox "Please select valid Items ", vbCritical, Me.Caption
'    cmbIName.SetFocus
'    Exit Sub
'Else
'    cmbSize.SetFocus
'End If
cmbSize.Clear
cmbSize.AddItem "All"
Dim rstmp As New ADODB.Recordset
    rstmp.Open "Select ISize from Items where Items.IName ='" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
        Do While Not rstmp.EOF
            cmbSize.AddItem rstmp!ISize
            rstmp.MoveNext
        Loop
    rstmp.Close
Set rstmp = Nothing

End Sub
Public Function AllSize_Focus()
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim i, a, b, c, d As Integer
LVItem.ListItems.Clear
LVTotal.ListItems.Clear

i = 1
    rs.Open "Select * from Items where Items.IName= '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            LVItem.ListItems.Add i, , rs!SrNo
            LVItem.ListItems(i).SubItems(1) = rs!ISize
            LVTotal.ListItems.Add i, , rs!SrNo
            LVTotal.ListItems(i).SubItems(1) = rs!ISize
            LVTotal.ListItems(i).SubItems(2) = rs!OpnStock
            
            a = i
            rs1.Open "Select * from Receive where Receive.RcvItems= '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs1.EOF
                If a = 1 Then
                    LVItem.ListItems(a).SubItems(2) = rs1!Receive
                    LVItem.ListItems(a).SubItems(3) = rs1!RcvDate
                    LVItem.ListItems(a).SubItems(4) = rs1!RcvBy
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Receive
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Receive
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & a
                    LVItem.ListItems(a).SubItems(1) = rs!ISize
                    LVTotal.ListItems(a).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(a).SubItems(2) = rs1!Receive
                    LVItem.ListItems(a).SubItems(3) = rs1!RcvDate
                    LVItem.ListItems(a).SubItems(4) = rs1!RcvBy
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Receive
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Receive
                    End If
                End If
                a = a + 1
                rs1.MoveNext
                Loop
            b = i
            rs2.Open "Select * from Issue where Issue.IssItems= '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs2.EOF
                If b = 1 Then
                    LVItem.ListItems(b).SubItems(5) = rs2!Issue
                    LVItem.ListItems(b).SubItems(6) = rs2!IssDate
                    LVItem.ListItems(b).SubItems(7) = rs2!Issueby
                    LVItem.ListItems(b).SubItems(8) = rs2!IReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Issue
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & b
                    LVItem.ListItems(b).SubItems(1) = rs!ISize
                    LVTotal.ListItems(b).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(b).SubItems(5) = rs2!Issue
                    LVItem.ListItems(b).SubItems(6) = rs2!IssDate
                    LVItem.ListItems(b).SubItems(7) = rs2!Issueby
                    LVItem.ListItems(b).SubItems(8) = rs2!IReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Issue
                    End If
                End If
                b = b + 1
                rs2.MoveNext
                Loop
            c = i
            rs3.Open "Select * from Return where Return.RtnItems= '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs3.EOF
                If c = 1 Then
                    LVItem.ListItems(c).SubItems(9) = rs3!Return
                    LVItem.ListItems(c).SubItems(10) = rs3!RtnDate
                    LVItem.ListItems(c).SubItems(11) = rs3!Returnby
                    LVItem.ListItems(c).SubItems(12) = rs3!RReceiveby
                    If LVTotal.ListItems(i).SubItems(5) = "" Then
                        LVTotal.ListItems(i).SubItems(5) = rs3!Return
                    Else
                        LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(5) + rs3!Return
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & c
                    LVItem.ListItems(c).SubItems(1) = rs!ISize
                    LVTotal.ListItems(c).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(c).SubItems(9) = rs3!Return
                    LVItem.ListItems(c).SubItems(10) = rs3!RtnDate
                    LVItem.ListItems(c).SubItems(11) = rs3!Returnby
                    LVItem.ListItems(c).SubItems(12) = rs3!RReceiveby
                    If LVTotal.ListItems(i).SubItems(5) = "" Then
                        LVTotal.ListItems(i).SubItems(5) = rs3!Return
                    Else
                        LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(5) + rs3!Return
                    End If
                End If
                c = c + 1
                rs3.MoveNext
                Loop
            d = i
            rs4.Open "Select * from Dad where Dad.DadItems= '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs4.EOF
                If d = 1 Then
                    LVItem.ListItems(d).SubItems(13) = rs4!Dad
                    LVItem.ListItems(d).SubItems(14) = rs4!DadDate
                    LVItem.ListItems(d).SubItems(15) = rs4!Dadby
                    If LVTotal.ListItems(i).SubItems(6) = "" Then
                        LVTotal.ListItems(i).SubItems(6) = rs4!Dad
                    Else
                        LVTotal.ListItems(i).SubItems(6) = LVTotal.ListItems(i).SubItems(6) + rs4!Dad
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & d
                    LVItem.ListItems(d).SubItems(1) = rs!ISize
                    LVTotal.ListItems(d).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(d).SubItems(13) = rs4!Dad
                    LVItem.ListItems(d).SubItems(14) = rs4!DadDate
                    LVItem.ListItems(d).SubItems(15) = rs4!Dadby
                    If LVTotal.ListItems(i).SubItems(6) = "" Then
                        LVTotal.ListItems(i).SubItems(6) = rs4!Dad
                    Else
                        LVTotal.ListItems(i).SubItems(6) = LVTotal.ListItems(i).SubItems(6) + rs4!Dad
                    End If
                End If
                d = d + 1
                rs4.MoveNext
                Loop
                LVTotal.ListItems(i).SubItems(7) = Val(LVTotal.ListItems(i).SubItems(2)) + Val(LVTotal.ListItems(i).SubItems(3)) - Val(LVTotal.ListItems(i).SubItems(4)) + Val(LVTotal.ListItems(i).SubItems(5)) - Val(LVTotal.ListItems(i).SubItems(6))
        i = i + 1
        rs.MoveNext
    Loop
                
                
                    
rs.Close
rs1.Close
rs2.Close
rs3.Close
rs4.Close
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
End Function

Public Function OtherSize_Focus()
On Error Resume Next
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim i, a, b, c, d As Integer
LVItem.ListItems.Clear
LVTotal.ListItems.Clear

i = 1
    rs.Open "Select * from Items where Items.IName= '" & UCase(cmbIName) & "' and Items.ISize= '" & cmbSize & "'", con, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            LVItem.ListItems.Add i, , rs!SrNo
            LVItem.ListItems(i).SubItems(1) = rs!ISize
            LVTotal.ListItems.Add i, , rs!SrNo
            LVTotal.ListItems(i).SubItems(1) = rs!ISize
            LVTotal.ListItems(i).SubItems(2) = rs!OpnStock
            
            a = i
            rs1.Open "Select * from Receive where Receive.RcvItems= '" & UCase(cmbIName) & "' and Receive.RcvSize= '" & cmbSize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs1.EOF
                If a = 1 Then
                    LVItem.ListItems(a).SubItems(2) = rs1!Receive
                    LVItem.ListItems(a).SubItems(3) = rs1!RcvDate
                    LVItem.ListItems(a).SubItems(4) = rs1!RcvBy
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Receive
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Receive
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & a
                    LVItem.ListItems(a).SubItems(1) = rs!ISize
                    LVTotal.ListItems(a).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(a).SubItems(2) = rs1!Receive
                    LVItem.ListItems(a).SubItems(3) = rs1!RcvDate
                    LVItem.ListItems(a).SubItems(4) = rs1!RcvBy
                    If LVTotal.ListItems(i).SubItems(3) = "" Then
                        LVTotal.ListItems(i).SubItems(3) = rs1!Receive
                    Else
                        LVTotal.ListItems(i).SubItems(3) = LVTotal.ListItems(i).SubItems(3) + rs1!Receive
                    End If
                End If
                a = a + 1
                rs1.MoveNext
                Loop
            b = i
            rs2.Open "Select * from Issue where Issue.IssItems= '" & UCase(cmbIName) & "' and Issue.IssSize= '" & cmbSize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs2.EOF
                If b = 1 Then
                    LVItem.ListItems(b).SubItems(5) = rs2!Issue
                    LVItem.ListItems(b).SubItems(6) = rs2!IssDate
                    LVItem.ListItems(b).SubItems(7) = rs2!Issby
                    LVItem.ListItems(b).SubItems(8) = rs2!IReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Issue
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & b
                    LVItem.ListItems(b).SubItems(1) = rs!ISize
                    LVTotal.ListItems(b).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(b).SubItems(5) = rs2!Issue
                    LVItem.ListItems(b).SubItems(6) = rs2!IssDate
                    LVItem.ListItems(b).SubItems(7) = rs2!Issby
                    LVItem.ListItems(b).SubItems(8) = rs2!IReceiveby
                    If LVTotal.ListItems(i).SubItems(4) = "" Then
                        LVTotal.ListItems(i).SubItems(4) = rs2!Issue
                    Else
                        LVTotal.ListItems(i).SubItems(4) = LVTotal.ListItems(i).SubItems(4) + rs2!Issue
                    End If
                End If
                b = b + 1
                rs2.MoveNext
                Loop
            c = i
            rs3.Open "Select * from Return where Return.RtnItems= '" & UCase(cmbIName) & "' and Return.RtnSize= '" & cmbSize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs3.EOF
                If c = 1 Then
                    LVItem.ListItems(c).SubItems(9) = rs3!Return
                    LVItem.ListItems(c).SubItems(10) = rs3!RtnDate
                    LVItem.ListItems(c).SubItems(11) = rs3!Returnby
                    LVItem.ListItems(c).SubItems(12) = rs3!RReceiveby
                    If LVTotal.ListItems(i).SubItems(5) = "" Then
                        LVTotal.ListItems(i).SubItems(5) = rs3!Return
                    Else
                        LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(5) + rs3!Return
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & c
                    LVItem.ListItems(c).SubItems(1) = rs!ISize
                    LVTotal.ListItems(c).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(c).SubItems(9) = rs3!Return
                    LVItem.ListItems(c).SubItems(10) = rs3!RtnDate
                    LVItem.ListItems(c).SubItems(11) = rs3!Returnby
                    LVItem.ListItems(c).SubItems(12) = rs3!RReceiveby
                    If LVTotal.ListItems(i).SubItems(5) = "" Then
                        LVTotal.ListItems(i).SubItems(5) = rs3!Return
                    Else
                        LVTotal.ListItems(i).SubItems(5) = LVTotal.ListItems(i).SubItems(5) + rs3!Return
                    End If
                End If
                c = c + 1
                rs3.MoveNext
                Loop
            d = i
            rs4.Open "Select * from Dad where Dad.DadItems= '" & UCase(cmbIName) & "' and Dad.DadSize= '" & cmbSize & "'", con, adOpenDynamic, adLockOptimistic
                Do While Not rs4.EOF
                If d = 1 Then
                    LVItem.ListItems(d).SubItems(13) = rs4!Dad
                    LVItem.ListItems(d).SubItems(14) = rs4!DadDate
                    LVItem.ListItems(d).SubItems(15) = rs4!Dadby
                    If LVTotal.ListItems(i).SubItems(6) = "" Then
                        LVTotal.ListItems(i).SubItems(6) = rs4!Dad
                    Else
                        LVTotal.ListItems(i).SubItems(6) = LVTotal.ListItems(i).SubItems(6) + rs4!Dad
                    End If
                Else
                    LVItem.ListItems.Add , , rs!SrNo & "." & d
                    LVItem.ListItems(d).SubItems(1) = rs!ISize
                    LVTotal.ListItems(d).SubItems(2) = rs!OpnStock
                    LVItem.ListItems(d).SubItems(13) = rs4!Dad
                    LVItem.ListItems(d).SubItems(14) = rs4!DadDate
                    LVItem.ListItems(d).SubItems(15) = rs4!Dadby
                    If LVTotal.ListItems(i).SubItems(6) = "" Then
                        LVTotal.ListItems(i).SubItems(6) = rs4!Dad
                    Else
                        LVTotal.ListItems(i).SubItems(6) = LVTotal.ListItems(i).SubItems(6) + rs4!Dad
                    End If
                End If
                d = d + 1
                rs4.MoveNext
                Loop
                LVTotal.ListItems(i).SubItems(7) = Val(LVTotal.ListItems(i).SubItems(2)) + Val(LVTotal.ListItems(i).SubItems(3)) - Val(LVTotal.ListItems(i).SubItems(4)) + Val(LVTotal.ListItems(i).SubItems(5)) - Val(LVTotal.ListItems(i).SubItems(6))
        i = i + 1
        rs.MoveNext
    Loop
                
                
                    
rs.Close
rs1.Close
rs2.Close
rs3.Close
rs4.Close
Set rs = Nothing
Set rs1 = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
Set rs4 = Nothing
End Function

Private Sub cmbSize_Change()
On Error Resume Next
    cmbSize = UCase(cmbSize)
    SendKeys "{End}"
End Sub
Private Sub cmbSize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbSize, cmbSize.Text)
End Sub

Private Sub cmbSize_LostFocus()
On Error Resume Next
    If cmbIName = "" Or cmbSize = "" Then
        MsgBox "Please select valid Item Size. ", vbCritical, Me.Caption
        'cmbSize.SetFocus
        cmbIName.SetFocus
        Exit Sub
    End If
    
    If cmbSize = "All" Then
        Call AllSize_Focus
    Else
        Call OtherSize_Focus
    End If
    
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
Private Sub cmbIName_Change()
On Error Resume Next
    cmbIName = UCase(cmbIName)
    SendKeys "{End}"
End Sub
Private Sub cmbIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbIName, cmbIName.Text)
End Sub
Public Function ClearAll()
On Error Resume Next
    FeedData "Items", "IName", cmbIName
    LVItem.ListItems.Clear
    cmbIName.SetFocus
End Function

Private Sub ImgClose_Click()
On Error Resume Next
    Unload Me
End Sub
