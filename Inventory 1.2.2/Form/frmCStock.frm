VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCStock 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "All Items Current Stocks"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Caption         =   "&Close"
      Height          =   375
      Left            =   9720
      MaskColor       =   &H80000009&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin MSComctlLib.ListView LVCStock 
      Height          =   6375
      Left            =   1050
      TabIndex        =   0
      Top             =   1200
      Width           =   9990
      _ExtentX        =   17621
      _ExtentY        =   11245
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sr.No"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Items Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Open Stock"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Receive"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Issue"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Return"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Dad"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Cur. Stock"
         Object.Width           =   2293
      EndProperty
   End
   Begin VB.Image Image5 
      Height          =   1020
      Left            =   1560
      Picture         =   "frmCStock.frx":0000
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK  INVENTORY"
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
      Left            =   3405
      TabIndex        =   2
      Top             =   360
      Width           =   5295
   End
End
Attribute VB_Name = "frmCStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    Con.Open
Call ClearAll
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    
    Set rs = Nothing
    Set Con = Nothing
End Sub
Public Function CStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
Dim rstmp1 As New ADODB.Recordset
Dim rstmp2 As New ADODB.Recordset
Dim rstmp3 As New ADODB.Recordset
Dim rstmp4 As New ADODB.Recordset

Dim i As Integer
i = 1

    rstmp.Open "select * from Items ", Con, adOpenDynamic, adLockOptimistic
        Do While Not rstmp.EOF
            LVCStock.ListItems.Add i, , rstmp!SrNo
            LVCStock.ListItems(i).SubItems(1) = rstmp!IName
            LVCStock.ListItems(i).SubItems(2) = rstmp!ISize
            LVCStock.ListItems(i).SubItems(3) = rstmp!OpnStock
            
            rstmp1.Open "Select sum(Receive) from Receive where Receive.RcvItems ='" & rstmp!IName & "' and Receive.RcvSize ='" & rstmp!ISize & "'", Con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp1.EOF
                    If IsNull(rstmp1(0)) Then
                        LVCStock.ListItems(i).SubItems(4) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(4) = rstmp1(0)
                    End If
                    rstmp1.MoveNext
                Loop
                If rstmp1.EOF = True Then
                    rstmp1.Close
                End If
                
            rstmp2.Open "Select sum(Issue) from Issue where Issue.IssItems ='" & rstmp!IName & "' and Issue.IssSize ='" & rstmp!ISize & "'", Con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp2.EOF
                    If IsNull(rstmp2(0)) Then
                        LVCStock.ListItems(i).SubItems(5) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(5) = rstmp2(0)
                    End If
                    rstmp2.MoveNext
                Loop
                If rstmp2.EOF = True Then
                    rstmp2.Close
                End If
            
            rstmp3.Open "Select sum(Return) from Return where Return.RtnItems ='" & rstmp!IName & "' and Return.RtnSize ='" & rstmp!ISize & "'", Con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp3.EOF
                    If IsNull(rstmp3(0)) Then
                        LVCStock.ListItems(i).SubItems(6) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(6) = rstmp3(0)
                    End If
                    rstmp3.MoveNext
                Loop
                If rstmp3.EOF = True Then
                    rstmp3.Close
                End If
            
            rstmp4.Open "Select sum(Dad) from Dad where Dad.DadItems ='" & rstmp!IName & "' and Dad.DadSize ='" & rstmp!ISize & "'", Con, adOpenDynamic, adLockOptimistic
                Do While Not rstmp4.EOF
                    If IsNull(rstmp4(0)) Then
                        LVCStock.ListItems(i).SubItems(7) = 0
                    Else
                        LVCStock.ListItems(i).SubItems(7) = rstmp4(0)
                    End If
                    rstmp4.MoveNext
                Loop
                If rstmp4.EOF = True Then
                    rstmp4.Close
                End If
            LVCStock.ListItems(i).SubItems(8) = Val(LVCStock.ListItems(i).SubItems(3)) + Val(LVCStock.ListItems(i).SubItems(4)) - Val(LVCStock.ListItems(i).SubItems(5)) + Val(LVCStock.ListItems(i).SubItems(6)) - Val(LVCStock.ListItems(i).SubItems(7))
            i = i + 1
            rstmp.MoveNext
        Loop
    
    
    
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
    LV.ListItems.Clear
    Call CStock
End Function
