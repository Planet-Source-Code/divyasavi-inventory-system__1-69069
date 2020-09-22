VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDad 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Dad Items"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbDadIName 
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
      Left            =   4440
      TabIndex        =   0
      Top             =   3000
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
      Left            =   4440
      TabIndex        =   1
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox txtDadIQty 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   3960
      Width           =   3855
   End
   Begin VB.ComboBox cmbDadIBy 
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
      Left            =   4440
      TabIndex        =   3
      Top             =   4440
      Width           =   3855
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
      Left            =   4440
      TabIndex        =   10
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
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
   Begin MSComCtl2.DTPicker DadDate 
      Height          =   330
      Left            =   4440
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39274
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   3000
      Picture         =   "frmDad.frx":0000
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DAD STOCK ENTRY"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   1320
      Width           =   5055
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
      Left            =   8760
      TabIndex        =   18
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Return Stock"
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
      Left            =   8640
      TabIndex        =   17
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Item Name :-"
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
      Left            =   2160
      TabIndex        =   16
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Item Qty :-"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   3960
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
      Left            =   3000
      TabIndex        =   14
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad  By :-"
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
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Dad Date :-"
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
      TabIndex        =   12
      Top             =   4920
      Width           =   1695
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
      Left            =   3000
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmDad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbDadIBy_LostFocus()
On Error Resume Next
If cmbDadIBy = "" Then
    Exit Sub
Else
    CheckData "EmpMaster", "EmpName", cmbDadIBy.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Name , " & cmbDadIBy.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbDadIBy.Text = ""
        cmbDadIBy.SetFocus
    'Else
    '    RcvDate.SetFocus
    End If
End If
End Sub

Private Sub cmbDadIName_LostFocus()
On Error Resume Next
If cmbDadIName = "" Then
    Exit Sub
Else
    CheckData "Items", "IName", cmbDadIName.Text
    If HH = "NOT OK" Then
        MsgBox "Please Enter valid Item , " & cmbDadIName.Text & "  is not in Database.", vbCritical, Me.Caption
        cmbDadIName.Text = ""
        cmbDadIName.SetFocus
    'Else
    '    cmbISize.SetFocus
    End If
End If

End Sub

Private Sub cmbISize_Change()
On Error Resume Next
    cmbISize = UCase(cmbISize)
    SendKeys "{End}"
End Sub

Private Sub cmbISize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbISize, cmbISize.Text)
End Sub

Private Sub cmbDadIBy_Change()
On Error Resume Next
    cmbDadIBy = UCase(cmbDadIBy)
    SendKeys "{End}"
End Sub

Private Sub cmbDadIBy_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbDadIBy, cmbDadIBy.Text)
End Sub
Private Sub cmbDadIName_Change()
On Error Resume Next
    cmbDadIName = UCase(cmbDadIName)
    SendKeys "{End}"
End Sub

Private Sub cmbDadIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If checkCharacter(KeyCode) Then Call findString(cmbDadIName, cmbDadIName.Text)
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
    '    txtRcvIQty.SetFocus
    End If
End If
Call RtnStock
If lblCStock = 0 Then
    MsgBox "You can't Dad " & UCase(cmbIssIName) & ", 0 (Zero) Stock Return.", vbCritical, Me.Caption
    Call ClearAll
End If
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub
Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Dad Where Dad.SrNo = " & Val(txtSrNo)
    MsgBox " Information is Deleted ", vbInformation, Me.Caption
    Call ClearAll
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    'frmSearch.Caption = "Dad"
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
If cmbDadIName = "" Then
    MsgBox "Plese Select Dad Item Name.", vbCritical, Me.Caption
    cmbDadIName.SetFocus
    Exit Sub
End If
If cmbISize = "" Then
    MsgBox "Please Select Dad Item Size ", vbCritical, Me.Caption
    cmbISize.SetFocus
    Exit Sub
End If
If txtDadIQty = "" Then
    MsgBox "Please Enter Dad Quantity ", vbCritical, Me.Caption
    txtDadIQty.SetFocus
    Exit Sub
End If
If cmbDadIBy = "" Then
    MsgBox "Please Select Dad By Name ", vbCritical, Me.Caption
    cmbDadIBy.SetFocus
    Exit Sub
End If

With rs
    '.Open "Select * from Dad where DadItems = '" & UCase(cmbDadIName) & "'", con, adOpenDynamic, adLockOptimistic
    .Open "Select * from Dad where SrNo = '" & Val(txtSrNo) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from Dad", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !SrNo = GetNewNo("Dad")
        !DadItems = UCase(cmbDadIName)
        !DadSize = UCase(cmbISize)
        !Dad = txtDadIQty
        !Dadby = UCase(cmbDadIBy)
        !DadDate = DadDate
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !SrNo = txtSrNo
        !DadItems = UCase(cmbDadIName)
        !DadSize = UCase(cmbISize)
        !Dad = txtDadIQty
        !Dadby = UCase(cmbDadIBy)
        !DadDate = DadDate
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

FeedData "Items", "IName", cmbDadIName
FeedData "Items", "ISize", cmbISize
FeedData "EmpMaster", "EmpName", cmbDadIBy

txtSrNo = GetNewNo("Dad")
cmbDadIName.Text = ""
cmbISize.Text = ""
txtDadIQty.Text = ""
cmbDadIBy.Text = ""
DadDate = Date

cmbDadIName.SetFocus
End Function
Public Function RtnStock()
On Error Resume Next
Dim rstmp As New ADODB.Recordset
    rstmp.Open "Select sum(Return) from Return where Return.RtnItems ='" & UCase(cmbDadIName) & "' and Return.RtnSize ='" & UCase(cmbISize) & "'", con, adOpenDynamic, adLockOptimistic
        If IsNull(rstmp(0)) Then
            lblCStock = 0
        Else
            lblCStock = rstmp(0)
        End If
        rstmp.Close
Set rstmp = Nothing
End Function

Private Sub txtDadIQty_LostFocus()
On Error Resume Next
If Val(lblCStock) < Val(txtDadIQty) Then
    MsgBox " Dad Quantity can not greater then Return Stock", vbCritical, Me.Caption
    txtDadIQty = ""
    txtDadIQty.SetFocus
End If
End Sub
