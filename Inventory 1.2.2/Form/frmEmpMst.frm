VERSION 5.00
Begin VB.Form frmEmpMst 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Employee Master"
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
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5400
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
      Left            =   7440
      TabIndex        =   19
      Text            =   "0"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000009&
      Caption         =   "&New"
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "&Save"
      Height          =   495
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000009&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000009&
      Caption         =   "&Close"
      Height          =   495
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ComboBox cmbDesig 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   3120
      Width           =   5655
   End
   Begin VB.ComboBox cmbCity 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Width           =   3735
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
      Left            =   3960
      TabIndex        =   1
      Top             =   2640
      Width           =   5655
   End
   Begin VB.TextBox txtMobile 
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
      Left            =   7320
      TabIndex        =   6
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtPhone 
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
      Left            =   3960
      TabIndex        =   5
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox txtAdd 
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
      Left            =   3960
      TabIndex        =   3
      Top             =   3600
      Width           =   5655
   End
   Begin VB.TextBox txtEmpCode 
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
      Left            =   3960
      TabIndex        =   0
      Text            =   "JAL"
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   2520
      Picture         =   "frmEmpMst.frx":0000
      Top             =   840
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE  MASTER"
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
      Left            =   3720
      TabIndex        =   21
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Emp.  No :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile :-"
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
      Left            =   6360
      TabIndex        =   18
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Phone :-"
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
      Left            =   2760
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "City :-"
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
      Left            =   2760
      TabIndex        =   16
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Address :-"
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
      Left            =   2760
      TabIndex        =   15
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Desig :-"
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
      Left            =   2760
      TabIndex        =   14
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name :-"
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
      Left            =   1920
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code :-"
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1695
   End
End
Attribute VB_Name = "frmEmpMst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbCity_Change()
On Error Resume Next
    cmbCity = UCase(cmbCity)
    SendKeys "{End}"
End Sub

Private Sub cmbDesig_Change()
On Error Resume Next
    cmbDesig = UCase(cmbDesig)
    SendKeys "{End}"
End Sub

Private Sub cmbDesig_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbDesig, cmbDesig.Text)
End Sub

Private Sub cmbEmpName_Change()
On Error Resume Next
    cmbEmpName = UCase(cmbEmpName)
    SendKeys "{End}"
End Sub
Private Sub cmbCity_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If checkCharacter(KeyCode) Then Call findString(cmbCity, cmbCity.Text)
End Sub
Private Sub cmbEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If checkCharacter(KeyCode) Then Call findString(cmbEmpName, cmbEmpName.Text)
End Sub

Private Sub cmbEmpName_LostFocus()
On Error Resume Next
Dim rsf As New ADODB.Recordset
    rsf.Open "Select * from EmpMaster where EmpMaster.EmpName = '" & UCase(cmbEmpName) & "'", con, adOpenDynamic, adLockOptimistic
    If rsf.BOF = True And rsf.EOF = True Then
        cmbDesig = ""
        txtAdd = ""
        cmbCity = ""
        txtPhone = ""
        txtMobile = ""
    Else
        'txtEmpCode = rsf!EmpCode
        'txtSrNo = rsf!SrNo
        cmbDesig = rsf!Desig
        txtAdd = rsf!Add
        cmbCity = rsf!City
        txtPhone = rsf!Phone
        txtMobile = rsf!Mobile
    End If
    rsf.Close
    Set rsf = Nothing
    
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
    Unload Me
End Sub

Public Function ClearAll()
On Error Resume Next

FeedData "EmpMaster", "EmpName", cmbEmpName
FeedData "EmpMaster", "Desig", cmbDesig
FeedData "EmpMaster", "City", cmbCity
txtEmpCode.Text = "JAL"
txtSrNo = GetNewNo("EmpMaster")
cmbEmpName.Text = ""
cmbDesig.Text = ""
txtAdd.Text = ""
cmbCity.Text = ""
txtPhone.Text = ""
txtMobile.Text = ""

cmbEmpName.SetFocus

End Function

Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from EmpMaster Where EmpMaster.SrNo = " & Val(txtSrNo)
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    'frmSearch.Caption "EmpMaster"
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
If cmbEmpName = "" Then
    MsgBox "Please Enter Employee Name ", vbCritical, Me.Caption
    Exit Sub
End If
If cmbDesig = "" Then
    MsgBox "Please Enter Employee Designation ", vbCritical, Me.Caption
    Exit Sub
End If
If txtPhone = "" And txtMobile = "" Then
    MsgBox "Please Enter Employee's Phone or Mobile ", vbCritical, Me.Caption
    Exit Sub
End If
With rs
    .Open "Select * from EmpMaster where EmpName = '" & UCase(cmbEmpName) & "'", con, adOpenDynamic, adLockOptimistic
    If .EOF = True And .BOF = True Then
        .Close
        .Open "Select * from EmpMaster", con, adOpenDynamic, adLockOptimistic
        .AddNew
        !EmpCode = UCase(txtEmpCode)
        !SrNo = GetNewNo("EmpMaster")
        !EmpName = UCase(cmbEmpName)
        !Desig = UCase(cmbDesig)
        !Add = UCase(txtAdd)
        !City = UCase(cmbCity)
        !Phone = txtPhone
        !Mobile = txtMobile
        .Update
        .Close
        MsgBox "Information is Saved", vbInformation, Me.Caption
    Else
        !EmpCode = UCase(txtEmpCode)
        !SrNo = txtSrNo
        !EmpName = UCase(cmbEmpName)
        !Desig = UCase(cmbDesig)
        !Add = UCase(txtAdd)
        !City = UCase(cmbCity)
        !Phone = txtPhone
        !Mobile = txtMobile
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
