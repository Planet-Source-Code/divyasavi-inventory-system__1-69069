VERSION 5.00
Begin VB.Form frmItems 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Items Master"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
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
      Left            =   5880
      TabIndex        =   11
      Text            =   "0"
      Top             =   2640
      Width           =   1215
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
      Left            =   5880
      TabIndex        =   1
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000009&
      Caption         =   "&Close"
      Height          =   495
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H80000009&
      Caption         =   "&Delete"
      Height          =   495
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000009&
      Caption         =   "&Save"
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      BackColor       =   &H80000009&
      Caption         =   "&New"
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
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
      Left            =   5880
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtItems 
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
      Left            =   5880
      TabIndex        =   2
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image Image10 
      Height          =   1020
      Left            =   2880
      Picture         =   "frmItems.frx":0000
      Top             =   1080
      Width           =   1020
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NEW ITEMS ENTRY"
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
      TabIndex        =   13
      Top             =   1320
      Width           =   5055
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
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
      Left            =   4440
      TabIndex        =   12
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Open Stock :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Size :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmbIName_Change()
On Error Resume Next
    cmbIName = UCase(cmbIName)
    SendKeys "{End}"
End Sub

Private Sub cmbIName_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If checkCharacter(KeyCode) Then Call findString(cmbIName, cmbIName.Text)
End Sub

Private Sub cmbISize_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
   If checkCharacter(KeyCode) Then Call findString(cmbISize, cmbISize.Text)
End Sub

Private Sub cmbISize_LostFocus()
On Error Resume Next
    Dim rsf As New ADODB.Recordset
    rsf.Open "Select ISize, OpnStock from Items where Items.IName = '" & UCase(cmbIName) & "'", con, adOpenDynamic, adLockOptimistic
    If rsf.BOF = True And rsf.EOF = True Then
        'cmbISize = ""
        txtItems = ""
        'cmbISize.SetFocus
        txtItems.SetFocus
    Else
        cmbISize = rsf!ISize
        txtItems = rsf!OpnStock
        cmdSave.Caption = "&Update"
        cmdSave.SetFocus
    End If
    rsf.Close
    Set rsf = Nothing
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    con.Execute "Delete * from Items Where Items.SrNo = " & Val(txtSrNo)
    MsgBox "Information is Deleted", vbInformation, Me.Caption
    Call ClearAll
    
End Sub

Private Sub cmdSave_Click()
On Error Resume Next
    If cmbIName.Text = "" Then
        MsgBox " Please Enter Item Name ", vbCritical, Me.Caption
        cmbIName.SetFocus
        Exit Sub
    End If
    If cmbISize.Text = "" Then
        MsgBox " Please Enter Item Size ", vbCritical, Me.Caption
        cmbISize.SetFocus
        Exit Sub
    End If
    If txtItems.Text = "" Then
        txtItems.Text = 0
    End If
    
    With rs
        .Open "Select * from Items where IName= '" & UCase(cmbIName) & "' and Items.ISize = '" & cmbISize & "'", con, adOpenDynamic, adLockOptimistic
        If .EOF = True And .BOF = True Then
            .Close
            .Open "Select * from Items", con, adOpenDynamic, adLockOptimistic
            .AddNew
            !SrNo = GetNewNo("Items")
            !IName = UCase(cmbIName)
            !ISize = cmbISize
            !OpnStock = txtItems
            .Update
            .Close
            MsgBox "Information is Saved ", vbInformation, Me.Caption
        Else
            !SrNo = txtSrNo
            !IName = UCase(cmbIName)
            !ISize = cmbISize
            !OpnStock = txtItems
            .Update
            .Close
            MsgBox "Information is Updated ", vbInformation, Me.Caption
        End If
    End With
    Set rs = Nothing
Call ClearAll
    
End Sub

Private Sub cmdFind_Click()
On Error Resume Next
    'frmSearch.Caption = "Items"
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

FeedData "Items", "IName", cmbIName
FeedData "Items", "ISize", cmbISize
txtSrNo = GetNewNo("Items")
cmbIName.Text = ""
cmbISize.Text = ""
txtItems.Text = ""
cmdSave.Caption = "&Save"
cmbIName.SetFocus
End Function
