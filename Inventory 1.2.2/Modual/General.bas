Attribute VB_Name = "General"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Dim con As New ADODB.Connection
Public HH As String
Sub main()
On Error Resume Next
    'Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    'Con.Open
  frmMain.Show
  'frmAbout.Show
End Sub

Public Function FeedData(ByVal TName As String, TField As String, com1 As ComboBox)
On Error Resume Next
'' feed data in cmb box
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
    com1.Clear
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    cn.Open

    rs.Open "Select " & TField & " from " & TName & " Group By " & TField, cn, adOpenDynamic, adLockOptimistic
        Do While Not rs.EOF
            com1.AddItem rs.Fields(0)
            rs.MoveNext
        Loop
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Function
Public Function CheckData(ByVal TName As String, TField As String, H As String)
On Error Resume Next
Dim cnh As New ADODB.Connection
Dim rsh As New ADODB.Recordset

    cnh.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    cnh.Open
    rsh.Open "Select " & TField & " From " & TName & " Where " & TName & "." & TField & " ='" & UCase(H) & "'", cnh, adOpenDynamic, adLockOptimistic
    If rsh.EOF = True And rsh.BOF = True Then
        HH = "NOT OK"
    Else
        HH = "OK"
    End If
    rsh.Close
    cnh.Close
Set rsh = Nothing
Set cnh = Nothing
End Function

Public Function findString(loc_Combo As ComboBox, loc_String As String)
On Error Resume Next
    Dim i As Integer
    i = SendMessage(loc_Combo.hWnd, CB_FINDSTRING, -1, ByVal loc_Combo.Text)
    If i >= 0 Then
        Dim varStartSelection As Integer
        If Left(loc_Combo.List(i), Len(loc_String)) = loc_String Then
            loc_Combo.ListIndex = i
            loc_Combo.SelStart = Len(loc_String)
            loc_Combo.SelLength = Len(loc_Combo.Text)
        End If
    End If
End Function

Public Function checkCharacter(ByVal KeyCode As Integer) As Boolean
On Error Resume Next
    Select Case KeyCode
        Case vbKeyBack, vbKeyUp, vbKeyDown, vbKeyRight, vbKeyLeft, vbKeyDelete, vbKeyReturn, vbKeyHome, vbKeyEnd, vbKeyPageDown, vbKeyPageUp, vbKeyCapital, vbKeyShift, vbKeyControl, 144
            checkCharacter = False
        Case Else
            checkCharacter = True
    End Select
End Function

Public Function GetNewNo(ByVal TName As String) As Integer
On Error Resume Next
Dim ECon As New ADODB.Connection
Dim ErsMax As New ADODB.Recordset
    ECon.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\InvData.mdb;Persist Security Info=False"
    ECon.Open
    
    ErsMax.Open "Select MAX (" & TName & ".SrNo) from " & TName, ECon, adOpenDynamic, adLockOptimistic
    If IsNull(ErsMax.Fields(0)) Then
        GetNewNo = 1
    Else
        GetNewNo = Val(ErsMax.Fields(0)) + 1
    End If
ErsMax.Close
Set ErsMax = Nothing
Set ECon = Nothing
End Function

