Attribute VB_Name = "Module1"
Option Explicit
Global adoConn As New ADODB.Connection
Global GConn As New ADODB.Connection
Global adoRS As New ADODB.Recordset
Global adoCmd As New ADODB.Command
Global strConn As String
Global strSQL As String

Public Const SC_CLOSE = &HF060
Public Const MF_BYCOMMAND = &H0

Public Declare Function GetSystemMenu Lib "user32" _
(ByVal hwnd As Long, ByVal bRevert As Long) As Long

Public Declare Function DeleteMenu Lib "user32" _
(ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Sub Main()
    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
    strConn = strConn & "Data\BRData.mdb"
    
    GConn.ConnectionString = strConn
    GConn.Open
    frmSplash.Show
End Sub
Public Sub FillCombo(Cmb As ComboBox, str1 As String)
    Dim rsF As New ADODB.Recordset
    
    If rsF.State = 1 Then rsF.Close
    rsF.Open str1, GConn, adOpenKeyset, adLockReadOnly
    With Cmb
    .Clear
    While rsF.EOF = False
        .AddItem IIf(IsNull(rsF(1)), "", rsF(1))
        .ItemData(.NewIndex) = IIf(IsNull(rsF(0)), 0, rsF(0))
        rsF.MoveNext
    Wend
    End With
End Sub

Public Sub gFormCenter(frm As Form)
    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 4
End Sub

Public Function OnlyNumber(ByVal MyTextBox As Control)
If Not IsNumeric(MyTextBox.Text) Then
MyTextBox.Text = ""
ElseIf IsNumeric(MyTextBox.Text) Then
If Val(MyTextBox.Text) < 0 Then
MyTextBox.Text = ""
End If
End If
End Function

Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As ADODB.Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean)
Dim X As Variant '|Optional to be declare as variant|
Dim i As Byte
On Error Resume Next
sRecordSource.MoveFirst
sListView.ListItems.Clear
Do While Not sRecordSource.EOF
    If with_num = True Then
        Set X = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
    Else
        Set X = sListView.ListItems.Add(, , sRecordSource.Fields(0), sNumIco, sNumIco)
    End If
        For i = 1 To sNumOfFields - 1
            If Not sRecordSource.Fields(Val(i)) = "" Then
                If show_first_rec = True Then
                    X.SubItems(i) = sRecordSource.Fields(Val(i) - 1)
                Else
                    X.SubItems(i) = sRecordSource.Fields(Val(i))
                End If
            End If
        Next i
    sRecordSource.MoveNext
Loop
i = 0
Set X = Nothing
End Sub




