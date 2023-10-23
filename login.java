Imports System.Data
Public Class login
 Dim con As New ADODB.Connection
 Dim rs, rs1 As New ADODB.Recordset
 Public str, temp1, temp2, temp3, temp4 As String
 Dim i As Integer
 Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button1.Click
 rs = New ADODB.Recordset
 rs1 = New ADODB.Recordset
 If String.Equals(TextBox1.Text, "Admin") Or String.Equals(TextBox1.Text,
"admin") Or String.Equals(TextBox1.Text, "ADMIN") And
String.Equals(TextBox2.Text, "Admin") Or String.Equals(TextBox2.Text, "admin") Or
String.Equals(TextBox2.Text, "ADMIN") Then
 temp4 = "MDIParent2"
 TextBox1.Text = ""
 TextBox2.Text = ""
 MDIParent2.Show()
 Me.Hide()
 i = 1
 Else
 Try
 str = "select * from logintable"
 rs.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockPessimistic)
 rs.MoveFirst()
 While (rs.EOF <> True)
 str = "select * from " & rs.Fields("tablename").Value & ""
 rs1.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockPessimistic)
 While (rs1.EOF <> True)
 If String.Equals(rs1.Fields("sname").Value, TextBox1.Text) And
String.Equals(rs1.Fields("pass").Value, TextBox2.Text) Then
 temp1 = rs1.Fields("sname").Value
 temp2 = rs1.Fields("scode").Value
 temp3 = rs1.Fields("ssname").Value
 temp4 = "MDIParent1"
 TextBox1.Text = ""
TextBox2.Text = ""
 MDIParent1.Show()
 Me.Hide()
 i = 1
 Exit While
 End If
 rs1.MoveNext()
 End While
 rs1.Close()
 rs.MoveNext()
 End While
 If i = 0 Then
 MsgBox("LOGIN NOT VAILD")
 End If
 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try
 End If
 End Sub
 Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles MyBase.Load
 con = New ADODB.Connection
 If (con.State = ConnectionState.Open) Then
 con.Close()
 End If
 con.Open("driver={microsoft ODBC for
Oracle};server=test;uid=M11MCA20;pwd=M11MCA20;")
 End Sub
 Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button2.Click
 End
 End Sub
End Class 
