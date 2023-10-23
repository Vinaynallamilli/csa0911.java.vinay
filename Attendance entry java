Public Class attentry
 Dim con As New ADODB.Connection
 Dim rs, rs1 As New ADODB.Recordset
 Dim str, dat As String
 Dim att As String
 Dim i As Integer = 1
 Dim flag As Integer = 1
 Dim chk1 As New DataGridViewCheckBoxColumn()
 Dim chk As New DataGridViewCheckBoxColumn()
 Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button3.Click
 rs = New ADODB.Recordset
 Try
 str = "select * from " & ComboBox1.SelectedItem & "_" &
ComboBox5.SelectedItem & ""
 rs.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockOptimistic)
 rs.MoveFirst()
 DataGridView1.Rows.Clear()
 i = 1
 While (rs.EOF <> True)
 Dim row As String() = New String() {i, rs.Fields("rollno").Value,
rs.Fields("name").Value}
 DataGridView1.Rows.Add(row)
 i = i + 1
 rs.MoveNext()
 End While
 rs.Close()
 DataGridView1.Columns.Add(chk)
 chk.HeaderText = "PRESENT/ABSENT"
 chk.Name = "chk"
 chk.Selected = True
 DataGridView1.Columns.Add(chk1)
 chk1.HeaderText = "ONDUTY"
 chk1.Name = "chk1"
 timetb()
 Catch ex As Exception
'rs.Close()
 MsgBox(ex.ToString)
 End Try
 End Sub
 Private Sub Form3_Load(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles MyBase.Load
 con = New ADODB.Connection
 'If (con.State = ConnectionState.Open) Then
 ' con.Close()
 'End If
 con.Open("driver={microsoft ODBC for
Oracle};server=test;uid=M11MCA20;pwd=M11MCA20;")
 Label15.Text = login.temp1
 Label16.Text = login.temp2
 Label7.Text = login.temp3
 End Sub
 Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button2.Click
 str = String.Empty
 att = ""
 flag = 1
 dat = DateTimePicker1.Value.Date.ToString("dd-MMM-yyyy")

 For Me.i = 0 To DataGridView1.RowCount - 1
 If DataGridView1.Rows(i).Cells(3).Value = True Then
 If (flag < 2) Then
 att = "'P'"
 flag = 3
 Else
 att = att + ",'P'"
 End If
 ElseIf DataGridView1.Rows(i).Cells(4).Value = True Then
 If (flag < 2) Then
 att = "'O'"
 flag = 3
 Else
att = att + ",'O'"
 End If
 Else
 If (flag < 2) Then
 att = "'A'"
 flag = 3
 Else
 att = att + ",'A'"
 End If
 End If
 Next
 Try
 str = "insert into " & ComboBox1.SelectedItem & "_" &
ComboBox5.SelectedItem & "_" & ComboBox2.SelectedItem & "_" &
ComboBox3.SelectedItem & "_att values('" & dat & "'," & ComboBox4.Text & ",'" &
Label7.Text & "'," & att & ")"
 con.Execute(str)
 MsgBox("insert")
 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try
 End Sub
 Private Sub CREATEToolStripMenuItem_Click(ByVal sender As System.Object,
ByVal e As System.EventArgs) Handles CREATEToolStripMenuItem.Click
 rs1 = New ADODB.Recordset
 str = "select * from " & ComboBox1.Text & "_" & ComboBox5.Text & ""
 rs1.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockPessimistic)
 rs1.MoveFirst()
 str = "create table " & ComboBox1.Text & "_" & ComboBox5.Text & "_" &
ComboBox2.Text & "_" & ComboBox3.Text & "_att(days Date,hour number,subject
varchar(15),primary key(days,hour))"
 con.Execute(str)
 While (rs1.EOF <> True)
 str = "alter table " & ComboBox1.Text & "_" & ComboBox5.Text & "_" &
ComboBox2.Text & "_" & ComboBox3.Text & "_att add(M" &
rs1.Fields("rollno").Value & " varchar(20))"
 con.Execute(str)
 rs1.MoveNext()
 End While
End Sub
 Private Sub timetb()
 Dim temp As String
 rs1 = New ADODB.Recordset
 ComboBox4.Text = "Select One"
 Try
 temp = "select * from " & ComboBox1.Text & "_" & ComboBox5.Text &
"_" & ComboBox2.Text & "_" & ComboBox3.Text & "_time where(day='" &
DateTimePicker1.Value.ToString("dddd") & "')"
 rs1.Open(temp, con, ADODB.CursorTypeEnum.adOpenUnspecified,
ADODB.LockTypeEnum.adLockPessimistic)
 ComboBox4.Items.Clear()
 For Me.i = 1 To 7
 If String.Equals(rs1.Fields(i).Value, Label7.Text) Then
 ComboBox4.Items.Add(i)
 End If
 Next
 rs1.Close()
 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try
 End Sub
 Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal
e As System.EventArgs) Handles DateTimePicker1.ValueChanged
 timetb()
 End Sub
 Private Sub DELETEToolStripMenuItem_Click(ByVal sender As System.Object,
ByVal e As System.EventArgs) Handles DELETEToolStripMenuItem.Click
 str = "drop table " & ComboBox1.Text & "_" & ComboBox5.Text & "_" &
ComboBox2.Text & "_" & ComboBox3.Text & "_" & Label7.Text & " "
 con.Execute(str)
 MsgBox("TABLE DELETED SUCCESSFULLY")
 End Sub
 Private Sub HOMEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal
e As System.EventArgs) Handles HOMEToolStripMenuItem.Click
 MDIParent1.Show()
 Me.Close()
End Sub

 Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs)
 End Sub
 Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e
As System.EventArgs) Handles CheckBox1.CheckedChanged
 If CheckBox1.Checked = True Then
 i = 0
 While (i < DataGridView1.Rows.Count)
 DataGridView1.Rows(i).Cells(3).Value = True
 i = i + 1
 End While
 Else
 i = 0
 While (i < DataGridView1.Rows.Count)
 DataGridView1.Rows(i).Cells(3).Value = False
 i = i + 1
 End While
 End If
 End Sub
 Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object,
ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles
DataGridView1.CellContentClick
 i = 0
 While (i < DataGridView1.Rows.Count)
 If DataGridView1.Rows(i).Cells(3).Value <> True Then
 DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.Red
 Else
 DataGridView1.Rows(i).Cells(3).Style.BackColor = Color.White
 End If
 i = i + 1
 End While
 End Sub
End Class
