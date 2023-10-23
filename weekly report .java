Imports Microsoft.Office.Interop
Public Class awreport
 Dim conn As New ADODB.Connection
 Dim rs, rs1 As New ADODB.Recordset
 Dim str, dat As String
 Dim i, j, flag, diff, count1 As New Integer
 Dim ro, temp, tot_day, pre_day, ab_day As Integer
 Dim holiday As String
 Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button3.Click
 holiday = String.Empty
 Try
 rs = New ADODB.Recordset
 rs1 = New ADODB.Recordset
 DataGridView1.Rows.Clear()
 DataGridView1.Columns.Clear()
 DataGridView2.Rows.Clear()
 DataGridView2.Columns.Clear()
 Dim clm1 As New DataGridViewTextBoxColumn()
 DataGridView2.Columns.Add(clm1)
 clm1.HeaderText = ComboBox1.Text + "-" + ComboBox5.Text
 clm1.Name = "clm1"
 Dim clm2 As New DataGridViewTextBoxColumn()
 DataGridView2.Columns.Add(clm2)
 clm2.HeaderText = "SEMESTER" + "-" + ComboBox3.Text
 clm2.Name = "clm3"
 DataGridView2.Columns(1).Width = 130
 Dim clm3 As New DataGridViewTextBoxColumn()
 DataGridView1.Columns.Add(clm3)
 clm3.HeaderText = "ROLLNO"
 clm3.Name = "clm3"
 Dim clm4 As New DataGridViewTextBoxColumn()
 DataGridView1.Columns.Add(clm4)
clm4.HeaderText = "STUDENT NAME"
 clm4.Name = "clm4"
 DataGridView1.Columns(1).Width = 130
 str = "select * from " & ComboBox1.SelectedItem & "_" &
ComboBox5.SelectedItem & ""
 rs.Open(str, conn, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockOptimistic)
 rs.MoveFirst()
 While (rs.EOF <> True)
 Dim row As String() = New String() {rs.Fields("rollno").Value,
rs.Fields("name").Value}
 DataGridView1.Rows.Add(row)
 rs.MoveNext()
 End While
 rs.Close()
 Dim d As Date
 d = DateTimePicker1.Value.Date
 Dim d1 As Date
 d1 = DateTimePicker2.Value.Date
 diff = DateDiff(DateInterval.Day, d, d1)
 j = 2
 While diff >= 0
 Try
 str = "Select * from " & ComboBox1.SelectedItem & "_" &
ComboBox5.Text & "_" & ComboBox2.SelectedItem & "_" &
ComboBox3.SelectedItem & "_att where(days='" & d.Date.ToString("dd-MMM-yyyy")
& "')order by hour asc "
 rs1.Open(str, conn, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockPessimistic)
 rs1.MoveFirst()
 count1 = 1
 Dim dtxt As New DataGridViewTextBoxColumn()
 DataGridView2.Columns.Add(dtxt)
 dtxt.HeaderText = d.Date.ToString("dd-MMM-yyyy")
 dtxt.Width = 140
 While (rs1.EOF <> True)
 Dim dtxt1 As New DataGridViewTextBoxColumn()
 DataGridView1.Columns.Add(dtxt1)
 dtxt1.HeaderText = rs1.Fields("hour").Value.ToString 
dtxt1.Width = 20
 Dim rcount As Integer = 0
 Dim count As Integer = 3
 While (rs1.Fields.Count > count)
 DataGridView1.Rows(rcount).Cells(j).Value = rs1.Fields(count).Value
 DataGridView1.Rows(rcount).HeaderCell.Value = (rcount +
1).ToString
 If String.Equals(rs1.Fields(count).Value, "A") Then
 DataGridView1.Rows(rcount).Cells(j).Style.BackColor = Color.Red
 End If
 rcount = rcount + 1
 count = count + 1
 End While
 count1 = count1 + 1
 j = j + 1
 rs1.MoveNext()
 End While
 rs1.Close()
 d = DateAdd(DateInterval.Day, 1, d)
 diff = diff - 1
 Catch ex As Exception
 holiday += "(" + d.Date.ToString("dd-MMM-yyyy") + "-HOLIDAY) "
 d = DateAdd(DateInterval.Day, 1, d)
 diff = diff - 1
 rs1.Close()
 End Try
 End While
 ' MsgBox(holiday)
 DataGridView1.Rows.Add(holiday)
 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try

 End Sub
 Private Sub creport_Load(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles MyBase.Load
 conn = New ADODB.Connection
 rs = New ADODB.Recordse
conn.Open("driver={microsoft ODBC for
Oracle};server=test;uid=M11MCA20;pwd=M11MCA20;")
 End Sub
 Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button1.Click
 Panel1.Visible = True
 ProgressBar1.Minimum = 0
 ProgressBar1.Maximum = 100
 Dim xlApp As Excel.Application
 Dim xlWorkBook As Excel.Workbook
 Dim xlWorkSheet As Excel.Worksheet
 Dim misValue As Object = System.Reflection.Missing.Value
 Dim i As Integer
 Dim j As Integer
 xlApp = New Excel.Application
 xlWorkBook = xlApp.Workbooks.Add(misValue)
 xlWorkSheet = xlWorkBook.Sheets("sheet1")
 flag = 0
 j = 1
 xlWorkSheet.Cells(1, 1) = "Dr.Mahalingam College of Engineering & Technology
".ToString
 xlWorkSheet.Cells(2, 1) = "NPT -MCET Campus, Udumalai Road -
Makkinaickenpatti - Pollachi".ToString
 xlWorkSheet.Cells(3, 1) = "Phone : 04259-236030 Fax : 04259-236070".ToString
 xlWorkSheet.Cells(4, 1) = "E-Mail : principal@drmcet.ac.in Web Site :
www.mcet.in".ToString
 xlWorkSheet.Range("A5").Value = "BATCH:" + ComboBox1.Text + "-" +
ComboBox5.Text + " ATTENDANCE DETAILS FROM " +
DateTimePicker1.Value.ToString("dd-MMM-yyyy") + " TO " +
DateTimePicker2.Value.ToString("dd-MMM-yyyy") + " SEMESTER:" + "-" +
ComboBox3.Text
 For Each col1 As DataGridViewColumn In DataGridView2.Columns
 If flag < 2 Then
xlWorkSheet.Cells(6, col1.Index + 1) = col1.HeaderText.ToString
 flag = flag + 1
 j = j + 1
 Else
 j = j + 1
 xlWorkSheet.Cells(6, j) = col1.HeaderText.ToString
 For i = 1 To 6
 j = j + 1
 xlWorkSheet.Cells(6, j + i - 1) = "".ToString
 Next
 End If
 Next
 xlWorkSheet.Cells(6, 1) = "SNO".ToString
 flag = 0
 For Each col As DataGridViewColumn In DataGridView1.Columns
 If flag < 2 Then
 xlWorkSheet.Cells(6, col.Index + 2) = col.HeaderText.ToString
 flag = flag + 1
 Else
 xlWorkSheet.Cells(7, col.Index + 2) = col.HeaderText.ToString
 End If
 Next
 For i = 1 To DataGridView1.Rows.Count - 1
 xlWorkSheet.Cells(i + 7, 1) = i.ToString
 flag = 0
 For j = 0 To DataGridView1.ColumnCount - 1
 Dim vv As String
 If DataGridView1(j, i - 1).Value Is Nothing Then
 vv = "Niet ingevuld"
 Else
 vv = DataGridView1(j, i - 1).Value.ToString
 xlWorkSheet.Cells(i + 7, j + 2) = vv
 If flag < 2 Then
 xlWorkSheet.Columns(j + 2).ColumnWidth = 15
 'xlWorkSheet.Columns.Merge(2)
 flag = flag + 1
 Else
 xlWorkSheet.Columns(j + 2).ColumnWidth = 1
 End If
 End IfxlWorkSheet.Cells(6, col1.Index + 1) = col1.HeaderText.ToString
 flag = flag + 1
 j = j + 1
 Else
 j = j + 1
 xlWorkSheet.Cells(6, j) = col1.HeaderText.ToString
 For i = 1 To 6
 j = j + 1
 xlWorkSheet.Cells(6, j + i - 1) = "".ToString
 Next
 End If
 Next
 xlWorkSheet.Cells(6, 1) = "SNO".ToString
 flag = 0
 For Each col As DataGridViewColumn In DataGridView1.Columns
 If flag < 2 Then
 xlWorkSheet.Cells(6, col.Index + 2) = col.HeaderText.ToString
 flag = flag + 1
 Else
 xlWorkSheet.Cells(7, col.Index + 2) = col.HeaderText.ToString
 End If
 Next
 For i = 1 To DataGridView1.Rows.Count - 1
 xlWorkSheet.Cells(i + 7, 1) = i.ToString
 flag = 0
 For j = 0 To DataGridView1.ColumnCount - 1
 Dim vv As String
 If DataGridView1(j, i - 1).Value Is Nothing Then
 vv = "Niet ingevuld"
 Else
 vv = DataGridView1(j, i - 1).Value.ToString
 xlWorkSheet.Cells(i + 7, j + 2) = vv
 If flag < 2 Then
 xlWorkSheet.Columns(j + 2).ColumnWidth = 15
 'xlWorkSheet.Columns.Merge(2)
 flag = flag + 1
 Else
 xlWorkSheet.Columns(j + 2).ColumnWidth = 1
 End If
 End If
ProgressBar1.Value = (i / DataGridView1.Rows.Count) * 100
 Next
 Next
 xlWorkSheet.Range("A1:AS1").Merge()
 xlWorkSheet.Range("A2:AS2").Merge()
 xlWorkSheet.Range("A3:AS3").Merge()
 xlWorkSheet.Range("A4:AS4").Merge()
 xlWorkSheet.Range("A5:AS5").Merge()
 xlWorkSheet.Range("D6:J6").Merge()
 xlWorkSheet.Range("K6:Q6").Merge()
 xlWorkSheet.Range("R6:X6").Merge()
 xlWorkSheet.Range("Y6:AE6").Merge()
 xlWorkSheet.Range("AF6:AL6").Merge()
 xlWorkSheet.Range("AM6:AS6").Merge()
 xlWorkBook.Activate()
 xlWorkBook.SaveAs("D:\export.xls")
 xlWorkBook.Close()
 xlApp.Quit()
 Panel1.Visible = False
 MsgBox("You can find your report at " & "D:\export.xls")
 End Sub
 Private Sub HOMEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal
e As System.EventArgs) Handles HOMEToolStripMenuItem.Click
 MDIParent2.Show()
 Me.Close()
 End Sub
End Class 
