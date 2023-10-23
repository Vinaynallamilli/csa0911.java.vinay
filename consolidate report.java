Imports Microsoft.Office.Interop
Public Class consli
 Dim con As New ADODB.Connection
 Dim rs, rs1 As New ADODB.Recordset
 Dim str, dat As String
 Dim i, j, k, diff, count1 As New Integer
 Dim pre_hours(100), tot_hours(100), ab_hours(100) As Integer
 Dim tot_day(100), pre_day(100), ab_day(100) As Double
 Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles Button3.Click
 DataGridView1.Rows.Clear()
 rollno()
 daycalc()
 End Sub
 Private Sub rollno()
 DataGridView1.Rows.Clear()
 Try
 str = "select * from " & ComboBox1.SelectedItem & "_" &
ComboBox5.SelectedItem & ""
 rs.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockOptimistic)
 rs.MoveFirst()
 i = 0
 While (rs.EOF <> True)
 Dim row As String() = New String() {rs.Fields("rollno").Value,
rs.Fields("name").Value}
 DataGridView1.Rows.Add(row)
 DataGridView1.Rows(i).HeaderCell.Value = (i + 1).ToString
 rs.MoveNext()
 i = i + 1
 End While
 rs.Close()
 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try
 End Sub
 Private Sub adconsoli_Load(ByVal sender As System.Object, ByVal e As
System.EventArgs) Handles MyBase.Load
 con = New ADODB.Connection
 rs = New ADODB.Recordset
 con.Open("driver={microsoft ODBC for
Oracle};server=test;uid=M11MCA20;pwd=M11MCA20;")
 Label8.Text = login.temp1
 Label11.Text = login.temp2
 Label10.Text = login.temp3
 End Sub
 Private Sub daycalc()
 Dim pre_hours(100), tot_hours(100), ab_hours(100) As Integer
 Try
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
& "' and subject='" & Label10.Text & "')order by hour asc "
 rs.Open(str, con, ADODB.CursorTypeEnum.adOpenDynamic,
ADODB.LockTypeEnum.adLockPessimistic)
 rs.MoveFirst()
 Dim temp(100), temp1(100) As Integer
 Dim flag1(100) As Integer
 While (rs.EOF <> True)
 Dim rcount As Integer = 0
 Dim count As Integer = 3
 k = 0
 While (rs.Fields.Count > count)
 If String.Equals(rs.Fields(count).Value, "P") Or
String.Equals(rs.Fields(count).Value, "O") Then
 pre_hours(k) = pre_hours(k) + 1
 ElseIf String.Equals(rs.Fields(count).Value, "A") Then
 ab_hours(k) = ab_hours(k) + 1
 End If
 tot_hours(k) = tot_hours(k) + 1

 rcount = rcount + 1
 count = count + 1
 k = k + 1
 End While
 j = j + 1
 rs.MoveNext()
 End While
 For Me.i = 0 To DataGridView1.Rows.Count - 1
 tot_day(i) = tot_day(i) + 1
 Next
 d = DateAdd(DateInterval.Day, 1, d)
 diff = diff - 1
 rs.Close()
 Catch ex As Exception
 d = DateAdd(DateInterval.Day, 1, d)
 diff = diff - 1
 rs.Close()
 End Try
 End While

 Catch ex As Exception
 MsgBox(ex.ToString)
 End Try

 i = 0
 While (i < DataGridView1.Rows.Count - 1)
 DataGridView1.Rows(i).Cells(2).Value = pre_hours(i)
 DataGridView1.Rows(i).Cells(3).Value = ab_hours(i)
 DataGridView1.Rows(i).Cells(4).Value = tot_hours(i)
 DataGridView1.Rows(i).Cells(5).Value = Math.Round((pre_hours(i) /
tot_hours(i) * 100), 2)
 i = i + 1
 End While
 End Sub
 Private Sub HOMEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal
e As System.EventArgs) Handles HOMEToolStripMenuItem.Click
 MDIParent1.Show()
 Me.Close()
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
 For Each col As DataGridViewColumn In DataGridView1.Columns
 xlWorkSheet.Cells(6, col.Index + 1) = col.HeaderText.ToString
 Next
 For i = 1 To DataGridView1.Rows.Count - 1
 For j = 0 To DataGridView1.ColumnCount - 1
 Dim vv As String
 If DataGridView1(j, i - 1).Value Is Nothing Then
 vv = "Niet ingevuld"
 Else
 vv = DataGridView1(j, i - 1).Value.ToString
 xlWorkSheet.Cells(i + 6, j + 1) = vv
 End If
 Next
 ProgressBar1.Value = (i / DataGridView1.Rows.Count) * 100
 Next
 xlWorkBook.Activate()
 xlWorkBook.SaveAs("D:\Consolidate.xls")
 xlWorkBook.Close()
 xlApp.Quit()
 Panel1.Visible = False
 MsgBox("You can find your report at " & "D:\Consolidate.xls")
 End Sub
End Class 
