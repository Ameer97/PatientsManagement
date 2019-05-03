Imports System.Data.OleDb
Public Class Form1
    Dim flag As Boolean = False
    Dim dt As DataTable = New DataTable()
    Dim view As DataView

    'Private Sub getXlFile()


    '    OpenFileDialog1.Filter = "Excel 2003 Documents (*.xls) | *.xls"

    '    If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
    '        Dim fileName As String
    '        fileName = OpenFileDialog1.FileName
    '        getExcel(fileName)
    '    End If

    'End Sub

    Private Sub getExcel()

        Dim s As String = ""
        Dim connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=""Excel 8.0; HDR = Yes; IMEX = 1""", "MyExcelFile.xls")

        Dim adapter As OleDbDataAdapter
        adapter = New OleDbDataAdapter("select * from [patient$]", connectionString)

        Dim Data = New DataTable()

        Dim ds = New DataSet()
        adapter.Fill(ds, "patient")
        Data = ds.Tables("patient")
        dt = Data
        DataGridView1.DataSource = Encrypter(Data)

    End Sub
    Function Encrypter(DataTableInstance As DataTable)

        If flag Then
            Return DataTableInstance
        End If
        For i = 0 To Convert.ToInt32(DataTableInstance.Rows.Count - 1)
            For j = 0 To Convert.ToInt32(DataTableInstance.Columns.Count - 1)
                DataTableInstance.Rows(i).Item(j) = Encrypt(DataTableInstance.Rows(i).Item(j))
            Next
        Next
        Return DataTableInstance
    End Function

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        'Dim path = "C:\Users\Ameer\Desktop\phpkey.txt"
        'Dim check As IO.StreamReader = New IO.StreamReader(path)
        'Dim checkstring = check.ReadToEnd
        Dim username = TextBox1.Text
        Dim password = TextBox2.Text

        If username = "doctor" AndAlso password = "1234" Then
            flag = True
            MessageBox.Show("You Are a Doctor")
            Button1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            TextBox1.Visible = False
            TextBox2.Visible = False
        Else
            MessageBox.Show("You are not a Doctor")
            flag = False

        End If
        Dim watch As Stopwatch = Stopwatch.StartNew()
        watch.Start()
        'getXlFile()
        getExcel()
        watch.Stop()
        MessageBox.Show(watch.Elapsed.TotalSeconds)

        If flag Then
            Button3.Enabled = True
        End If

    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim dt As DataTable = New DataTable()
        For Each col As DataGridViewColumn In DataGridView1.Columns
            dt.Columns.Add(col.Name)
        Next

        For Each row As DataGridViewRow In DataGridView1.Rows
            Dim dRow As DataRow = dt.NewRow()

            For Each cell As DataGridViewCell In row.Cells
                dRow(cell.ColumnIndex) = cell.Value
            Next

            dt.Rows.Add(dRow)
        Next
        Dim customerTable As DataSet = New DataSet()
        dt.TableName = "patient"
        customerTable.Tables.Add(dt)
        ExcelLibrary.DataSetHelper.CreateWorkbook("MyExcelFile", customerTable)
        MessageBox.Show("Your file has been saved")
    End Sub


    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        Dim query = _
       From order In dt.AsEnumerable() _
       Where order.Field(Of String)(ComboBox1.SelectedItem.ToString).Contains(TextBox3.Text)
       Select order

        view = query.AsDataView()
        DataGridView1.DataSource = view

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button2.Enabled = False
        ComboBox1.Enabled = False
        TextBox3.Enabled = False
        Button3.Enabled = False
        Button4.Enabled = False



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim activate As String
        If flag Then
            activate = InputBox("Enter Activate Code")
        End If
        If activate = "JKHX300590-eyJsaWNlbnNlSWQiOiJKS0hYMzAwNTkwIiwibGljZW5zZWVOYW1lIjoiTmljb2xlIFBvd2VsbCIsImFzc2lnbmVlTmFtZSI6IiIsImFzc2lnbmVlRW1haWwiOiIiLCJsaWNlbnNlUmVzdHJpY3Rpb24iOiJGb3IgZWR1Y2F0aW9uYWwgdXNlIG9ubHkiLCJjaGVja0NvbmN1cnJlbnRVc2UiOmZhbHNlLCJwcm9kdWN0cyI6W3siY29kZSI6IklJIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In0seyJjb2RlIjoiQUMiLCJwYWlkVXBUbyI6IjIwMTktMDktMjQifSx7ImNvZGUiOiJEUE4iLCJwYWlkVXBUbyI6IjIwMTktMDktMjQifSx7ImNvZGUiOiJQUyIsInBhaWRVcFRvIjoiMjAxOS0wOS0yNCJ9LHsiY29kZSI6IkdPIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In0seyJjb2RlIjoiRE0iLCJwYWlkVXBUbyI6IjIwMTktMDktMjQifSx7ImNvZGUiOiJDTCIsInBhaWRVcFRvIjoiMjAxOS0wOS0yNCJ9LHsiY29kZSI6IlJTMCIsInBhaWRVcFRvIjoiMjAxOS0wOS0yNCJ9LHsiY29kZSI6IlJDIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In0seyJjb2RlIjoiUkQiLCJwYWlkVXBUbyI6IjIwMTktMDktMjQifSx7ImNvZGUiOiJQQyIsInBhaWRVcFRvIjoiMjAxOS0wOS0yNCJ9LHsiY29kZSI6IlJNIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In0seyJjb2RlIjoiV1MiLCJwYWlkVXBUbyI6IjIwMTktMDktMjQifSx7ImNvZGUiOiJEQiIsInBhaWRVcFRvIjoiMjAxOS0wOS0yNCJ9LHsiY29kZSI6IkRDIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In0seyJjb2RlIjoiUlNVIiwicGFpZFVwVG8iOiIyMDE5LTA5LTI0In1dLCJoYXNoIjoiMTAzMTY2ODUvMCIsImdyYWNlUGVyaW9kRGF5cyI6MCwiYXV0b1Byb2xvbmdhdGVkIjpmYWxzZSwiaXNBdXRvUHJvbG9uZ2F0ZWQiOmZhbHNlfQ==-GvkOuUgCVPdyynFuSG+GNmcDZKp643apInM159fRXb69urSBIFyKO46umkRbl89lwr25SrAcl2TfRG1NMP/zPMRmGvd5VHiXDxa/xatzyPpkGf/czv0GeyuP/XhfX8332kXh9Dnowt3Z++IKUlkTjYInkpg09G9OHSwYcIcHAZ51CsqbrWIaemvDH3P9v+k6EUwwhgDZYA/TplavU/2d9J0EZg8kwzo/TK5P7Za09RFx91YBE558Ncl6VMgdhcwgF+oYHGEfs4Bez5xawJwagLymf3mLhq9acihxGnFsfqcyM/EeKDLKWOAveLQIk1NhtU7YR3fFC0EHEGwb04MavA==-MIIEPjCCAiagAwIBAgIBBTANBgkqhkiG9w0BAQsFADAYMRYwFAYDVQQDDA1KZXRQcm9maWxlIENBMB4XDTE1MTEwMjA4MjE0OFoXDTE4MTEwMTA4MjE0OFowETEPMA0GA1UEAwwGcHJvZDN5MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAxcQkq+zdxlR2mmRYBPzGbUNdMN6OaXiXzxIWtMEkrJMO/5oUfQJbLLuMSMK0QHFmaI37WShyxZcfRCidwXjot4zmNBKnlyHodDij/78TmVqFl8nOeD5+07B8VEaIu7c3E1N+e1doC6wht4I4+IEmtsPAdoaj5WCQVQbrI8KeT8M9VcBIWX7fD0fhexfg3ZRt0xqwMcXGNp3DdJHiO0rCdU+Itv7EmtnSVq9jBG1usMSFvMowR25mju2JcPFp1+I4ZI+FqgR8gyG8oiNDyNEoAbsR3lOpI7grUYSvkB/xVy/VoklPCK2h0f0GJxFjnye8NT1PAywoyl7RmiAVRE/EKwIDAQABo4GZMIGWMAkGA1UdEwQCMAAwHQYDVR0OBBYEFGEpG9oZGcfLMGNBkY7SgHiMGgTcMEgGA1UdIwRBMD+AFKOetkhnQhI2Qb1t4Lm0oFKLl/GzoRykGjAYMRYwFAYDVQQDDA1KZXRQcm9maWxlIENBggkA0myxg7KDeeEwEwYDVR0lBAwwCgYIKwYBBQUHAwEwCwYDVR0PBAQDAgWgMA0GCSqGSIb3DQEBCwUAA4ICAQC9WZuYgQedSuOc5TOUSrRigMw4/+wuC5EtZBfvdl4HT/8vzMW/oUlIP4YCvA0XKyBaCJ2iX+ZCDKoPfiYXiaSiH+HxAPV6J79vvouxKrWg2XV6ShFtPLP+0gPdGq3x9R3+kJbmAm8w+FOdlWqAfJrLvpzMGNeDU14YGXiZ9bVzmIQbwrBA+c/F4tlK/DV07dsNExihqFoibnqDiVNTGombaU2dDup2gwKdL81ua8EIcGNExHe82kjF4zwfadHk3bQVvbfdAwxcDy4xBjs3L4raPLU3yenSzr/OEur1+jfOxnQSmEcMXKXgrAQ9U55gwjcOFKrgOxEdek/Sk1VfOjvS+nuM4eyEruFMfaZHzoQiuw4IqgGc45ohFH0UUyjYcuFxxDSU9lMCv8qdHKm+wnPRb0l9l5vXsCBDuhAGYD6ss+Ga+aDY6f/qXZuUCEUOH3QUNbbCUlviSz6+GiRnt1kA9N2Qachl+2yBfaqUqr8h7Z2gsx5LcIf5kYNsqJ0GavXTVyWh7PYiKX4bs354ZQLUwwa/cG++2+wNWP+HtBhVxMRNTdVhSm38AknZlD+PTAsWGu9GyLmhti2EnVwGybSD2Dxmhxk3IPCkhKAK+pl0eWYGZWG3tJ9mZ7SowcXLWDFAk0lRJnKGFMTggrWjV8GYpw5bq23VmIqqDLgkNzuoog==" Then
            MessageBox.Show("the Code is valid")

            Button3.Visible = False

            For Each column In dt.Columns
                ComboBox1.Items.Add(column.ColumnName)
            Next
            ComboBox1.SelectedIndex = 0
            Button2.Enabled = True
            ComboBox1.Enabled = True
            TextBox3.Enabled = True
            Button4.Enabled = True
        Else
            MessageBox.Show("the Code is invalied")
            Application.Exit()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim selected = DataGridView1.SelectedRows.Count
        If selected > 0 Then
            Dim result = MessageBox.Show("Do you want to delete these " & selected & " Rows", "caption", MessageBoxButtons.YesNo)
            Select Case result
                Case 6 ' i.e "yes"
                    For Each row As DataGridViewRow In DataGridView1.SelectedRows
                        DataGridView1.Rows.Remove(row)
                    Next
                Case 7 ' i.e "no"
                    MessageBox.Show("No any row has been deleted")
            End Select
        End If
        MessageBox.Show("Susuessfilly deleted these " & selected & " Rows")

    End Sub
End Class
