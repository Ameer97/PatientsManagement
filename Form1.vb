Imports System.Data.OleDb
Public Class Form1
    Dim flag As Boolean = False
    Dim dt As DataTable = New DataTable()
    Dim view As DataView

    Private Sub getXlFile()


        OpenFileDialog1.Filter = "Excel 2003 Documents (*.xls) | *.xls"

        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim fileName As String
            fileName = OpenFileDialog1.FileName
            getExcel(fileName)
        End If

    End Sub

    Private Sub getExcel(openFDFilename As String)

        Dim s As String = ""
        Dim connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=""Excel 8.0; HDR = Yes; IMEX = 1""", openFDFilename)

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
        Else
            MessageBox.Show("You are not a Doctor")
            flag = False

        End If
        Dim watch As Stopwatch = Stopwatch.StartNew()
        watch.Start()
        getXlFile()
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
        ExcelLibrary.DataSetHelper.CreateWorkbook(OpenFileDialog1.FileName, customerTable)
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
        If activate = "a" Then
            For Each column In dt.Columns
                ComboBox1.Items.Add(column.ColumnName)
            Next
            ComboBox1.SelectedIndex = 0
            Button2.Enabled = True
            ComboBox1.Enabled = True
            TextBox3.Enabled = True
            Button4.Enabled = True
        Else
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
