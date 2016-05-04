Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class Form1

    Public Shared App As New Excel.Application
    Public Shared worksheet As Excel.Worksheet
    Public Shared workbook As Excel.Workbook

    Public Shared appDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    Public Shared filePath = System.IO.Path.Combine(appDir, "temp.xls")

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        LoginForm1.UsernameTextBox.Clear()
        LoginForm1.PasswordTextBox.Clear()


        LoginForm1.Enabled = True
        Me.Hide()

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If My.Computer.FileSystem.FileExists(filePath) Then
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim DtSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & filePath & "';Extended Properties=Excel 8.0;")
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
            DtSet = New System.Data.DataSet
            MyCommand.Fill(DtSet)
            DataGridView1.DataSource = DtSet.Tables(0)
            MyConnection.Close()
        Else
        End If
    End Sub

    Private Sub DisplayLowStockItems() Handles MyBase.Shown
        Dim MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Form1.filePath & "';Extended Properties=Excel 8.0;")
        MyConnection.Open()
        Dim query As String = "SELECT * FROM [sheet1$]"
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(query, MyConnection)
        Dim reader As OleDb.OleDbDataReader = command.ExecuteReader()
        Dim lowStock = ""
        While reader.Read()
            Dim i As Item = New Item()
            ' read in a string
            i.Name = reader.Item(0)
            i.Qty = reader.Item(1)
            i.Price = reader.Item(2)
            i.Desc = reader.Item(3)

            If (i.Qty <= 2) Then
                lowStock += i.Name + ": " + i.Qty.ToString + Environment.NewLine
            End If
        End While

        If Not String.IsNullOrWhiteSpace(lowStock) Then
            lowStock = "Following items are low on stock: " + Environment.NewLine + lowStock
            MessageBox.Show(lowStock)
        End If
    End Sub
    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Form2.Show()
        Me.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form3.Show()
        Me.Enabled = False


    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Button4.PerformClick()
        LoginForm1.CancelButton.PerformClick()



    End Sub

End Class
