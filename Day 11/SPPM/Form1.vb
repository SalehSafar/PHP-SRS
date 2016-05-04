Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class Form1

    Public Shared App As New Excel.Application
    Public Shared worksheet As Excel.Worksheet
    Public Shared workbook As Excel.Workbook

    Public Shared appDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    Public Shared filePath = System.IO.Path.Combine(appDir, "temp.xls")

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim managerdisplay As String
        managerdisplay = ""
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = workbook.Worksheets("sheet1")

        Dim x As Integer
        Dim lLastRow As Long
        Dim stock As Integer

        stock = 0

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With

        If lLastRow = 1 Then

        End If

        For x = 1 To lLastRow
            managerdisplay += Form1.worksheet.Cells(x, 1).value & vbTab & Form1.worksheet.Cells(x, 2).value & vbTab & Form1.worksheet.Cells(x, 3).value & vbTab & Form1.worksheet.Cells(x, 4).value & vbNewLine
            Form4.csvsave += Form1.worksheet.Cells(x, 1).value & "," & Form1.worksheet.Cells(x, 2).value & "," & Form1.worksheet.Cells(x, 3).value & "," & Form1.worksheet.Cells(x, 4).value & vbNewLine

        Next

        managerdisplay += "Sales From Lu" & vbNewLine
        Form4.csvsave += "Sales From Lu" & vbNewLine


        Form1.worksheet = workbook.Worksheets("sheet3")
        For x = 1 To lLastRow
            managerdisplay += Form1.worksheet.Cells(x, 1).value & vbTab & Form1.worksheet.Cells(x, 2).value & vbTab & Form1.worksheet.Cells(x, 3).value & vbTab & Form1.worksheet.Cells(x, 5).value & vbNewLine
            Form4.csvsave += Form1.worksheet.Cells(x, 1).value & "," & Form1.worksheet.Cells(x, 2).value & "," & Form1.worksheet.Cells(x, 3).value & "," & Form1.worksheet.Cells(x, 5).value & vbNewLine

        Next
        Form4.RichTextBox1.Text = managerdisplay

        workbook.Save()
        workbook.Close()
        App.Quit()

        releaseObject(Form1.worksheet)
        releaseObject(Form1.workbook)
        releaseObject(Form1.App)



        Form4.Show()
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
        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub

End Class
