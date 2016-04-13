Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form3

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form1.DataGridView1.DataSource = Nothing
        Form1.DataGridView1.Rows.Clear()
        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim DtSet As System.Data.DataSet
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Form1.filePath & "';Extended Properties=Excel 8.0;")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        DtSet = New System.Data.DataSet
        MyCommand.Fill(DtSet)
        Form1.DataGridView1.DataSource = DtSet.Tables(0)
        MyConnection.Close()

        Form1.Enabled = True

        Me.Hide()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet1")

        Dim x As Integer
        Dim delete As Boolean
        Dim lLastRow As Long
        Dim stock As Integer

        stock = 0

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With

        If lLastRow = 1 Then
            MessageBox.Show("Nothing to Remove", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If

        For x = 2 To lLastRow
            If Form1.worksheet.Cells(x, 1).Value = TextBox1.Text Then
                delete = True
                Exit For
            Else
                delete = False

            End If

        Next

        If delete = True Then

            Form1.worksheet.Rows(x).Delete()


        Else

            MessageBox.Show("No item found", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If

        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.APP.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)
        Form1.releaseObject(Form1.APP)

    End Sub
    Private Sub Form3_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Button2.PerformClick()

        Form1.Button4.PerformClick()
        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub


    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub
End Class