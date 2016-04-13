Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Form1.Enabled = True
        Me.Hide()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet1")

        Dim x As Integer
        Dim replace As Boolean
        Dim lLastRow As Long
        Dim stock As Integer

        stock = 0

        If TextBox1.Text = "" Then
            MessageBox.Show("Item name cannot be empty")
            Return
        End If

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With


        If lLastRow = 1 Then
            replace = False
            GoTo end_for
        End If


        For x = 2 To lLastRow
            If Form1.worksheet.Cells(x, 1).Value.ToString = TextBox1.Text Then
                replace = True
                stock = Form1.worksheet.Cells(x, 2).Value
                stock = stock + Integer.Parse(TextBox2.Text)
                Exit For
            Else

                replace = False
            End If

        Next
end_for:
        Dim result As DialogResult = MessageBox.Show("Confirm the changes:" & vbNewLine & "Item:" & TextBox1.Text & vbNewLine & "Available Stock:" & stock & vbNewLine & "Price:" & TextBox3.Text & vbNewLine & "Description:" & TextBox4.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)

        If result = DialogResult.Yes Then

            GoTo coming
        Else

            Return

        End If

coming:
        If replace = True Then
            Form1.worksheet.Cells(x, 2).Value = stock
            Form1.worksheet.Cells(x, 3).Value = TextBox3.Text
            Form1.worksheet.Cells(x, 4).Value = TextBox4.Text

            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
        Else
            Form1.worksheet.Rows(lLastRow + 1).Insert()
            Dim formatrange As Excel.Range
            formatrange = Form1.worksheet.Range("a" & lLastRow + 1)
            formatrange.NumberFormat = "@"


            formatrange = Form1.worksheet.Range("d" & lLastRow + 1)
            formatrange.NumberFormat = "@"


            Form1.worksheet.Cells(lLastRow + 1, 1).Value = TextBox1.Text
            Form1.worksheet.Cells(lLastRow + 1, 2).Value = TextBox2.Text
            Form1.worksheet.Cells(lLastRow + 1, 3).Value = TextBox3.Text
            Form1.worksheet.Cells(lLastRow + 1, 4).Value = TextBox4.Text

            TextBox1.Clear()
            TextBox2.Clear()
            TextBox3.Clear()
            TextBox4.Clear()
        End If



        'add the data a row after the end of the list


        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.APP.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)

        Form1.releaseObject(Form1.App)

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


    End Sub

    Private Sub TextBox2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox3.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If
    End Sub

    Private Sub Form2_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Button2.PerformClick()

        Form1.Button4.PerformClick()
        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub

End Class