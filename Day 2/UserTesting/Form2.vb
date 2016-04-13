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

       

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With


        If lLastRow = 1 Then
            replace = False

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




        'add the data a row after the end of the list


        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.APP.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)

        Form1.releaseObject(Form1.App)

    End Sub

End Class