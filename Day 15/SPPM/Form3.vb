Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form3

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Form4.Show()
        Me.Hide()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        

        '  Dim Conn As OleDbConnection = New OleDbConnection

        '     Conn = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Form1.filePath & "';Extended Properties=Excel 8.0;")
        '   Conn.Open()
        '   Dim cmd As OleDbCommand = New OleDbCommand("SELECT [Date], [TotalSales] FROM [sheet3$]", Conn)
        '   Dim dr As OleDbDataReader = cmd.ExecuteReader
        '   While dr.Read
        'Chart1.Series("sheet3$").Points.AddXY(dr("Date").ToString, dr("TotalSales").ToString)
        '   End While
        '   dr.Close()
        'Reference: http://www.visual-basic-tutorials.com/display-data-as-charts-and-graph-in-visual-basic.html#sthash.aRIaIOEN.dpuf

        Dim lLastRow As Integer
        Dim dates() As String
        Dim total() As Integer
        Dim count As Integer
        Dim x As Integer

        dates = {0}
        total = {0}
        count = 0
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet3")

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "H").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header


        End With

        For x = 2 To lLastRow
            If String.IsNullOrEmpty(Form1.worksheet.Cells(x, 8).Value) Then
            Else
                count = count + 1
            End If
        Next

        


        

        For x = 2 To lLastRow

            If String.IsNullOrEmpty(Form1.worksheet.Cells(x, 8).Value) Then
            Else
                Array.Resize(dates, dates.Length + 1)
                dates(dates.Length - 1) = Form1.worksheet.Cells(x, 8).value

                Array.Resize(total, total.Length + 1)
                total(total.Length - 1) = Form1.worksheet.Cells(x, 9).value.ToString
            End If

            ' dates() = Form1.worksheet.Cells(lLastRow, 8).value.ToString


        Next

        Chart1.Series(0).Points.DataBindXY(dates, total)



        

        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.App.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)

        Form1.releaseObject(Form1.App)


















    End Sub


    

    Private Sub Form2_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Button2.PerformClick()

        Form1.Button4.PerformClick()
        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub
End Class