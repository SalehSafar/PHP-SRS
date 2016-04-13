Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class LoginForm1

    ' TODO: Insert code to perform custom authentication using the provided username and password 
    ' (See http://go.microsoft.com/fwlink/?LinkId=35339).  
    ' The custom principal can then be attached to the current thread's principal as follows: 
    '     My.User.CurrentPrincipal = CustomPrincipal
    ' where CustomPrincipal is the IPrincipal implementation used to perform authentication. 
    ' Subsequently, My.User will return identity information encapsulated in the CustomPrincipal object
    ' such as the username, display name, etc.

    Private Sub OK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK.Click
        Dim misValue As Object = System.Reflection.Missing.Value

        If Form1.App Is Nothing Then
            MessageBox.Show("Excel is not properly installed")
            Return

        End If

        If My.Computer.FileSystem.FileExists(Form1.filePath) Then

            GoTo _next

        Else

            Form1.App = New Microsoft.Office.Interop.Excel.Application()
            Form1.workbook = Form1.App.Workbooks.Add(misValue)
            Form1.worksheet = Form1.workbook.Sheets("sheet1")

            Form1.worksheet.Cells(1, 1) = "Item"
            Form1.worksheet.Cells(1, 2) = "Quantity"
            Form1.worksheet.Cells(1, 3) = "Price"
            Form1.worksheet.Cells(1, 4) = "Description"

            Dim formatrange As Excel.Range
            formatrange = Form1.worksheet.Range("a1")
            formatrange.NumberFormat = "@"


            formatrange = Form1.worksheet.Range("d1")
            formatrange.NumberFormat = "@"

            Form1.worksheet = Form1.workbook.Sheets("sheet2")
            Form1.worksheet.Cells(1, 1) = "Username"
            Form1.worksheet.Cells(1, 2) = "Password"
            Form1.worksheet.Cells(1, 3) = "Status"

            Form1.worksheet.Cells(2, 1) = "admin"
            Form1.worksheet.Cells(2, 2) = "1234"
            Form1.worksheet.Cells(2, 3) = "admin"

            Form1.worksheet.Cells(3, 1) = "Lu"
            Form1.worksheet.Cells(3, 2) = "4321"
            Form1.worksheet.Cells(3, 3) = "clerk"


            Form1.workbook.SaveAs(Form1.filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, _
             Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            Form1.workbook.Close(True, misValue, misValue)
            Form1.App.Quit()

            Form1.releaseObject(Form1.worksheet)
            Form1.releaseObject(Form1.workbook)

            Form1.releaseObject(Form1.App)
            GoTo _next
        End If

_next:
        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet2")

        Dim x As Integer
        Dim lLastRow As Long

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With


        For x = 2 To lLastRow

            If Form1.worksheet.Cells(x, 1).Value = UsernameTextBox.Text Then
                If Form1.worksheet.Cells(x, 2).Value = PasswordTextBox.Text Then

                    If Form1.worksheet.Cells(x, 3).Value = "admin" Then
                        MessageBox.Show("Log in as admin: " & Form1.worksheet.Cells(x, 1).Value & Form1.worksheet.Cells(x, 2).Value, "Class", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)

                        Form1.workbook.Save()
                        Form1.workbook.Close()
                        Form1.App.Quit()

                        Form1.releaseObject(Form1.worksheet)
                        Form1.releaseObject(Form1.workbook)
                        Form1.releaseObject(Form1.App)

                        Form4.Show()
                        Me.Enabled = False
                        Exit For
                    ElseIf Form1.worksheet.Cells(x, 3).Value = "clerk" Then
                        MessageBox.Show("Log in as general worker: " & Form1.worksheet.Cells(x, 1).Value & Form1.worksheet.Cells(x, 2).Value, "Class", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)
                        Form1.workbook.Save()
                        Form1.workbook.Close()
                        Form1.App.Quit()

                        Form1.releaseObject(Form1.worksheet)
                        Form1.releaseObject(Form1.workbook)
                        Form1.releaseObject(Form1.App)

                        Form5.Show()
                        Me.Enabled = False
                        Exit For
                    Else
                        MessageBox.Show("Wrong username or password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button2)

                    End If
                Else

                End If

            Else

            End If
        Next



    End Sub

    Private Sub Cancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel.Click
        Me.Close()
    End Sub


End Class
