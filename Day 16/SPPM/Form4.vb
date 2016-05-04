Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form4

    Public csvsave As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LoginForm1.Enabled = True
        LoginForm1.UsernameTextBox.Text = ""
        LoginForm1.PasswordTextBox.Text = ""
        LoginForm1.UsernameTextBox.Focus()



        Me.Hide()
        

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        LoginForm1.UsernameTextBox.Clear()
        LoginForm1.PasswordTextBox.Clear()
        Form1.Enabled = True

        Form1.Show()
        Me.Hide()

    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub


    Private Sub Form4_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Button2.PerformClick()

        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim csvFile As String = My.Application.Info.DirectoryPath & "\Test.csv"
        Dim outFile As IO.StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(csvFile, False)


        outFile.WriteLine(csvsave)
        
        outFile.Close()
        Console.WriteLine(My.Computer.FileSystem.ReadAllText(csvFile))

        Label1.Text = "CSV successfully saved."

    End Sub

    
    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Form3.Show()
        Me.Hide()

    End Sub
End Class