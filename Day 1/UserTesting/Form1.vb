Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class Form1

    Public Shared App As New Excel.Application
    Public Shared worksheet As Excel.Worksheet
    Public Shared workbook As Excel.Workbook

    Public Shared appDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
    Public Shared filePath = System.IO.Path.Combine(appDir, "temp.xls")


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

End Class
