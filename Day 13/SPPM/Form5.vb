﻿Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form5
    Private _myConnection As System.Data.OleDb.OleDbConnection
    Private _data As AutoCompleteStringCollection
    Private _items As List(Of Item)
    Private _printer As PCPrint

    Private Sub Form5_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        _printer = New PCPrint()
        _data = New AutoCompleteStringCollection()
        GetAutoCompleteData(_data)
        Populate()
        ReadAllItems()
    End Sub

    Private Sub Populate()
        OrderGrid.Rows.Clear()
        Dim cmb As New DataGridViewComboBoxColumn()
        cmb.HeaderText = "Item"
        cmb.Name = "cmb"
        cmb.AutoComplete = True
        cmb.MaxDropDownItems = 10
        For Each i As String In _data
            cmb.Items.AddRange(i)
        Next

        OrderGrid.Columns.Add(cmb)
        Dim priceCol As New DataGridViewTextBoxColumn()
        priceCol.HeaderText = "Price"
        OrderGrid.Columns.Add(priceCol)
        Dim qtyCol As New DataGridViewTextBoxColumn()
        qtyCol.HeaderText = "Quantity"
        OrderGrid.Columns.Add(qtyCol)

        OrderGrid.Columns(1).ReadOnly = True

        'OrderGrid.
        'OrderGrid.Rows.Add(New String() {"panadol", 10, 10})
    End Sub

    ''' <summary>
    ''' Only allow number on colum quantity
    ''' Autocomplete for item
    ''' </summary>
    ''' <param name="sender">The source of the event.</param>
    ''' <param name="e">The <see cref="System.Windows.Forms.DataGridViewEditingControlShowingEventArgs"/> instance containing the event data.</param>
    Private Sub OrderGrid_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles OrderGrid.EditingControlShowing
        'Only allow number on colum quantity
        If OrderGrid.CurrentCell.ColumnIndex = 2 Then
            AddHandler CType(e.Control, TextBox).KeyPress, AddressOf TextBox_keyPress
        End If

        If OrderGrid.CurrentCell.ColumnIndex = 0 Then
            Dim cbx As ComboBox = TryCast(e.Control, ComboBox)
            If cbx IsNot Nothing Then
                cbx.AutoCompleteMode = AutoCompleteMode.Suggest
                cbx.AutoCompleteSource = AutoCompleteSource.ListItems

                cbx.DropDownStyle = ComboBoxStyle.DropDown
                'Data
                cbx.AutoCompleteCustomSource = _data

                'Add the handle to your IndexChanged Event
                AddHandler cbx.SelectedIndexChanged, AddressOf editingComboBox_SelectedIndexChanged
                'Prevent this event from firing twice, as is normally the case.
                RemoveHandler OrderGrid.EditingControlShowing, AddressOf OrderGrid_EditingControlShowing
            End If
        End If
    End Sub

    Private Sub editingComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        'ref: http://www.vbforums.com/showthread.php?656274-RESOLVED-Datagridview-Combobox-SelectedIndexChanged
        'Get the editing control
        Dim editingComboBox As ComboBox = TryCast(sender, ComboBox)
        If editingComboBox Is Nothing Then Exit Sub

        'Show your Message Boxes
        'MessageBox.Show(editingComboBox.SelectedIndex.ToString()) ' Display index
        'MessageBox.Show(editingComboBox.Text) ' Display value
        'MessageBox.Show(editingComboBox.Text) ' Display value

        'set the price based on the item name
        Dim i = SearchItem(editingComboBox.Text)
        OrderGrid.Rows(OrderGrid.CurrentCell.RowIndex).Cells(1).Value = i.Price

        'Remove the handle to this event. It will be readded each time a new combobox selection causes the EditingControlShowing Event to fire
        RemoveHandler editingComboBox.SelectedIndexChanged, AddressOf editingComboBox_SelectedIndexChanged
        'Re-enable the EditingControlShowing event so the above can take place.
        AddHandler OrderGrid.EditingControlShowing, AddressOf OrderGrid_EditingControlShowing
    End Sub

    Private Sub GetAutoCompleteData(ByVal data)
        OpenExcelFile()
        Dim query As String = "SELECT * FROM [Sheet1$]"
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(query, _myConnection)

        Dim reader As OleDb.OleDbDataReader = command.ExecuteReader()

        While reader.Read()
            ' read in a string
            data.Add(reader.Item(0))
        End While
        reader.Close()
        CloseExcelFile()
    End Sub

    Private Sub TextBox_keyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)
        If Char.IsDigit(CChar(CStr(e.KeyChar))) = False Then e.Handled = True
    End Sub

    Private Sub ReadAllItems()
        _items = New List(Of Item)

        OpenExcelFile()
        Dim query As String = "SELECT * FROM [sheet1$]"
        Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(query, _myConnection)

        Dim reader As OleDb.OleDbDataReader = command.ExecuteReader()

        While reader.Read()
            Dim i As Item = New Item()
            ' read in a string
            i.Name = reader.Item(0)
            i.Qty = reader.Item(1)
            i.Price = reader.Item(2)
            i.Desc = reader.Item(3)
            _items.Add(i)
        End While
        reader.Close()
        CloseExcelFile()
    End Sub
    ''' <summary>
    ''' Searches the item from the excel file
    ''' </summary>
    ''' <param name="name">The name.</param>
    ''' <returns></returns>
    Private Function SearchItem(ByVal name As String) As Item
        For Each i As Item In _items
            If (i.Name = name) Then
                Return i
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>
    ''' Opens the excel file.
    ''' </summary>
    Private Sub OpenExcelFile()
        _myConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" & Form1.filePath & "';Extended Properties=Excel 8.0;")
        _myConnection.Open()
    End Sub

    ''' <summary>
    ''' Closes the excel file.
    ''' </summary>
    Private Sub CloseExcelFile()
        _myConnection.Close()
    End Sub



    Private Sub SellBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles SellBtn.Click



        For Each row As DataGridViewRow In OrderGrid.Rows
            Dim name = row.Cells(0).Value
            Dim qty = row.Cells(2).Value
            If (name = "") Then
                Exit For
            End If
            OpenExcelFile()
            Dim query = String.Format("Update [sheet1$]" +
                " SET quantity = quantity - {0}" +
                " WHERE item = '{1}'", qty, name)

            Dim updatestock = String.Format("Update [sheet3$]" +
                " SET quantity = quantity + {0}" +
                " WHERE item = '{1}'", qty, name)

            Dim command As OleDbCommand = New OleDbCommand(query, _myConnection)

            Dim commands As OleDbCommand = New OleDbCommand(updatestock, _myConnection)

            'run command
            Dim result = command.ExecuteNonQuery()
            Dim results = commands.ExecuteNonQuery()

            CloseExcelFile()
        Next
        OrderGrid.Rows.Clear()


        Form1.App = New Microsoft.Office.Interop.Excel.Application()
        Form1.workbook = Form1.App.Workbooks.Open(Form1.filePath)

        Form1.worksheet = Form1.workbook.Worksheets("sheet3")
        Dim x As Integer
        Dim stock As Integer
        Dim lLastRow As Long
        stock = 0

        With Form1.worksheet
            'find the last row of the list
            lLastRow = Form1.worksheet.Cells(Form1.worksheet.Rows.Count, "A").End(Excel.XlDirection.xlUp).Row
            'shift from an extra row if list has header
        End With

        For x = 2 To lLastRow

            Form1.worksheet.Cells(x, 5).Value = Form1.worksheet.Cells(x, 5).value + Form1.worksheet.Cells(x, 2).Value * Form1.worksheet.Cells(x, 3).Value



        Next
        Form1.workbook.Save()
        Form1.workbook.Close()
        Form1.App.Quit()

        Form1.releaseObject(Form1.worksheet)
        Form1.releaseObject(Form1.workbook)

        Form1.releaseObject(Form1.App)

        x = 0


    End Sub

    Private Sub PrintBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PrintBtn.Click
        _printer.TextToPrint = DescTxt.Text
        'Issue print command
        _printer.Print()
    End Sub


    Private Sub BackBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackBtn.Click
        LoginForm1.Enabled = True
        LoginForm1.UsernameTextBox.Text = ""
        LoginForm1.PasswordTextBox.Text = ""
        LoginForm1.UsernameTextBox.Focus()
        Me.Hide()

    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        BackBtn.PerformClick()


        LoginForm1.Enabled = True
        LoginForm1.CancelButton.PerformClick()



    End Sub


    Private Sub OrderGrid_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles OrderGrid.CellContentClick

    End Sub

End Class