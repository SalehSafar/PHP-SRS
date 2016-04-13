Imports System.Data.OleDb

Public Class Form5
    Private _myConnection As System.Data.OleDb.OleDbConnection
    Private _data As AutoCompleteStringCollection
    Private _items As List(Of Item)


    Private Sub Form5_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load

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


        'Remove the handle to this event. It will be readded each time a new combobox selection causes the EditingControlShowing Event to fire
        RemoveHandler editingComboBox.SelectedIndexChanged, AddressOf editingComboBox_SelectedIndexChanged
        'Re-enable the EditingControlShowing event so the above can take place.
        AddHandler OrderGrid.EditingControlShowing, AddressOf OrderGrid_EditingControlShowing
    End Sub


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

            Dim command As OleDbCommand = New OleDbCommand(query, _myConnection)

            'run command
            Dim result = command.ExecuteNonQuery()

            CloseExcelFile()
        Next
        OrderGrid.Rows.Clear()
        MessageBox.Show("Items are decreased from the stock")
    End Sub


    Private Sub BackBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackBtn.Click
        LoginForm1.Enabled = True
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