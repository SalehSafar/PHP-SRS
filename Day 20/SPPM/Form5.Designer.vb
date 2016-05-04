<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form5
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.DescTxt = New System.Windows.Forms.RichTextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.OrderGrid = New System.Windows.Forms.DataGridView()
        Me.SellBtn = New System.Windows.Forms.Button()
        Me.PrintBtn = New System.Windows.Forms.Button()
        Me.BackBtn = New System.Windows.Forms.Button()
        CType(Me.OrderGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DescTxt
        '
        Me.DescTxt.Location = New System.Drawing.Point(40, 55)
        Me.DescTxt.Name = "DescTxt"
        Me.DescTxt.Size = New System.Drawing.Size(205, 233)
        Me.DescTxt.TabIndex = 0
        Me.DescTxt.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Description"
        '
        'OrderGrid
        '
        Me.OrderGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.OrderGrid.Location = New System.Drawing.Point(336, 55)
        Me.OrderGrid.Name = "OrderGrid"
        Me.OrderGrid.Size = New System.Drawing.Size(432, 233)
        Me.OrderGrid.TabIndex = 21
        '
        'SellBtn
        '
        Me.SellBtn.Location = New System.Drawing.Point(459, 341)
        Me.SellBtn.Name = "SellBtn"
        Me.SellBtn.Size = New System.Drawing.Size(75, 23)
        Me.SellBtn.TabIndex = 22
        Me.SellBtn.Text = "Sell"
        Me.SellBtn.UseVisualStyleBackColor = True
        '
        'PrintBtn
        '
        Me.PrintBtn.Location = New System.Drawing.Point(575, 341)
        Me.PrintBtn.Name = "PrintBtn"
        Me.PrintBtn.Size = New System.Drawing.Size(75, 23)
        Me.PrintBtn.TabIndex = 23
        Me.PrintBtn.Text = "Print"
        Me.PrintBtn.UseVisualStyleBackColor = True
        '
        'BackBtn
        '
        Me.BackBtn.Location = New System.Drawing.Point(40, 431)
        Me.BackBtn.Name = "BackBtn"
        Me.BackBtn.Size = New System.Drawing.Size(75, 23)
        Me.BackBtn.TabIndex = 24
        Me.BackBtn.Text = "Back"
        Me.BackBtn.UseVisualStyleBackColor = True
        '
        'Form5
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(838, 466)
        Me.Controls.Add(Me.BackBtn)
        Me.Controls.Add(Me.PrintBtn)
        Me.Controls.Add(Me.SellBtn)
        Me.Controls.Add(Me.OrderGrid)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DescTxt)
        Me.Name = "Form5"
        Me.Text = "Form5"
        CType(Me.OrderGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DescTxt As RichTextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents OrderGrid As DataGridView
    Friend WithEvents SellBtn As Button
    Friend WithEvents PrintBtn As Button
    Friend WithEvents BackBtn As Button
End Class
