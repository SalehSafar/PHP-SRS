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
        Me.BackBtn = New System.Windows.Forms.Button()
        Me.PrintBtn = New System.Windows.Forms.Button()
        Me.SellBtn = New System.Windows.Forms.Button()
        Me.OrderGrid = New System.Windows.Forms.DataGridView()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DescTxt = New System.Windows.Forms.RichTextBox()
        CType(Me.OrderGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BackBtn
        '
        Me.BackBtn.Location = New System.Drawing.Point(12, 415)
        Me.BackBtn.Name = "BackBtn"
        Me.BackBtn.Size = New System.Drawing.Size(75, 23)
        Me.BackBtn.TabIndex = 36
        Me.BackBtn.Text = "Back"
        Me.BackBtn.UseVisualStyleBackColor = True
        '
        'PrintBtn
        '
        Me.PrintBtn.Location = New System.Drawing.Point(547, 325)
        Me.PrintBtn.Name = "PrintBtn"
        Me.PrintBtn.Size = New System.Drawing.Size(75, 23)
        Me.PrintBtn.TabIndex = 35
        Me.PrintBtn.Text = "Print"
        Me.PrintBtn.UseVisualStyleBackColor = True
        '
        'SellBtn
        '
        Me.SellBtn.Location = New System.Drawing.Point(431, 325)
        Me.SellBtn.Name = "SellBtn"
        Me.SellBtn.Size = New System.Drawing.Size(75, 23)
        Me.SellBtn.TabIndex = 34
        Me.SellBtn.Text = "Sell"
        Me.SellBtn.UseVisualStyleBackColor = True
        '
        'OrderGrid
        '
        Me.OrderGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.OrderGrid.Location = New System.Drawing.Point(308, 39)
        Me.OrderGrid.Name = "OrderGrid"
        Me.OrderGrid.Size = New System.Drawing.Size(432, 233)
        Me.OrderGrid.TabIndex = 33
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 32
        Me.Label1.Text = "Description"
        '
        'DescTxt
        '
        Me.DescTxt.Location = New System.Drawing.Point(12, 39)
        Me.DescTxt.Name = "DescTxt"
        Me.DescTxt.Size = New System.Drawing.Size(205, 233)
        Me.DescTxt.TabIndex = 31
        Me.DescTxt.Text = ""
        '
        'Form4
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(761, 452)
        Me.Controls.Add(Me.BackBtn)
        Me.Controls.Add(Me.PrintBtn)
        Me.Controls.Add(Me.SellBtn)
        Me.Controls.Add(Me.OrderGrid)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DescTxt)
        Me.Name = "Form4"
        Me.Text = "Form4"
        CType(Me.OrderGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BackBtn As System.Windows.Forms.Button
    Friend WithEvents PrintBtn As System.Windows.Forms.Button
    Friend WithEvents SellBtn As System.Windows.Forms.Button
    Friend WithEvents OrderGrid As System.Windows.Forms.DataGridView
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DescTxt As System.Windows.Forms.RichTextBox
End Class
