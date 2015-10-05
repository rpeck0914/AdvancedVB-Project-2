<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataBaseForm
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
        Me.BindButton = New System.Windows.Forms.Button()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.FilterButton = New System.Windows.Forms.Button()
        Me.FilterTextBox = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ProductNameTextBox = New System.Windows.Forms.TextBox()
        Me.UnitPriceTextBox = New System.Windows.Forms.TextBox()
        Me.InsertButton = New System.Windows.Forms.Button()
        Me.DeleteButton = New System.Windows.Forms.Button()
        Me.UpdateButton = New System.Windows.Forms.Button()
        Me.ProductIDLabel = New System.Windows.Forms.Label()
        Me.IDTextBox = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.ExitButton = New System.Windows.Forms.Button()
        Me.ClearDataButton = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BindButton
        '
        Me.BindButton.Location = New System.Drawing.Point(12, 110)
        Me.BindButton.Name = "BindButton"
        Me.BindButton.Size = New System.Drawing.Size(99, 23)
        Me.BindButton.TabIndex = 11
        Me.BindButton.Text = "&Bind To List"
        Me.BindButton.UseVisualStyleBackColor = True
        '
        'ListBox1
        '
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 147)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(130, 394)
        Me.ListBox1.TabIndex = 14
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(148, 139)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(381, 402)
        Me.DataGridView1.TabIndex = 15
        '
        'FilterButton
        '
        Me.FilterButton.Location = New System.Drawing.Point(254, 110)
        Me.FilterButton.Name = "FilterButton"
        Me.FilterButton.Size = New System.Drawing.Size(75, 23)
        Me.FilterButton.TabIndex = 13
        Me.FilterButton.Text = "&Filter"
        Me.FilterButton.UseVisualStyleBackColor = True
        '
        'FilterTextBox
        '
        Me.FilterTextBox.Location = New System.Drawing.Point(148, 112)
        Me.FilterTextBox.Name = "FilterTextBox"
        Me.FilterTextBox.Size = New System.Drawing.Size(100, 20)
        Me.FilterTextBox.TabIndex = 12
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(78, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Product Name:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Unit Price:"
        '
        'ProductNameTextBox
        '
        Me.ProductNameTextBox.Location = New System.Drawing.Point(96, 16)
        Me.ProductNameTextBox.Name = "ProductNameTextBox"
        Me.ProductNameTextBox.Size = New System.Drawing.Size(100, 20)
        Me.ProductNameTextBox.TabIndex = 1
        '
        'UnitPriceTextBox
        '
        Me.UnitPriceTextBox.Location = New System.Drawing.Point(96, 42)
        Me.UnitPriceTextBox.Name = "UnitPriceTextBox"
        Me.UnitPriceTextBox.Size = New System.Drawing.Size(100, 20)
        Me.UnitPriceTextBox.TabIndex = 3
        '
        'InsertButton
        '
        Me.InsertButton.Location = New System.Drawing.Point(230, 14)
        Me.InsertButton.Name = "InsertButton"
        Me.InsertButton.Size = New System.Drawing.Size(99, 23)
        Me.InsertButton.TabIndex = 4
        Me.InsertButton.Text = "&Insert"
        Me.InsertButton.UseVisualStyleBackColor = True
        '
        'DeleteButton
        '
        Me.DeleteButton.Location = New System.Drawing.Point(230, 72)
        Me.DeleteButton.Name = "DeleteButton"
        Me.DeleteButton.Size = New System.Drawing.Size(99, 23)
        Me.DeleteButton.TabIndex = 6
        Me.DeleteButton.Text = "&Delete"
        Me.DeleteButton.UseVisualStyleBackColor = True
        '
        'UpdateButton
        '
        Me.UpdateButton.Location = New System.Drawing.Point(230, 43)
        Me.UpdateButton.Name = "UpdateButton"
        Me.UpdateButton.Size = New System.Drawing.Size(99, 23)
        Me.UpdateButton.TabIndex = 5
        Me.UpdateButton.Text = "&Update"
        Me.UpdateButton.UseVisualStyleBackColor = True
        '
        'ProductIDLabel
        '
        Me.ProductIDLabel.AutoSize = True
        Me.ProductIDLabel.Location = New System.Drawing.Point(335, 48)
        Me.ProductIDLabel.Name = "ProductIDLabel"
        Me.ProductIDLabel.Size = New System.Drawing.Size(61, 13)
        Me.ProductIDLabel.TabIndex = 7
        Me.ProductIDLabel.Text = "Product ID:"
        '
        'IDTextBox
        '
        Me.IDTextBox.Location = New System.Drawing.Point(402, 45)
        Me.IDTextBox.Name = "IDTextBox"
        Me.IDTextBox.Size = New System.Drawing.Size(100, 20)
        Me.IDTextBox.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(335, 72)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(124, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "*ID Required For Update"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(335, 85)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(111, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = " And Deleting Of Data"
        '
        'ExitButton
        '
        Me.ExitButton.Location = New System.Drawing.Point(442, 553)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(87, 40)
        Me.ExitButton.TabIndex = 16
        Me.ExitButton.Text = "E&xit"
        Me.ExitButton.UseVisualStyleBackColor = True
        '
        'ClearDataButton
        '
        Me.ClearDataButton.Location = New System.Drawing.Point(349, 553)
        Me.ClearDataButton.Name = "ClearDataButton"
        Me.ClearDataButton.Size = New System.Drawing.Size(87, 40)
        Me.ClearDataButton.TabIndex = 17
        Me.ClearDataButton.Text = "&Clear Data"
        Me.ClearDataButton.UseVisualStyleBackColor = True
        '
        'DataBaseForm
        '
        Me.AcceptButton = Me.FilterButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ExitButton
        Me.ClientSize = New System.Drawing.Size(541, 605)
        Me.Controls.Add(Me.ClearDataButton)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.IDTextBox)
        Me.Controls.Add(Me.ProductIDLabel)
        Me.Controls.Add(Me.UpdateButton)
        Me.Controls.Add(Me.DeleteButton)
        Me.Controls.Add(Me.InsertButton)
        Me.Controls.Add(Me.UnitPriceTextBox)
        Me.Controls.Add(Me.ProductNameTextBox)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.BindButton)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.FilterButton)
        Me.Controls.Add(Me.FilterTextBox)
        Me.Name = "DataBaseForm"
        Me.Text = "Data Base Northwind"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BindButton As System.Windows.Forms.Button
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents FilterButton As System.Windows.Forms.Button
    Friend WithEvents FilterTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ProductNameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents UnitPriceTextBox As System.Windows.Forms.TextBox
    Friend WithEvents InsertButton As System.Windows.Forms.Button
    Friend WithEvents DeleteButton As System.Windows.Forms.Button
    Friend WithEvents UpdateButton As System.Windows.Forms.Button
    Friend WithEvents ProductIDLabel As System.Windows.Forms.Label
    Friend WithEvents IDTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents ClearDataButton As System.Windows.Forms.Button

End Class
