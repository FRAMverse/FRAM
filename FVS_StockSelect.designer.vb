<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_StockSelect
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
      Me.StockSelectionTitle = New System.Windows.Forms.Label()
      Me.SelectTypeLabel = New System.Windows.Forms.Label()
      Me.StockListBox = New System.Windows.Forms.ListBox()
      Me.SSDoneButton = New System.Windows.Forms.Button()
      Me.SSCancelButton = New System.Windows.Forms.Button()
      Me.StockGroupNameTextBox = New System.Windows.Forms.TextBox()
      Me.StockGroupLabel = New System.Windows.Forms.Label()
      Me.SuspendLayout()
      '
      'StockSelectionTitle
      '
      Me.StockSelectionTitle.AutoSize = True
      Me.StockSelectionTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.StockSelectionTitle.Location = New System.Drawing.Point(173, 9)
      Me.StockSelectionTitle.Name = "StockSelectionTitle"
      Me.StockSelectionTitle.Size = New System.Drawing.Size(178, 26)
      Me.StockSelectionTitle.TabIndex = 0
      Me.StockSelectionTitle.Text = "Stock Selection"
      '
      'SelectTypeLabel
      '
      Me.SelectTypeLabel.AutoSize = True
      Me.SelectTypeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SelectTypeLabel.Location = New System.Drawing.Point(157, 60)
      Me.SelectTypeLabel.Name = "SelectTypeLabel"
      Me.SelectTypeLabel.Size = New System.Drawing.Size(220, 17)
      Me.SelectTypeLabel.TabIndex = 2
      Me.SelectTypeLabel.Text = "Multi-Stock Selection Allowed"
      '
      'StockListBox
      '
      Me.StockListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.StockListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.StockListBox.FormattingEnabled = True
      Me.StockListBox.ItemHeight = 17
      Me.StockListBox.Location = New System.Drawing.Point(12, 94)
      Me.StockListBox.Name = "StockListBox"
      Me.StockListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
      Me.StockListBox.Size = New System.Drawing.Size(539, 616)
      Me.StockListBox.TabIndex = 3
      '
      'SSDoneButton
      '
      Me.SSDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SSDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SSDoneButton.Location = New System.Drawing.Point(96, 780)
      Me.SSDoneButton.Name = "SSDoneButton"
      Me.SSDoneButton.Size = New System.Drawing.Size(137, 39)
      Me.SSDoneButton.TabIndex = 4
      Me.SSDoneButton.Text = "OK - Done"
      Me.SSDoneButton.UseVisualStyleBackColor = False
      '
      'SSCancelButton
      '
      Me.SSCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SSCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SSCancelButton.Location = New System.Drawing.Point(336, 780)
      Me.SSCancelButton.Name = "SSCancelButton"
      Me.SSCancelButton.Size = New System.Drawing.Size(137, 39)
      Me.SSCancelButton.TabIndex = 5
      Me.SSCancelButton.Text = "Cancel"
      Me.SSCancelButton.UseVisualStyleBackColor = False
      '
      'StockGroupNameTextBox
      '
      Me.StockGroupNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.StockGroupNameTextBox.Location = New System.Drawing.Point(12, 747)
      Me.StockGroupNameTextBox.Name = "StockGroupNameTextBox"
      Me.StockGroupNameTextBox.Size = New System.Drawing.Size(539, 23)
      Me.StockGroupNameTextBox.TabIndex = 6
      '
      'StockGroupLabel
      '
      Me.StockGroupLabel.AutoSize = True
      Me.StockGroupLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.StockGroupLabel.Location = New System.Drawing.Point(9, 727)
      Me.StockGroupLabel.Name = "StockGroupLabel"
      Me.StockGroupLabel.Size = New System.Drawing.Size(114, 13)
      Me.StockGroupLabel.TabIndex = 7
      Me.StockGroupLabel.Text = "Stock Group Name"
      '
      'FVS_StockSelect
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(572, 831)
      Me.Controls.Add(Me.StockGroupLabel)
      Me.Controls.Add(Me.StockGroupNameTextBox)
      Me.Controls.Add(Me.SSCancelButton)
      Me.Controls.Add(Me.SSDoneButton)
      Me.Controls.Add(Me.StockListBox)
      Me.Controls.Add(Me.SelectTypeLabel)
      Me.Controls.Add(Me.StockSelectionTitle)
      Me.Name = "FVS_StockSelect"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Stock Selection"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents StockSelectionTitle As System.Windows.Forms.Label
   Friend WithEvents SelectTypeLabel As System.Windows.Forms.Label
   Friend WithEvents StockListBox As System.Windows.Forms.ListBox
   Friend WithEvents SSDoneButton As System.Windows.Forms.Button
   Friend WithEvents SSCancelButton As System.Windows.Forms.Button
   Friend WithEvents StockGroupNameTextBox As System.Windows.Forms.TextBox
   Friend WithEvents StockGroupLabel As System.Windows.Forms.Label
End Class
