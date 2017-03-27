<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_FisherySelect
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
      Me.FisherySelectionTitle = New System.Windows.Forms.Label()
      Me.FisheryGroupLabel = New System.Windows.Forms.Label()
      Me.FSCancelButton = New System.Windows.Forms.Button()
      Me.FSDoneButton = New System.Windows.Forms.Button()
      Me.FisheryGroupNameTextBox = New System.Windows.Forms.TextBox()
      Me.SelectTypeLabel = New System.Windows.Forms.Label()
      Me.FisheryListBox = New System.Windows.Forms.ListBox()
      Me.FSAllButton = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'FisherySelectionTitle
      '
      Me.FisherySelectionTitle.AutoSize = True
      Me.FisherySelectionTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisherySelectionTitle.Location = New System.Drawing.Point(207, 9)
      Me.FisherySelectionTitle.Name = "FisherySelectionTitle"
      Me.FisherySelectionTitle.Size = New System.Drawing.Size(172, 24)
      Me.FisherySelectionTitle.TabIndex = 0
      Me.FisherySelectionTitle.Text = "Fishery Selection"
      '
      'FisheryGroupLabel
      '
      Me.FisheryGroupLabel.AutoSize = True
      Me.FisheryGroupLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisheryGroupLabel.Location = New System.Drawing.Point(12, 700)
      Me.FisheryGroupLabel.Name = "FisheryGroupLabel"
      Me.FisheryGroupLabel.Size = New System.Drawing.Size(121, 13)
      Me.FisheryGroupLabel.TabIndex = 11
      Me.FisheryGroupLabel.Text = "Fishery Group Name"
      '
      'FSCancelButton
      '
      Me.FSCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCancelButton.Location = New System.Drawing.Point(418, 734)
      Me.FSCancelButton.Name = "FSCancelButton"
      Me.FSCancelButton.Size = New System.Drawing.Size(137, 39)
      Me.FSCancelButton.TabIndex = 9
      Me.FSCancelButton.Text = "Cancel"
      Me.FSCancelButton.UseVisualStyleBackColor = False
      '
      'FSDoneButton
      '
      Me.FSDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSDoneButton.Location = New System.Drawing.Point(224, 734)
      Me.FSDoneButton.Name = "FSDoneButton"
      Me.FSDoneButton.Size = New System.Drawing.Size(137, 39)
      Me.FSDoneButton.TabIndex = 8
      Me.FSDoneButton.Text = "OK - Done"
      Me.FSDoneButton.UseVisualStyleBackColor = False
      '
      'FisheryGroupNameTextBox
      '
      Me.FisheryGroupNameTextBox.Location = New System.Drawing.Point(172, 697)
      Me.FisheryGroupNameTextBox.Name = "FisheryGroupNameTextBox"
      Me.FisheryGroupNameTextBox.Size = New System.Drawing.Size(411, 20)
      Me.FisheryGroupNameTextBox.TabIndex = 12
      '
      'SelectTypeLabel
      '
      Me.SelectTypeLabel.AutoSize = True
      Me.SelectTypeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SelectTypeLabel.Location = New System.Drawing.Point(180, 55)
      Me.SelectTypeLabel.Name = "SelectTypeLabel"
      Me.SelectTypeLabel.Size = New System.Drawing.Size(233, 17)
      Me.SelectTypeLabel.TabIndex = 13
      Me.SelectTypeLabel.Text = "Multi-Fishery Selection Allowed"
      '
      'FisheryListBox
      '
      Me.FisheryListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FisheryListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisheryListBox.FormattingEnabled = True
      Me.FisheryListBox.ItemHeight = 17
      Me.FisheryListBox.Location = New System.Drawing.Point(41, 99)
      Me.FisheryListBox.Name = "FisheryListBox"
      Me.FisheryListBox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
      Me.FisheryListBox.Size = New System.Drawing.Size(567, 548)
      Me.FisheryListBox.TabIndex = 14
      '
      'FSAllButton
      '
      Me.FSAllButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSAllButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSAllButton.Location = New System.Drawing.Point(41, 734)
      Me.FSAllButton.Name = "FSAllButton"
      Me.FSAllButton.Size = New System.Drawing.Size(137, 39)
      Me.FSAllButton.TabIndex = 15
      Me.FSAllButton.Text = "Select All"
      Me.FSAllButton.UseVisualStyleBackColor = False
      '
      'FVS_FisherySelect
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(647, 784)
      Me.Controls.Add(Me.FSAllButton)
      Me.Controls.Add(Me.FisheryListBox)
      Me.Controls.Add(Me.SelectTypeLabel)
      Me.Controls.Add(Me.FisheryGroupNameTextBox)
      Me.Controls.Add(Me.FisheryGroupLabel)
      Me.Controls.Add(Me.FSCancelButton)
      Me.Controls.Add(Me.FSDoneButton)
      Me.Controls.Add(Me.FisherySelectionTitle)
      Me.Name = "FVS_FisherySelect"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Fishery Selection"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents FisherySelectionTitle As System.Windows.Forms.Label
   Friend WithEvents FisheryGroupLabel As System.Windows.Forms.Label
   Friend WithEvents FSCancelButton As System.Windows.Forms.Button
   Friend WithEvents FSDoneButton As System.Windows.Forms.Button
   Friend WithEvents FisheryGroupNameTextBox As System.Windows.Forms.TextBox
   Friend WithEvents SelectTypeLabel As System.Windows.Forms.Label
   Friend WithEvents FisheryListBox As System.Windows.Forms.ListBox
   Friend WithEvents FSAllButton As System.Windows.Forms.Button
End Class
