<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_ReportSelection
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
      Me.Label1 = New System.Windows.Forms.Label()
      Me.ReportCheckedListBox = New System.Windows.Forms.CheckedListBox()
      Me.ReportSelectedListBox = New System.Windows.Forms.ListBox()
      Me.DrvDoneButton = New System.Windows.Forms.Button()
      Me.DrvCancelButton = New System.Windows.Forms.Button()
      Me.DrvListLabel = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.RepDriverNameTextBox = New System.Windows.Forms.TextBox()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(271, 24)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(324, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Report Selection for Output Driver"
      '
      'ReportCheckedListBox
      '
      Me.ReportCheckedListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReportCheckedListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReportCheckedListBox.FormattingEnabled = True
      Me.ReportCheckedListBox.Location = New System.Drawing.Point(12, 77)
      Me.ReportCheckedListBox.Name = "ReportCheckedListBox"
      Me.ReportCheckedListBox.Size = New System.Drawing.Size(438, 424)
      Me.ReportCheckedListBox.TabIndex = 1
      '
      'ReportSelectedListBox
      '
      Me.ReportSelectedListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReportSelectedListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReportSelectedListBox.FormattingEnabled = True
      Me.ReportSelectedListBox.ItemHeight = 20
      Me.ReportSelectedListBox.Location = New System.Drawing.Point(493, 77)
      Me.ReportSelectedListBox.Name = "ReportSelectedListBox"
      Me.ReportSelectedListBox.Size = New System.Drawing.Size(431, 644)
      Me.ReportSelectedListBox.TabIndex = 2
      '
      'DrvDoneButton
      '
      Me.DrvDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.DrvDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DrvDoneButton.Location = New System.Drawing.Point(235, 740)
      Me.DrvDoneButton.Name = "DrvDoneButton"
      Me.DrvDoneButton.Size = New System.Drawing.Size(161, 47)
      Me.DrvDoneButton.TabIndex = 3
      Me.DrvDoneButton.Text = "SAVE"
      Me.DrvDoneButton.UseVisualStyleBackColor = False
      '
      'DrvCancelButton
      '
      Me.DrvCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.DrvCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DrvCancelButton.Location = New System.Drawing.Point(558, 740)
      Me.DrvCancelButton.Name = "DrvCancelButton"
      Me.DrvCancelButton.Size = New System.Drawing.Size(161, 47)
      Me.DrvCancelButton.TabIndex = 4
      Me.DrvCancelButton.Text = "CANCEL"
      Me.DrvCancelButton.UseVisualStyleBackColor = False
      '
      'DrvListLabel
      '
      Me.DrvListLabel.AutoSize = True
      Me.DrvListLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DrvListLabel.Location = New System.Drawing.Point(21, 540)
      Me.DrvListLabel.Name = "DrvListLabel"
      Me.DrvListLabel.Size = New System.Drawing.Size(57, 17)
      Me.DrvListLabel.TabIndex = 5
      Me.DrvListLabel.Text = "Label2"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label2.Location = New System.Drawing.Point(21, 648)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(152, 17)
      Me.Label2.TabIndex = 6
      Me.Label2.Text = "Report Driver Name"
      '
      'RepDriverNameTextBox
      '
      Me.RepDriverNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RepDriverNameTextBox.Location = New System.Drawing.Point(25, 671)
      Me.RepDriverNameTextBox.Name = "RepDriverNameTextBox"
      Me.RepDriverNameTextBox.Size = New System.Drawing.Size(371, 23)
      Me.RepDriverNameTextBox.TabIndex = 7
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(103, 853)
      Me.RecordSetNameLabel.Name = "RecordSetNameLabel"
      Me.RecordSetNameLabel.Size = New System.Drawing.Size(121, 17)
      Me.RecordSetNameLabel.TabIndex = 30
      Me.RecordSetNameLabel.Text = "recordset name"
      '
      'DatabaseNameLabel
      '
      Me.DatabaseNameLabel.AutoSize = True
      Me.DatabaseNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.DatabaseNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(103, 819)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(20, 853)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(20, 819)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_ReportSelection
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(946, 875)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.RepDriverNameTextBox)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.DrvListLabel)
      Me.Controls.Add(Me.DrvCancelButton)
      Me.Controls.Add(Me.DrvDoneButton)
      Me.Controls.Add(Me.ReportSelectedListBox)
      Me.Controls.Add(Me.ReportCheckedListBox)
      Me.Controls.Add(Me.Label1)
      Me.Name = "FVS_ReportSelection"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Report Selection"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents ReportCheckedListBox As System.Windows.Forms.CheckedListBox
   Friend WithEvents ReportSelectedListBox As System.Windows.Forms.ListBox
   Friend WithEvents DrvDoneButton As System.Windows.Forms.Button
   Friend WithEvents DrvCancelButton As System.Windows.Forms.Button
   Friend WithEvents DrvListLabel As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents RepDriverNameTextBox As System.Windows.Forms.TextBox
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
