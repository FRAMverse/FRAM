<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_RunModel
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
      Me.ModelRunTitleLabel = New System.Windows.Forms.Label()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.TammTextLabel = New System.Windows.Forms.Label()
      Me.TammNameLabel = New System.Windows.Forms.Label()
      Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
      Me.SelectTAMMButton = New System.Windows.Forms.Button()
      Me.ReplaceQuotaCheck = New System.Windows.Forms.CheckBox()
      Me.OldTammCheck = New System.Windows.Forms.CheckBox()
      Me.TammFwsCheck = New System.Windows.Forms.CheckBox()
      Me.ChinookBYCheck = New System.Windows.Forms.CheckBox()
      Me.RunModelButton = New System.Windows.Forms.Button()
      Me.CancelRunButton = New System.Windows.Forms.Button()
      Me.RunProgressLabel = New System.Windows.Forms.Label()
      Me.MRProgressBar = New System.Windows.Forms.ProgressBar()
      Me.MSFBias = New System.Windows.Forms.CheckBox()
      Me.SuspendLayout()
      '
      'ModelRunTitleLabel
      '
      Me.ModelRunTitleLabel.AutoSize = True
      Me.ModelRunTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ModelRunTitleLabel.Location = New System.Drawing.Point(286, 40)
      Me.ModelRunTitleLabel.Name = "ModelRunTitleLabel"
      Me.ModelRunTitleLabel.Size = New System.Drawing.Size(247, 24)
      Me.ModelRunTitleLabel.TabIndex = 0
      Me.ModelRunTitleLabel.Text = "Model Run Specifications"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(245, 143)
      Me.RecordSetNameLabel.Name = "RecordSetNameLabel"
      Me.RecordSetNameLabel.Size = New System.Drawing.Size(121, 17)
      Me.RecordSetNameLabel.TabIndex = 19
      Me.RecordSetNameLabel.Text = "recordset name"
      '
      'DatabaseNameLabel
      '
      Me.DatabaseNameLabel.AutoSize = True
      Me.DatabaseNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.DatabaseNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(245, 109)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 18
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(142, 143)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(84, 17)
      Me.RecordSetTextLabel.TabIndex = 17
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AccessibleRole = System.Windows.Forms.AccessibleRole.None
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(140, 109)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(77, 17)
      Me.DatabaseTextLabel.TabIndex = 16
      Me.DatabaseTextLabel.Text = "Database"
      '
      'TammTextLabel
      '
      Me.TammTextLabel.AutoSize = True
      Me.TammTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.TammTextLabel.Location = New System.Drawing.Point(243, 185)
      Me.TammTextLabel.Name = "TammTextLabel"
      Me.TammTextLabel.Size = New System.Drawing.Size(149, 17)
      Me.TammTextLabel.TabIndex = 20
      Me.TammTextLabel.Text = "TAMM Spreadsheet"
      '
      'TammNameLabel
      '
      Me.TammNameLabel.AutoSize = True
      Me.TammNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.TammNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.TammNameLabel.Location = New System.Drawing.Point(422, 185)
      Me.TammNameLabel.Name = "TammNameLabel"
      Me.TammNameLabel.Size = New System.Drawing.Size(142, 17)
      Me.TammNameLabel.TabIndex = 21
      Me.TammNameLabel.Text = "spreadsheet name"
      '
      'OpenFileDialog1
      '
      Me.OpenFileDialog1.FileName = "OpenFileDialog1"
      '
      'SelectTAMMButton
      '
      Me.SelectTAMMButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SelectTAMMButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SelectTAMMButton.Location = New System.Drawing.Point(47, 174)
      Me.SelectTAMMButton.Name = "SelectTAMMButton"
      Me.SelectTAMMButton.Size = New System.Drawing.Size(148, 44)
      Me.SelectTAMMButton.TabIndex = 22
      Me.SelectTAMMButton.Text = "Select TAMM"
      Me.SelectTAMMButton.UseVisualStyleBackColor = False
      '
      'ReplaceQuotaCheck
      '
      Me.ReplaceQuotaCheck.AutoSize = True
      Me.ReplaceQuotaCheck.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReplaceQuotaCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReplaceQuotaCheck.Location = New System.Drawing.Point(247, 387)
      Me.ReplaceQuotaCheck.Name = "ReplaceQuotaCheck"
      Me.ReplaceQuotaCheck.Size = New System.Drawing.Size(293, 21)
      Me.ReplaceQuotaCheck.TabIndex = 23
      Me.ReplaceQuotaCheck.Text = "Replace Quotas with Fishery Scalers"
      Me.ReplaceQuotaCheck.UseVisualStyleBackColor = False
      '
      'OldTammCheck
      '
      Me.OldTammCheck.AutoSize = True
      Me.OldTammCheck.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.OldTammCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.OldTammCheck.Location = New System.Drawing.Point(247, 245)
      Me.OldTammCheck.Name = "OldTammCheck"
      Me.OldTammCheck.Size = New System.Drawing.Size(401, 21)
      Me.OldTammCheck.TabIndex = 24
      Me.OldTammCheck.Text = "Old Chinook TAMM Format (10+11 Sport Combined)"
      Me.OldTammCheck.UseVisualStyleBackColor = False
      '
      'TammFwsCheck
      '
      Me.TammFwsCheck.AutoSize = True
      Me.TammFwsCheck.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.TammFwsCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.TammFwsCheck.Location = New System.Drawing.Point(247, 290)
      Me.TammFwsCheck.Name = "TammFwsCheck"
      Me.TammFwsCheck.Size = New System.Drawing.Size(315, 21)
      Me.TammFwsCheck.TabIndex = 25
      Me.TammFwsCheck.Text = "Use Chinook TAMM FWS (No Iterations)"
      Me.TammFwsCheck.UseVisualStyleBackColor = False
      '
      'ChinookBYCheck
      '
      Me.ChinookBYCheck.AutoSize = True
      Me.ChinookBYCheck.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ChinookBYCheck.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ChinookBYCheck.Location = New System.Drawing.Point(247, 339)
      Me.ChinookBYCheck.Name = "ChinookBYCheck"
      Me.ChinookBYCheck.Size = New System.Drawing.Size(263, 21)
      Me.ChinookBYCheck.TabIndex = 26
      Me.ChinookBYCheck.Text = "Chinook Brood Year AEQ Report"
      Me.ChinookBYCheck.UseVisualStyleBackColor = False
      '
      'RunModelButton
      '
      Me.RunModelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.RunModelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RunModelButton.Location = New System.Drawing.Point(228, 582)
      Me.RunModelButton.Name = "RunModelButton"
      Me.RunModelButton.Size = New System.Drawing.Size(148, 44)
      Me.RunModelButton.TabIndex = 27
      Me.RunModelButton.Text = "RUN Model"
      Me.RunModelButton.UseVisualStyleBackColor = False
      '
      'CancelRunButton
      '
      Me.CancelRunButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.CancelRunButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.CancelRunButton.Location = New System.Drawing.Point(508, 582)
      Me.CancelRunButton.Name = "CancelRunButton"
      Me.CancelRunButton.Size = New System.Drawing.Size(148, 44)
      Me.CancelRunButton.TabIndex = 28
      Me.CancelRunButton.Text = "CANCEL"
      Me.CancelRunButton.UseVisualStyleBackColor = False
      '
      'RunProgressLabel
      '
      Me.RunProgressLabel.AutoSize = True
      Me.RunProgressLabel.BackColor = System.Drawing.Color.Yellow
      Me.RunProgressLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RunProgressLabel.Location = New System.Drawing.Point(376, 486)
      Me.RunProgressLabel.Name = "RunProgressLabel"
      Me.RunProgressLabel.Size = New System.Drawing.Size(106, 20)
      Me.RunProgressLabel.TabIndex = 29
      Me.RunProgressLabel.Text = "Run Progress"
      '
      'MRProgressBar
      '
      Me.MRProgressBar.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.MRProgressBar.Cursor = System.Windows.Forms.Cursors.IBeam
      Me.MRProgressBar.Location = New System.Drawing.Point(280, 535)
      Me.MRProgressBar.Name = "MRProgressBar"
      Me.MRProgressBar.Size = New System.Drawing.Size(315, 22)
      Me.MRProgressBar.TabIndex = 30
      '
      'MSFBias
      '
      Me.MSFBias.AutoSize = True
      Me.MSFBias.BackColor = System.Drawing.Color.Fuchsia
      Me.MSFBias.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.MSFBias.Location = New System.Drawing.Point(16, 645)
      Me.MSFBias.Name = "MSFBias"
      Me.MSFBias.Size = New System.Drawing.Size(210, 21)
      Me.MSFBias.TabIndex = 31
      Me.MSFBias.Text = "Use MSF bias adjustment"
      Me.MSFBias.UseVisualStyleBackColor = False
      '
      'FVS_RunModel
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(882, 678)
      Me.Controls.Add(Me.MSFBias)
      Me.Controls.Add(Me.MRProgressBar)
      Me.Controls.Add(Me.RunProgressLabel)
      Me.Controls.Add(Me.CancelRunButton)
      Me.Controls.Add(Me.RunModelButton)
      Me.Controls.Add(Me.ChinookBYCheck)
      Me.Controls.Add(Me.TammFwsCheck)
      Me.Controls.Add(Me.OldTammCheck)
      Me.Controls.Add(Me.ReplaceQuotaCheck)
      Me.Controls.Add(Me.SelectTAMMButton)
      Me.Controls.Add(Me.TammNameLabel)
      Me.Controls.Add(Me.TammTextLabel)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.ModelRunTitleLabel)
      Me.Name = "FVS_RunModel"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_ModelRun"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents ModelRunTitleLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents TammTextLabel As System.Windows.Forms.Label
   Friend WithEvents TammNameLabel As System.Windows.Forms.Label
   Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
   Friend WithEvents SelectTAMMButton As System.Windows.Forms.Button
   Friend WithEvents ReplaceQuotaCheck As System.Windows.Forms.CheckBox
   Friend WithEvents OldTammCheck As System.Windows.Forms.CheckBox
   Friend WithEvents TammFwsCheck As System.Windows.Forms.CheckBox
   Friend WithEvents ChinookBYCheck As System.Windows.Forms.CheckBox
   Friend WithEvents RunModelButton As System.Windows.Forms.Button
   Friend WithEvents CancelRunButton As System.Windows.Forms.Button
   Friend WithEvents RunProgressLabel As System.Windows.Forms.Label
   Friend WithEvents MRProgressBar As System.Windows.Forms.ProgressBar
   Friend WithEvents MSFBias As System.Windows.Forms.CheckBox
End Class
