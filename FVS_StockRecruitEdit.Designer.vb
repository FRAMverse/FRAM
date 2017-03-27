<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_StockRecruitEdit
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
        Me.StockRecruitGrid = New System.Windows.Forms.DataGridView()
        Me.SRDoneButton = New System.Windows.Forms.Button()
        Me.SRCancelButton = New System.Windows.Forms.Button()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.RecordSetNameLabel = New System.Windows.Forms.Label()
        Me.DatabaseNameLabel = New System.Windows.Forms.Label()
        Me.RecordSetTextLabel = New System.Windows.Forms.Label()
        Me.DatabaseTextLabel = New System.Windows.Forms.Label()
        Me.ReadRecruitsButton = New System.Windows.Forms.Button()
        Me.FillRecruitSSButton = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        CType(Me.StockRecruitGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(368, 54)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(234, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Stock Recruit (Forecast)"
        '
        'StockRecruitGrid
        '
        Me.StockRecruitGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.StockRecruitGrid.Location = New System.Drawing.Point(22, 96)
        Me.StockRecruitGrid.Name = "StockRecruitGrid"
        Me.StockRecruitGrid.RowTemplate.Height = 24
        Me.StockRecruitGrid.Size = New System.Drawing.Size(997, 535)
        Me.StockRecruitGrid.TabIndex = 1
        '
        'SRDoneButton
        '
        Me.SRDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SRDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SRDoneButton.Location = New System.Drawing.Point(141, 655)
        Me.SRDoneButton.Name = "SRDoneButton"
        Me.SRDoneButton.Size = New System.Drawing.Size(172, 40)
        Me.SRDoneButton.TabIndex = 2
        Me.SRDoneButton.Text = "OK - Done"
        Me.SRDoneButton.UseVisualStyleBackColor = False
        '
        'SRCancelButton
        '
        Me.SRCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SRCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SRCancelButton.Location = New System.Drawing.Point(319, 655)
        Me.SRCancelButton.Name = "SRCancelButton"
        Me.SRCancelButton.Size = New System.Drawing.Size(172, 40)
        Me.SRCancelButton.TabIndex = 3
        Me.SRCancelButton.Text = "Cancel Changes"
        Me.SRCancelButton.UseVisualStyleBackColor = False
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1031, 24)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(147, 20)
        Me.ToolStripMenuItem1.Text = "Copy-Grid-to-ClipBoard"
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(107, 765)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(107, 731)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 29
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(24, 765)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 28
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(24, 731)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 27
        Me.DatabaseTextLabel.Text = "Database"
        '
        'ReadRecruitsButton
        '
        Me.ReadRecruitsButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ReadRecruitsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReadRecruitsButton.Location = New System.Drawing.Point(562, 655)
        Me.ReadRecruitsButton.Name = "ReadRecruitsButton"
        Me.ReadRecruitsButton.Size = New System.Drawing.Size(172, 40)
        Me.ReadRecruitsButton.TabIndex = 31
        Me.ReadRecruitsButton.Text = "Read Recruits"
        Me.ReadRecruitsButton.UseVisualStyleBackColor = False
        '
        'FillRecruitSSButton
        '
        Me.FillRecruitSSButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.FillRecruitSSButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FillRecruitSSButton.Location = New System.Drawing.Point(753, 655)
        Me.FillRecruitSSButton.Name = "FillRecruitSSButton"
        Me.FillRecruitSSButton.Size = New System.Drawing.Size(172, 40)
        Me.FillRecruitSSButton.TabIndex = 32
        Me.FillRecruitSSButton.Text = "Fill Spreadsheet"
        Me.FillRecruitSSButton.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'FVS_StockRecruitEdit
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1031, 791)
        Me.Controls.Add(Me.FillRecruitSSButton)
        Me.Controls.Add(Me.ReadRecruitsButton)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.SRCancelButton)
        Me.Controls.Add(Me.SRDoneButton)
        Me.Controls.Add(Me.StockRecruitGrid)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FVS_StockRecruitEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit StockRecruit Parameters"
        CType(Me.StockRecruitGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents StockRecruitGrid As System.Windows.Forms.DataGridView
   Friend WithEvents SRDoneButton As System.Windows.Forms.Button
   Friend WithEvents SRCancelButton As System.Windows.Forms.Button
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents ReadRecruitsButton As System.Windows.Forms.Button
   Friend WithEvents FillRecruitSSButton As System.Windows.Forms.Button
   Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
End Class
