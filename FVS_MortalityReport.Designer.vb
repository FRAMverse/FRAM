<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_MortalityReport
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
        Me.CRCancelButton = New System.Windows.Forms.Button()
        Me.MortalityGrid = New System.Windows.Forms.DataGridView()
        Me.MortalityReportTitle = New System.Windows.Forms.Label()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.ClipBoardCopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StocksSelectedLabel = New System.Windows.Forms.Label()
        Me.StockTitleLabel = New System.Windows.Forms.Label()
        Me.MortalityTypeComboBox = New System.Windows.Forms.ComboBox()
        Me.MortalityTypeTitleLabel = New System.Windows.Forms.Label()
        Me.AgeSumButton = New System.Windows.Forms.Button()
        Me.RecordSetNameLabel = New System.Windows.Forms.Label()
        Me.DatabaseNameLabel = New System.Windows.Forms.Label()
        Me.RecordSetTextLabel = New System.Windows.Forms.Label()
        Me.DatabaseTextLabel = New System.Windows.Forms.Label()
        CType(Me.MortalityGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'CRCancelButton
        '
        Me.CRCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.CRCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CRCancelButton.Location = New System.Drawing.Point(395, 801)
        Me.CRCancelButton.Name = "CRCancelButton"
        Me.CRCancelButton.Size = New System.Drawing.Size(168, 42)
        Me.CRCancelButton.TabIndex = 1
        Me.CRCancelButton.Text = "EXIT"
        Me.CRCancelButton.UseVisualStyleBackColor = False
        '
        'MortalityGrid
        '
        Me.MortalityGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.MortalityGrid.Location = New System.Drawing.Point(29, 118)
        Me.MortalityGrid.Name = "MortalityGrid"
        Me.MortalityGrid.RowTemplate.Height = 24
        Me.MortalityGrid.Size = New System.Drawing.Size(895, 656)
        Me.MortalityGrid.TabIndex = 2
        '
        'MortalityReportTitle
        '
        Me.MortalityReportTitle.AutoSize = True
        Me.MortalityReportTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MortalityReportTitle.Location = New System.Drawing.Point(393, 33)
        Me.MortalityReportTitle.Name = "MortalityReportTitle"
        Me.MortalityReportTitle.Size = New System.Drawing.Size(139, 24)
        Me.MortalityReportTitle.TabIndex = 3
        Me.MortalityReportTitle.Text = "Landed Catch"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipBoardCopyToolStripMenuItem, Me.PrintToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(959, 24)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'ClipBoardCopyToolStripMenuItem
        '
        Me.ClipBoardCopyToolStripMenuItem.Name = "ClipBoardCopyToolStripMenuItem"
        Me.ClipBoardCopyToolStripMenuItem.Size = New System.Drawing.Size(99, 20)
        Me.ClipBoardCopyToolStripMenuItem.Text = "ClipBoardCopy"
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.PrintToolStripMenuItem.Text = "Print"
        Me.PrintToolStripMenuItem.Visible = False
        '
        'StocksSelectedLabel
        '
        Me.StocksSelectedLabel.AutoSize = True
        Me.StocksSelectedLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.StocksSelectedLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StocksSelectedLabel.Location = New System.Drawing.Point(206, 78)
        Me.StocksSelectedLabel.Name = "StocksSelectedLabel"
        Me.StocksSelectedLabel.Size = New System.Drawing.Size(56, 17)
        Me.StocksSelectedLabel.TabIndex = 5
        Me.StocksSelectedLabel.Text = "Stocks"
        '
        'StockTitleLabel
        '
        Me.StockTitleLabel.AutoSize = True
        Me.StockTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StockTitleLabel.Location = New System.Drawing.Point(207, 61)
        Me.StockTitleLabel.Name = "StockTitleLabel"
        Me.StockTitleLabel.Size = New System.Drawing.Size(100, 13)
        Me.StockTitleLabel.TabIndex = 6
        Me.StockTitleLabel.Text = "Stocks Selected"
        '
        'MortalityTypeComboBox
        '
        Me.MortalityTypeComboBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.MortalityTypeComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MortalityTypeComboBox.FormattingEnabled = True
        Me.MortalityTypeComboBox.Items.AddRange(New Object() {"None Selected", "Landed Catch", "Non Retention", "Shakers", "Legal Shakers", "DropOff", "Total Mortality", "AEQ Total Mortality"})
        Me.MortalityTypeComboBox.Location = New System.Drawing.Point(29, 78)
        Me.MortalityTypeComboBox.Name = "MortalityTypeComboBox"
        Me.MortalityTypeComboBox.Size = New System.Drawing.Size(171, 25)
        Me.MortalityTypeComboBox.TabIndex = 8
        '
        'MortalityTypeTitleLabel
        '
        Me.MortalityTypeTitleLabel.AutoSize = True
        Me.MortalityTypeTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MortalityTypeTitleLabel.Location = New System.Drawing.Point(26, 61)
        Me.MortalityTypeTitleLabel.Name = "MortalityTypeTitleLabel"
        Me.MortalityTypeTitleLabel.Size = New System.Drawing.Size(87, 13)
        Me.MortalityTypeTitleLabel.TabIndex = 9
        Me.MortalityTypeTitleLabel.Text = "Mortality Type"
        '
        'AgeSumButton
        '
        Me.AgeSumButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.AgeSumButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AgeSumButton.Location = New System.Drawing.Point(659, 801)
        Me.AgeSumButton.Name = "AgeSumButton"
        Me.AgeSumButton.Size = New System.Drawing.Size(168, 42)
        Me.AgeSumButton.TabIndex = 10
        Me.AgeSumButton.Text = "Sum Age Only"
        Me.AgeSumButton.UseVisualStyleBackColor = False
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(111, 902)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(111, 868)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 29
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(28, 902)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 28
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(28, 868)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 27
        Me.DatabaseTextLabel.Text = "Database"
        '
        'FVS_MortalityReport
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(959, 926)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.AgeSumButton)
        Me.Controls.Add(Me.MortalityTypeTitleLabel)
        Me.Controls.Add(Me.MortalityTypeComboBox)
        Me.Controls.Add(Me.StockTitleLabel)
        Me.Controls.Add(Me.StocksSelectedLabel)
        Me.Controls.Add(Me.MortalityReportTitle)
        Me.Controls.Add(Me.MortalityGrid)
        Me.Controls.Add(Me.CRCancelButton)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FVS_MortalityReport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Mortality Report"
        CType(Me.MortalityGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
   Friend WithEvents CRCancelButton As System.Windows.Forms.Button
   Friend WithEvents MortalityGrid As System.Windows.Forms.DataGridView
   Friend WithEvents MortalityReportTitle As System.Windows.Forms.Label
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipBoardCopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents StocksSelectedLabel As System.Windows.Forms.Label
   Friend WithEvents StockTitleLabel As System.Windows.Forms.Label
   Friend WithEvents MortalityTypeComboBox As System.Windows.Forms.ComboBox
   Friend WithEvents MortalityTypeTitleLabel As System.Windows.Forms.Label
   Friend WithEvents AgeSumButton As System.Windows.Forms.Button
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
