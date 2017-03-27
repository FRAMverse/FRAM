<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_StockImpactsPer1000Screen
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
      Me.SIPExitButton = New System.Windows.Forms.Button()
      Me.SIPGrid = New System.Windows.Forms.DataGridView()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.ClipBoardCopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
      Me.FSCTitleLabel = New System.Windows.Forms.Label()
      Me.SIPSelectedLabel = New System.Windows.Forms.Label()
      Me.SIPComboBox = New System.Windows.Forms.ComboBox()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      CType(Me.SIPGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(250, 41)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(375, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Stock Impacts Per 1000 (Total Mortality)"
      '
      'SIPExitButton
      '
      Me.SIPExitButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SIPExitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SIPExitButton.Location = New System.Drawing.Point(495, 753)
      Me.SIPExitButton.Name = "SIPExitButton"
      Me.SIPExitButton.Size = New System.Drawing.Size(140, 41)
      Me.SIPExitButton.TabIndex = 1
      Me.SIPExitButton.Text = "EXIT"
      Me.SIPExitButton.UseVisualStyleBackColor = False
      '
      'SIPGrid
      '
      Me.SIPGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.SIPGrid.Location = New System.Drawing.Point(12, 128)
      Me.SIPGrid.Name = "SIPGrid"
      Me.SIPGrid.RowTemplate.Height = 24
      Me.SIPGrid.Size = New System.Drawing.Size(953, 610)
      Me.SIPGrid.TabIndex = 2
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipBoardCopyToolStripMenuItem})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Size = New System.Drawing.Size(977, 24)
      Me.MenuStrip1.TabIndex = 3
      Me.MenuStrip1.Text = "MenuStrip1"
      '
      'ClipBoardCopyToolStripMenuItem
      '
      Me.ClipBoardCopyToolStripMenuItem.Name = "ClipBoardCopyToolStripMenuItem"
      Me.ClipBoardCopyToolStripMenuItem.Size = New System.Drawing.Size(102, 20)
      Me.ClipBoardCopyToolStripMenuItem.Text = "ClipBoard Copy"
      '
      'FSCTitleLabel
      '
      Me.FSCTitleLabel.AutoSize = True
      Me.FSCTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCTitleLabel.Location = New System.Drawing.Point(9, 66)
      Me.FSCTitleLabel.Name = "FSCTitleLabel"
      Me.FSCTitleLabel.Size = New System.Drawing.Size(86, 13)
      Me.FSCTitleLabel.TabIndex = 19
      Me.FSCTitleLabel.Text = "Choose Stock"
      '
      'SIPSelectedLabel
      '
      Me.SIPSelectedLabel.AutoSize = True
      Me.SIPSelectedLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.SIPSelectedLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SIPSelectedLabel.Location = New System.Drawing.Point(392, 91)
      Me.SIPSelectedLabel.Name = "SIPSelectedLabel"
      Me.SIPSelectedLabel.Size = New System.Drawing.Size(121, 17)
      Me.SIPSelectedLabel.TabIndex = 17
      Me.SIPSelectedLabel.Text = "Stock-Selection"
      '
      'SIPComboBox
      '
      Me.SIPComboBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SIPComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SIPComboBox.FormattingEnabled = True
      Me.SIPComboBox.Location = New System.Drawing.Point(12, 86)
      Me.SIPComboBox.Name = "SIPComboBox"
      Me.SIPComboBox.Size = New System.Drawing.Size(374, 25)
      Me.SIPComboBox.TabIndex = 20
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label2.Location = New System.Drawing.Point(12, 766)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(398, 17)
      Me.Label2.TabIndex = 21
      Me.Label2.Text = " (-----) No Fishery    (*****) No Stock Impact in Fishery"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(96, 849)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(96, 815)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(13, 849)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(13, 815)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_StockImpactsPer1000Screen
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(977, 867)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.SIPComboBox)
      Me.Controls.Add(Me.FSCTitleLabel)
      Me.Controls.Add(Me.SIPSelectedLabel)
      Me.Controls.Add(Me.SIPGrid)
      Me.Controls.Add(Me.SIPExitButton)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.MenuStrip1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Name = "FVS_StockImpactsPer1000Screen"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_StockImpactsPer1000Screen"
      CType(Me.SIPGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents SIPExitButton As System.Windows.Forms.Button
   Friend WithEvents SIPGrid As System.Windows.Forms.DataGridView
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipBoardCopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents FSCTitleLabel As System.Windows.Forms.Label
   Friend WithEvents SIPSelectedLabel As System.Windows.Forms.Label
   Friend WithEvents SIPComboBox As System.Windows.Forms.ComboBox
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
