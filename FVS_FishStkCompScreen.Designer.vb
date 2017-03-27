<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_FishStkCompScreen
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
      Me.FSCGrid = New System.Windows.Forms.DataGridView()
      Me.FSCTitleLabel = New System.Windows.Forms.Label()
      Me.FSCComboBox = New System.Windows.Forms.ComboBox()
      Me.FSCSelectedLabel = New System.Windows.Forms.Label()
      Me.FSCExitButton = New System.Windows.Forms.Button()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.ClipboardCopyToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      CType(Me.FSCGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(317, 50)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(350, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Fishery Stock Composition (Percent)"
      '
      'FSCGrid
      '
      Me.FSCGrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSCGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.FSCGrid.Location = New System.Drawing.Point(23, 147)
      Me.FSCGrid.Name = "FSCGrid"
      Me.FSCGrid.RowTemplate.Height = 24
      Me.FSCGrid.Size = New System.Drawing.Size(1037, 635)
      Me.FSCGrid.TabIndex = 1
      '
      'FSCTitleLabel
      '
      Me.FSCTitleLabel.AutoSize = True
      Me.FSCTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCTitleLabel.Location = New System.Drawing.Point(18, 87)
      Me.FSCTitleLabel.Name = "FSCTitleLabel"
      Me.FSCTitleLabel.Size = New System.Drawing.Size(93, 13)
      Me.FSCTitleLabel.TabIndex = 16
      Me.FSCTitleLabel.Text = "Choose Fishery"
      '
      'FSCComboBox
      '
      Me.FSCComboBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSCComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCComboBox.FormattingEnabled = True
      Me.FSCComboBox.Location = New System.Drawing.Point(21, 104)
      Me.FSCComboBox.Name = "FSCComboBox"
      Me.FSCComboBox.Size = New System.Drawing.Size(374, 25)
      Me.FSCComboBox.TabIndex = 15
      '
      'FSCSelectedLabel
      '
      Me.FSCSelectedLabel.AutoSize = True
      Me.FSCSelectedLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.FSCSelectedLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCSelectedLabel.Location = New System.Drawing.Point(401, 112)
      Me.FSCSelectedLabel.Name = "FSCSelectedLabel"
      Me.FSCSelectedLabel.Size = New System.Drawing.Size(134, 17)
      Me.FSCSelectedLabel.TabIndex = 14
      Me.FSCSelectedLabel.Text = "Fishery-Selection"
      '
      'FSCExitButton
      '
      Me.FSCExitButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FSCExitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FSCExitButton.Location = New System.Drawing.Point(453, 803)
      Me.FSCExitButton.Name = "FSCExitButton"
      Me.FSCExitButton.Size = New System.Drawing.Size(168, 42)
      Me.FSCExitButton.TabIndex = 17
      Me.FSCExitButton.Text = "EXIT"
      Me.FSCExitButton.UseVisualStyleBackColor = False
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipboardCopyToolStripMenuItem1})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Size = New System.Drawing.Size(1072, 24)
      Me.MenuStrip1.TabIndex = 18
      Me.MenuStrip1.Text = "MenuStrip1"
      '
      'ClipboardCopyToolStripMenuItem1
      '
      Me.ClipboardCopyToolStripMenuItem1.Name = "ClipboardCopyToolStripMenuItem1"
      Me.ClipboardCopyToolStripMenuItem1.Size = New System.Drawing.Size(102, 20)
      Me.ClipboardCopyToolStripMenuItem1.Text = "Clipboard Copy"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(112, 895)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(112, 861)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(29, 895)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(29, 861)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_FishStkCompScreen
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(1072, 930)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.MenuStrip1)
      Me.Controls.Add(Me.FSCGrid)
      Me.Controls.Add(Me.FSCTitleLabel)
      Me.Controls.Add(Me.FSCComboBox)
      Me.Controls.Add(Me.FSCExitButton)
      Me.Controls.Add(Me.FSCSelectedLabel)
      Me.Controls.Add(Me.Label1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Name = "FVS_FishStkCompScreen"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_FishStkCompScreen"
      CType(Me.FSCGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents FSCGrid As System.Windows.Forms.DataGridView
   Friend WithEvents FSCTitleLabel As System.Windows.Forms.Label
   Friend WithEvents FSCComboBox As System.Windows.Forms.ComboBox
   Friend WithEvents FSCSelectedLabel As System.Windows.Forms.Label
   Friend WithEvents FSCExitButton As System.Windows.Forms.Button
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipboardCopyToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
