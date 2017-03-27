<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_SizeLimitEdit
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
      Me.SizeLimitGrid = New System.Windows.Forms.DataGridView()
      Me.SLDoneButton = New System.Windows.Forms.Button()
      Me.SLCancelButton = New System.Windows.Forms.Button()
      Me.SizeLimitTitle = New System.Windows.Forms.Label()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.ClipBoardCopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
      Me.btnLimitChange = New System.Windows.Forms.Button()
      Me.SizeLimitBox = New System.Windows.Forms.CheckBox()
      CType(Me.SizeLimitGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'SizeLimitGrid
      '
      Me.SizeLimitGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.SizeLimitGrid.Location = New System.Drawing.Point(30, 68)
      Me.SizeLimitGrid.Name = "SizeLimitGrid"
      Me.SizeLimitGrid.RowTemplate.Height = 24
      Me.SizeLimitGrid.Size = New System.Drawing.Size(839, 600)
      Me.SizeLimitGrid.TabIndex = 0
      '
      'SLDoneButton
      '
      Me.SLDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SLDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SLDoneButton.Location = New System.Drawing.Point(158, 693)
      Me.SLDoneButton.Name = "SLDoneButton"
      Me.SLDoneButton.Size = New System.Drawing.Size(170, 44)
      Me.SLDoneButton.TabIndex = 1
      Me.SLDoneButton.Text = "OK - Done"
      Me.SLDoneButton.UseVisualStyleBackColor = False
      '
      'SLCancelButton
      '
      Me.SLCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SLCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SLCancelButton.Location = New System.Drawing.Point(384, 693)
      Me.SLCancelButton.Name = "SLCancelButton"
      Me.SLCancelButton.Size = New System.Drawing.Size(170, 44)
      Me.SLCancelButton.TabIndex = 2
      Me.SLCancelButton.Text = "Cancel"
      Me.SLCancelButton.UseVisualStyleBackColor = False
      '
      'SizeLimitTitle
      '
      Me.SizeLimitTitle.AutoSize = True
      Me.SizeLimitTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SizeLimitTitle.Location = New System.Drawing.Point(341, 22)
      Me.SizeLimitTitle.Name = "SizeLimitTitle"
      Me.SizeLimitTitle.Size = New System.Drawing.Size(184, 24)
      Me.SizeLimitTitle.TabIndex = 3
      Me.SizeLimitTitle.Text = "Fishery Size Limits"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(108, 802)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(108, 768)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(25, 802)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(25, 768)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipBoardCopyToolStripMenuItem})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Size = New System.Drawing.Size(902, 24)
      Me.MenuStrip1.TabIndex = 31
      Me.MenuStrip1.Text = "MenuStrip1"
      '
      'ClipBoardCopyToolStripMenuItem
      '
      Me.ClipBoardCopyToolStripMenuItem.Name = "ClipBoardCopyToolStripMenuItem"
      Me.ClipBoardCopyToolStripMenuItem.Size = New System.Drawing.Size(102, 20)
      Me.ClipBoardCopyToolStripMenuItem.Text = "ClipBoard Copy"
      '
      'btnLimitChange
      '
      Me.btnLimitChange.BackColor = System.Drawing.Color.Plum
      Me.btnLimitChange.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.btnLimitChange.Location = New System.Drawing.Point(608, 693)
      Me.btnLimitChange.Name = "btnLimitChange"
      Me.btnLimitChange.Size = New System.Drawing.Size(170, 44)
      Me.btnLimitChange.TabIndex = 32
      Me.btnLimitChange.Text = "Load Limit Changes"
      Me.btnLimitChange.UseVisualStyleBackColor = False
      '
      'SizeLimitBox
      '
      Me.SizeLimitBox.AutoSize = True
      Me.SizeLimitBox.Font = New System.Drawing.Font("Arial", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SizeLimitBox.Location = New System.Drawing.Point(504, 759)
      Me.SizeLimitBox.Name = "SizeLimitBox"
      Me.SizeLimitBox.Size = New System.Drawing.Size(302, 22)
      Me.SizeLimitBox.TabIndex = 33
      Me.SizeLimitBox.Text = "Alternative Size Limits (check for yes)?"
      Me.SizeLimitBox.UseVisualStyleBackColor = True
      '
      'FVS_SizeLimitEdit
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(902, 822)
      Me.Controls.Add(Me.SizeLimitBox)
      Me.Controls.Add(Me.btnLimitChange)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.SizeLimitTitle)
      Me.Controls.Add(Me.SLCancelButton)
      Me.Controls.Add(Me.SLDoneButton)
      Me.Controls.Add(Me.SizeLimitGrid)
      Me.Controls.Add(Me.MenuStrip1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Name = "FVS_SizeLimitEdit"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Size Limit Edit"
      CType(Me.SizeLimitGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents SizeLimitGrid As System.Windows.Forms.DataGridView
   Friend WithEvents SLDoneButton As System.Windows.Forms.Button
   Friend WithEvents SLCancelButton As System.Windows.Forms.Button
   Friend WithEvents SizeLimitTitle As System.Windows.Forms.Label
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipBoardCopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents btnLimitChange As System.Windows.Forms.Button
   Friend WithEvents SizeLimitBox As System.Windows.Forms.CheckBox
End Class
