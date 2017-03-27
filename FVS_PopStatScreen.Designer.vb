<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_PopStatScreen
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
      Me.PopStatGrid = New System.Windows.Forms.DataGridView()
      Me.PSExit = New System.Windows.Forms.Button()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.ClipBoardCopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      CType(Me.PopStatGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(481, 36)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(228, 25)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Population Statistics"
      '
      'PopStatGrid
      '
      Me.PopStatGrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.PopStatGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.PopStatGrid.Location = New System.Drawing.Point(21, 77)
      Me.PopStatGrid.Name = "PopStatGrid"
      Me.PopStatGrid.Size = New System.Drawing.Size(1150, 698)
      Me.PopStatGrid.TabIndex = 1
      '
      'PSExit
      '
      Me.PSExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.PSExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.PSExit.Location = New System.Drawing.Point(522, 792)
      Me.PSExit.Name = "PSExit"
      Me.PSExit.Size = New System.Drawing.Size(158, 36)
      Me.PSExit.TabIndex = 2
      Me.PSExit.Text = "EXIT"
      Me.PSExit.UseVisualStyleBackColor = False
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipBoardCopyToolStripMenuItem})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Size = New System.Drawing.Size(1199, 24)
      Me.MenuStrip1.TabIndex = 3
      Me.MenuStrip1.Text = "MenuStrip1"
      '
      'ClipBoardCopyToolStripMenuItem
      '
      Me.ClipBoardCopyToolStripMenuItem.Name = "ClipBoardCopyToolStripMenuItem"
      Me.ClipBoardCopyToolStripMenuItem.Size = New System.Drawing.Size(102, 20)
      Me.ClipBoardCopyToolStripMenuItem.Text = "ClipBoard Copy"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(101, 878)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(101, 844)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(18, 878)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(18, 844)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_PopStatScreen
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(1199, 900)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.PSExit)
      Me.Controls.Add(Me.PopStatGrid)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.MenuStrip1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Name = "FVS_PopStatScreen"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Population Statistics"
      CType(Me.PopStatGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents PopStatGrid As System.Windows.Forms.DataGridView
   Friend WithEvents PSExit As System.Windows.Forms.Button
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipBoardCopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
