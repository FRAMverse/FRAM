<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_SelectiveFisheryScreen
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
      Me.MSFGrid = New System.Windows.Forms.DataGridView()
      Me.MSFTitleLabel = New System.Windows.Forms.Label()
      Me.MSFComboBox = New System.Windows.Forms.ComboBox()
      Me.MSFSelectedLabel = New System.Windows.Forms.Label()
      Me.MSFExitButton = New System.Windows.Forms.Button()
      Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
      Me.ClipBoardCopyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.btn_MSFreport = New System.Windows.Forms.Button()
      CType(Me.MSFGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.MenuStrip1.SuspendLayout()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(453, 32)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(247, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Selective Fishery Impacts"
      '
      'MSFGrid
      '
      Me.MSFGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.MSFGrid.Location = New System.Drawing.Point(34, 120)
      Me.MSFGrid.Name = "MSFGrid"
      Me.MSFGrid.RowTemplate.Height = 24
      Me.MSFGrid.Size = New System.Drawing.Size(1170, 613)
      Me.MSFGrid.TabIndex = 1
      '
      'MSFTitleLabel
      '
      Me.MSFTitleLabel.AutoSize = True
      Me.MSFTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.MSFTitleLabel.Location = New System.Drawing.Point(33, 49)
      Me.MSFTitleLabel.Name = "MSFTitleLabel"
      Me.MSFTitleLabel.Size = New System.Drawing.Size(182, 13)
      Me.MSFTitleLabel.TabIndex = 13
      Me.MSFTitleLabel.Text = "Choose Mark Selective Fishery"
      '
      'MSFComboBox
      '
      Me.MSFComboBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.MSFComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.MSFComboBox.FormattingEnabled = True
      Me.MSFComboBox.Location = New System.Drawing.Point(36, 66)
      Me.MSFComboBox.Name = "MSFComboBox"
      Me.MSFComboBox.Size = New System.Drawing.Size(274, 25)
      Me.MSFComboBox.TabIndex = 12
      '
      'MSFSelectedLabel
      '
      Me.MSFSelectedLabel.AutoSize = True
      Me.MSFSelectedLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.MSFSelectedLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.MSFSelectedLabel.Location = New System.Drawing.Point(316, 74)
      Me.MSFSelectedLabel.Name = "MSFSelectedLabel"
      Me.MSFSelectedLabel.Size = New System.Drawing.Size(112, 17)
      Me.MSFSelectedLabel.TabIndex = 10
      Me.MSFSelectedLabel.Text = "MSF-Selection"
      '
      'MSFExitButton
      '
      Me.MSFExitButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.MSFExitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.MSFExitButton.Location = New System.Drawing.Point(520, 753)
      Me.MSFExitButton.Name = "MSFExitButton"
      Me.MSFExitButton.Size = New System.Drawing.Size(168, 42)
      Me.MSFExitButton.TabIndex = 14
      Me.MSFExitButton.Text = "EXIT"
      Me.MSFExitButton.UseVisualStyleBackColor = False
      '
      'MenuStrip1
      '
      Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClipBoardCopyToolStripMenuItem})
      Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
      Me.MenuStrip1.Name = "MenuStrip1"
      Me.MenuStrip1.Size = New System.Drawing.Size(1204, 24)
      Me.MenuStrip1.TabIndex = 15
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
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(96, 865)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(96, 831)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(13, 865)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(13, 831)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'btn_MSFreport
      '
      Me.btn_MSFreport.BackColor = System.Drawing.Color.Salmon
      Me.btn_MSFreport.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.btn_MSFreport.Location = New System.Drawing.Point(54, 753)
      Me.btn_MSFreport.Name = "btn_MSFreport"
      Me.btn_MSFreport.Size = New System.Drawing.Size(175, 46)
      Me.btn_MSFreport.TabIndex = 31
      Me.btn_MSFreport.Text = "WDFW MSF Report"
      Me.btn_MSFreport.UseVisualStyleBackColor = False
      '
      'FVS_SelectiveFisheryScreen
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(1216, 807)
      Me.Controls.Add(Me.btn_MSFreport)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.MSFExitButton)
      Me.Controls.Add(Me.MSFTitleLabel)
      Me.Controls.Add(Me.MSFComboBox)
      Me.Controls.Add(Me.MSFSelectedLabel)
      Me.Controls.Add(Me.MSFGrid)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.MenuStrip1)
      Me.MainMenuStrip = Me.MenuStrip1
      Me.Name = "FVS_SelectiveFisheryScreen"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_SelectiveFisheryScreen"
      CType(Me.MSFGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.MenuStrip1.ResumeLayout(False)
      Me.MenuStrip1.PerformLayout()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents MSFGrid As System.Windows.Forms.DataGridView
   Friend WithEvents MSFTitleLabel As System.Windows.Forms.Label
   Friend WithEvents MSFComboBox As System.Windows.Forms.ComboBox
   Friend WithEvents MSFSelectedLabel As System.Windows.Forms.Label
   Friend WithEvents MSFExitButton As System.Windows.Forms.Button
   Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
   Friend WithEvents ClipBoardCopyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents btn_MSFreport As System.Windows.Forms.Button
End Class
