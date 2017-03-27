<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_StockFisheryScalerEdit
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
      Me.SFDoneButton = New System.Windows.Forms.Button()
      Me.SFCancelButton = New System.Windows.Forms.Button()
      Me.StockFisheryGrid = New System.Windows.Forms.DataGridView()
      Me.SFTitle = New System.Windows.Forms.Label()
      Me.SFSTitleLabel = New System.Windows.Forms.Label()
      Me.SFSComboBox = New System.Windows.Forms.ComboBox()
      Me.SFSSelectedLabel = New System.Windows.Forms.Label()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      CType(Me.StockFisheryGrid, System.ComponentModel.ISupportInitialize).BeginInit()
      Me.SuspendLayout()
      '
      'SFDoneButton
      '
      Me.SFDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SFDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFDoneButton.Location = New System.Drawing.Point(303, 720)
      Me.SFDoneButton.Name = "SFDoneButton"
      Me.SFDoneButton.Size = New System.Drawing.Size(145, 42)
      Me.SFDoneButton.TabIndex = 0
      Me.SFDoneButton.Text = "OK - Done"
      Me.SFDoneButton.UseVisualStyleBackColor = False
      '
      'SFCancelButton
      '
      Me.SFCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SFCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFCancelButton.Location = New System.Drawing.Point(524, 720)
      Me.SFCancelButton.Name = "SFCancelButton"
      Me.SFCancelButton.Size = New System.Drawing.Size(145, 42)
      Me.SFCancelButton.TabIndex = 1
      Me.SFCancelButton.Text = "Cancel"
      Me.SFCancelButton.UseVisualStyleBackColor = False
      '
      'StockFisheryGrid
      '
      Me.StockFisheryGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
      Me.StockFisheryGrid.Location = New System.Drawing.Point(27, 100)
      Me.StockFisheryGrid.Name = "StockFisheryGrid"
      Me.StockFisheryGrid.RowTemplate.Height = 24
      Me.StockFisheryGrid.Size = New System.Drawing.Size(928, 604)
      Me.StockFisheryGrid.TabIndex = 2
      '
      'SFTitle
      '
      Me.SFTitle.AutoSize = True
      Me.SFTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFTitle.Location = New System.Drawing.Point(203, 18)
      Me.SFTitle.Name = "SFTitle"
      Me.SFTitle.Size = New System.Drawing.Size(444, 24)
      Me.SFTitle.TabIndex = 3
      Me.SFTitle.Text = "Stock/Fishery Specific Exploitation Rate Scaler"
      '
      'SFSTitleLabel
      '
      Me.SFSTitleLabel.AutoSize = True
      Me.SFSTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFSTitleLabel.Location = New System.Drawing.Point(24, 40)
      Me.SFSTitleLabel.Name = "SFSTitleLabel"
      Me.SFSTitleLabel.Size = New System.Drawing.Size(93, 13)
      Me.SFSTitleLabel.TabIndex = 19
      Me.SFSTitleLabel.Text = "Choose Fishery"
      '
      'SFSComboBox
      '
      Me.SFSComboBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SFSComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFSComboBox.FormattingEnabled = True
      Me.SFSComboBox.Location = New System.Drawing.Point(27, 57)
      Me.SFSComboBox.Name = "SFSComboBox"
      Me.SFSComboBox.Size = New System.Drawing.Size(374, 25)
      Me.SFSComboBox.TabIndex = 18
      '
      'SFSSelectedLabel
      '
      Me.SFSSelectedLabel.AutoSize = True
      Me.SFSSelectedLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.SFSSelectedLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SFSSelectedLabel.Location = New System.Drawing.Point(407, 65)
      Me.SFSSelectedLabel.Name = "SFSSelectedLabel"
      Me.SFSSelectedLabel.Size = New System.Drawing.Size(134, 17)
      Me.SFSSelectedLabel.TabIndex = 17
      Me.SFSSelectedLabel.Text = "Fishery-Selection"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(103, 821)
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
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(103, 787)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(20, 821)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(20, 787)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_StockFisheryScalerEdit
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(967, 848)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.SFSTitleLabel)
      Me.Controls.Add(Me.SFSComboBox)
      Me.Controls.Add(Me.SFSSelectedLabel)
      Me.Controls.Add(Me.SFTitle)
      Me.Controls.Add(Me.StockFisheryGrid)
      Me.Controls.Add(Me.SFCancelButton)
      Me.Controls.Add(Me.SFDoneButton)
      Me.Name = "FVS_StockFisheryScalerEdit"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Stock/Fishery Scaler Edit"
      CType(Me.StockFisheryGrid, System.ComponentModel.ISupportInitialize).EndInit()
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents SFDoneButton As System.Windows.Forms.Button
   Friend WithEvents SFCancelButton As System.Windows.Forms.Button
   Friend WithEvents StockFisheryGrid As System.Windows.Forms.DataGridView
   Friend WithEvents SFTitle As System.Windows.Forms.Label
   Friend WithEvents SFSTitleLabel As System.Windows.Forms.Label
   Friend WithEvents SFSComboBox As System.Windows.Forms.ComboBox
   Friend WithEvents SFSSelectedLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
