<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_BackwardsResults
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
        Me.BROKButton = New System.Windows.Forms.Button()
        Me.BFResultsGrid = New System.Windows.Forms.DataGridView()
        Me.RecordSetNameLabel = New System.Windows.Forms.Label()
        Me.DatabaseNameLabel = New System.Windows.Forms.Label()
        Me.RecordSetTextLabel = New System.Windows.Forms.Label()
        Me.DatabaseTextLabel = New System.Windows.Forms.Label()
        CType(Me.BFResultsGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(445, 20)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(322, 24)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Results for Backwards FRAM Run"
        '
        'BROKButton
        '
        Me.BROKButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BROKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BROKButton.Location = New System.Drawing.Point(551, 780)
        Me.BROKButton.Name = "BROKButton"
        Me.BROKButton.Size = New System.Drawing.Size(159, 45)
        Me.BROKButton.TabIndex = 3
        Me.BROKButton.Text = "OK - Done"
        Me.BROKButton.UseVisualStyleBackColor = False
        '
        'BFResultsGrid
        '
        Me.BFResultsGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.BFResultsGrid.Location = New System.Drawing.Point(12, 64)
        Me.BFResultsGrid.Name = "BFResultsGrid"
        Me.BFResultsGrid.RowTemplate.Height = 24
        Me.BFResultsGrid.Size = New System.Drawing.Size(1189, 691)
        Me.BFResultsGrid.TabIndex = 4
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(98, 814)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(98, 780)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 18
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(15, 814)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 17
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(15, 780)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 16
        Me.DatabaseTextLabel.Text = "Database"
        '
        'FVS_BackwardsResults
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1213, 837)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.BFResultsGrid)
        Me.Controls.Add(Me.BROKButton)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FVS_BackwardsResults"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FVS_BackwardsResults"
        CType(Me.BFResultsGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents BROKButton As System.Windows.Forms.Button
   Friend WithEvents BFResultsGrid As System.Windows.Forms.DataGridView
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
