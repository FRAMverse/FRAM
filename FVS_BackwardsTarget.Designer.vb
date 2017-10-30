<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_BackwardsTarget
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.Label1 = New System.Windows.Forms.Label
        Me.BFTargetGrid = New System.Windows.Forms.DataGridView
        Me.BTOKButton = New System.Windows.Forms.Button
        Me.BTCancelButton = New System.Windows.Forms.Button
        Me.BTEscapementButton = New System.Windows.Forms.Button
        Me.BTFillSSButton = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.RecordSetNameLabel = New System.Windows.Forms.Label
        Me.DatabaseNameLabel = New System.Windows.Forms.Label
        Me.RecordSetTextLabel = New System.Windows.Forms.Label
        Me.DatabaseTextLabel = New System.Windows.Forms.Label
        Me.BTCatchButton = New System.Windows.Forms.Button
        CType(Me.BFTargetGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(216, 25)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(402, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Target Escapements for Backwards FRAM"
        '
        'BFTargetGrid
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.BFTargetGrid.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.BFTargetGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.BFTargetGrid.Location = New System.Drawing.Point(44, 74)
        Me.BFTargetGrid.Name = "BFTargetGrid"
        Me.BFTargetGrid.RowTemplate.Height = 24
        Me.BFTargetGrid.Size = New System.Drawing.Size(935, 607)
        Me.BFTargetGrid.TabIndex = 1
        '
        'BTOKButton
        '
        Me.BTOKButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BTOKButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTOKButton.Location = New System.Drawing.Point(44, 704)
        Me.BTOKButton.Name = "BTOKButton"
        Me.BTOKButton.Size = New System.Drawing.Size(159, 45)
        Me.BTOKButton.TabIndex = 2
        Me.BTOKButton.Text = "OK - Done"
        Me.BTOKButton.UseVisualStyleBackColor = False
        '
        'BTCancelButton
        '
        Me.BTCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BTCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTCancelButton.Location = New System.Drawing.Point(221, 704)
        Me.BTCancelButton.Name = "BTCancelButton"
        Me.BTCancelButton.Size = New System.Drawing.Size(159, 45)
        Me.BTCancelButton.TabIndex = 3
        Me.BTCancelButton.Text = "Cancel"
        Me.BTCancelButton.UseVisualStyleBackColor = False
        '
        'BTEscapementButton
        '
        Me.BTEscapementButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BTEscapementButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTEscapementButton.Location = New System.Drawing.Point(410, 704)
        Me.BTEscapementButton.Name = "BTEscapementButton"
        Me.BTEscapementButton.Size = New System.Drawing.Size(208, 61)
        Me.BTEscapementButton.TabIndex = 4
        Me.BTEscapementButton.Text = "Import Escapements"
        Me.BTEscapementButton.UseVisualStyleBackColor = False
        '
        'BTFillSSButton
        '
        Me.BTFillSSButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BTFillSSButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTFillSSButton.Location = New System.Drawing.Point(653, 704)
        Me.BTFillSSButton.Name = "BTFillSSButton"
        Me.BTFillSSButton.Size = New System.Drawing.Size(208, 61)
        Me.BTFillSSButton.TabIndex = 5
        Me.BTFillSSButton.Text = "Export to Spreadsheet"
        Me.BTFillSSButton.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(43, 768)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(456, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "FLAGS: 0=Don't Use, 1=Exact Value, 2=Split into M/UM" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(128, 834)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(128, 800)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 18
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(45, 834)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 17
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(45, 800)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 16
        Me.DatabaseTextLabel.Text = "Database"
        '
        'BTCatchButton
        '
        Me.BTCatchButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.BTCatchButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BTCatchButton.Location = New System.Drawing.Point(784, 783)
        Me.BTCatchButton.Name = "BTCatchButton"
        Me.BTCatchButton.Size = New System.Drawing.Size(208, 45)
        Me.BTCatchButton.TabIndex = 20
        Me.BTCatchButton.Text = "Load Back-Fram Catch"
        Me.BTCatchButton.UseVisualStyleBackColor = False
        '
        'FVS_BackwardsTarget
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1004, 855)
        Me.Controls.Add(Me.BTCatchButton)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.BTFillSSButton)
        Me.Controls.Add(Me.BTEscapementButton)
        Me.Controls.Add(Me.BTCancelButton)
        Me.Controls.Add(Me.BTOKButton)
        Me.Controls.Add(Me.BFTargetGrid)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FVS_BackwardsTarget"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FVS_BackwardsTarget"
        CType(Me.BFTargetGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents BFTargetGrid As System.Windows.Forms.DataGridView
   Friend WithEvents BTOKButton As System.Windows.Forms.Button
   Friend WithEvents BTCancelButton As System.Windows.Forms.Button
   Friend WithEvents BTEscapementButton As System.Windows.Forms.Button
   Friend WithEvents BTFillSSButton As System.Windows.Forms.Button
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents BTCatchButton As System.Windows.Forms.Button
End Class
