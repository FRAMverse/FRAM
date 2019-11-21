<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_ScreenReports
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
        Me.Label1 = New System.Windows.Forms.Label
        Me.FisheryMortalityCheckBox = New System.Windows.Forms.CheckBox
        Me.StockCatchCheckBox = New System.Windows.Forms.CheckBox
        Me.RepCancelButton = New System.Windows.Forms.Button
        Me.FisheryScalerCheckBox = New System.Windows.Forms.CheckBox
        Me.MSFCheckBox = New System.Windows.Forms.CheckBox
        Me.FishStkCompCheckBox = New System.Windows.Forms.CheckBox
        Me.PSCCohoERCheckBox = New System.Windows.Forms.CheckBox
        Me.StockPer1000CheckBox = New System.Windows.Forms.CheckBox
        Me.PopStatCheckBox = New System.Windows.Forms.CheckBox
        Me.RecordSetNameLabel = New System.Windows.Forms.Label
        Me.DatabaseNameLabel = New System.Windows.Forms.Label
        Me.RecordSetTextLabel = New System.Windows.Forms.Label
        Me.DatabaseTextLabel = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(341, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(155, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Screen Reports"
        '
        'FisheryMortalityCheckBox
        '
        Me.FisheryMortalityCheckBox.AutoSize = True
        Me.FisheryMortalityCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FisheryMortalityCheckBox.Location = New System.Drawing.Point(35, 100)
        Me.FisheryMortalityCheckBox.Name = "FisheryMortalityCheckBox"
        Me.FisheryMortalityCheckBox.Size = New System.Drawing.Size(249, 24)
        Me.FisheryMortalityCheckBox.TabIndex = 1
        Me.FisheryMortalityCheckBox.Text = "FISHERY Mortality Reports"
        Me.FisheryMortalityCheckBox.UseVisualStyleBackColor = True
        '
        'StockCatchCheckBox
        '
        Me.StockCatchCheckBox.AutoSize = True
        Me.StockCatchCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StockCatchCheckBox.Location = New System.Drawing.Point(35, 158)
        Me.StockCatchCheckBox.Name = "StockCatchCheckBox"
        Me.StockCatchCheckBox.Size = New System.Drawing.Size(228, 24)
        Me.StockCatchCheckBox.TabIndex = 7
        Me.StockCatchCheckBox.Text = "STOCK Mortality Reports"
        Me.StockCatchCheckBox.UseVisualStyleBackColor = True
        '
        'RepCancelButton
        '
        Me.RepCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.RepCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RepCancelButton.Location = New System.Drawing.Point(392, 635)
        Me.RepCancelButton.Name = "RepCancelButton"
        Me.RepCancelButton.Size = New System.Drawing.Size(162, 45)
        Me.RepCancelButton.TabIndex = 8
        Me.RepCancelButton.Text = "EXIT"
        Me.RepCancelButton.UseVisualStyleBackColor = False
        '
        'FisheryScalerCheckBox
        '
        Me.FisheryScalerCheckBox.AutoSize = True
        Me.FisheryScalerCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FisheryScalerCheckBox.Location = New System.Drawing.Point(35, 487)
        Me.FisheryScalerCheckBox.Name = "FisheryScalerCheckBox"
        Me.FisheryScalerCheckBox.Size = New System.Drawing.Size(202, 24)
        Me.FisheryScalerCheckBox.TabIndex = 10
        Me.FisheryScalerCheckBox.Text = "Fishery Scaler Report"
        Me.FisheryScalerCheckBox.UseVisualStyleBackColor = True
        Me.FisheryScalerCheckBox.Visible = False
        '
        'MSFCheckBox
        '
        Me.MSFCheckBox.AutoSize = True
        Me.MSFCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MSFCheckBox.Location = New System.Drawing.Point(35, 274)
        Me.MSFCheckBox.Name = "MSFCheckBox"
        Me.MSFCheckBox.Size = New System.Drawing.Size(612, 24)
        Me.MSFCheckBox.TabIndex = 11
        Me.MSFCheckBox.Text = "Mark-Selective Fishery Reports (Includes WDFW Chin MSF spreadsheet)"
        Me.MSFCheckBox.UseVisualStyleBackColor = True
        '
        'FishStkCompCheckBox
        '
        Me.FishStkCompCheckBox.AutoSize = True
        Me.FishStkCompCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FishStkCompCheckBox.Location = New System.Drawing.Point(33, 328)
        Me.FishStkCompCheckBox.Name = "FishStkCompCheckBox"
        Me.FishStkCompCheckBox.Size = New System.Drawing.Size(301, 24)
        Me.FishStkCompCheckBox.TabIndex = 12
        Me.FishStkCompCheckBox.Text = "Fishery Stock Composition Report"
        Me.FishStkCompCheckBox.UseVisualStyleBackColor = True
        '
        'PSCCohoERCheckBox
        '
        Me.PSCCohoERCheckBox.AutoSize = True
        Me.PSCCohoERCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PSCCohoERCheckBox.Location = New System.Drawing.Point(33, 380)
        Me.PSCCohoERCheckBox.Name = "PSCCohoERCheckBox"
        Me.PSCCohoERCheckBox.Size = New System.Drawing.Size(200, 24)
        Me.PSCCohoERCheckBox.TabIndex = 13
        Me.PSCCohoERCheckBox.Text = "PSC Coho ER Report"
        Me.PSCCohoERCheckBox.UseVisualStyleBackColor = True
        '
        'StockPer1000CheckBox
        '
        Me.StockPer1000CheckBox.AutoSize = True
        Me.StockPer1000CheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StockPer1000CheckBox.Location = New System.Drawing.Point(33, 432)
        Me.StockPer1000CheckBox.Name = "StockPer1000CheckBox"
        Me.StockPer1000CheckBox.Size = New System.Drawing.Size(220, 24)
        Me.StockPer1000CheckBox.TabIndex = 14
        Me.StockPer1000CheckBox.Text = "Stock Impacts Per 1000"
        Me.StockPer1000CheckBox.UseVisualStyleBackColor = True
        '
        'PopStatCheckBox
        '
        Me.PopStatCheckBox.AutoSize = True
        Me.PopStatCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PopStatCheckBox.Location = New System.Drawing.Point(35, 219)
        Me.PopStatCheckBox.Name = "PopStatCheckBox"
        Me.PopStatCheckBox.Size = New System.Drawing.Size(193, 24)
        Me.PopStatCheckBox.TabIndex = 15
        Me.PopStatCheckBox.Text = "Population Statistics"
        Me.PopStatCheckBox.UseVisualStyleBackColor = True
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(118, 747)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(118, 713)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 29
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(35, 747)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 28
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(35, 713)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 27
        Me.DatabaseTextLabel.Text = "Database"
        '
        'FVS_ScreenReports
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(945, 783)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.PopStatCheckBox)
        Me.Controls.Add(Me.StockPer1000CheckBox)
        Me.Controls.Add(Me.PSCCohoERCheckBox)
        Me.Controls.Add(Me.FishStkCompCheckBox)
        Me.Controls.Add(Me.MSFCheckBox)
        Me.Controls.Add(Me.FisheryScalerCheckBox)
        Me.Controls.Add(Me.RepCancelButton)
        Me.Controls.Add(Me.StockCatchCheckBox)
        Me.Controls.Add(Me.FisheryMortalityCheckBox)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FVS_ScreenReports"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Screen Reports"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents FisheryMortalityCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents StockCatchCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents RepCancelButton As System.Windows.Forms.Button
   Friend WithEvents FisheryScalerCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents MSFCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents FishStkCompCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents PSCCohoERCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents StockPer1000CheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents PopStatCheckBox As System.Windows.Forms.CheckBox
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
