<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_ActiveRateScaler
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
        Me.StkFishRateScalerGrid = New System.Windows.Forms.DataGridView
        Me.lblSFRSTitle = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox1 = New System.Windows.Forms.TextBox
        CType(Me.StkFishRateScalerGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StkFishRateScalerGrid
        '
        Me.StkFishRateScalerGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.StkFishRateScalerGrid.Location = New System.Drawing.Point(50, 89)
        Me.StkFishRateScalerGrid.Name = "StkFishRateScalerGrid"
        Me.StkFishRateScalerGrid.Size = New System.Drawing.Size(980, 507)
        Me.StkFishRateScalerGrid.TabIndex = 0
        '
        'lblSFRSTitle
        '
        Me.lblSFRSTitle.AutoSize = True
        Me.lblSFRSTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSFRSTitle.Location = New System.Drawing.Point(46, 31)
        Me.lblSFRSTitle.Name = "lblSFRSTitle"
        Me.lblSFRSTitle.Size = New System.Drawing.Size(322, 24)
        Me.lblSFRSTitle.TabIndex = 1
        Me.lblSFRSTitle.Text = "Active Stock Fishery Rate Scalers"
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(466, 621)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(151, 35)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "OK - Done"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.SystemColors.Control
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.Red
        Me.TextBox1.Location = New System.Drawing.Point(50, 63)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(481, 15)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = "Use previous screen for updating scalers, if desired!!!!"
        '
        'FVS_ActiveRateScaler
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1122, 686)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.lblSFRSTitle)
        Me.Controls.Add(Me.StkFishRateScalerGrid)
        Me.Name = "FVS_ActiveRateScaler"
        Me.Text = "FVS_ActiveRateScaler"
        CType(Me.StkFishRateScalerGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StkFishRateScalerGrid As System.Windows.Forms.DataGridView
    Friend WithEvents lblSFRSTitle As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
End Class
