<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_BackwardsFram
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
        Me.Label2 = New System.Windows.Forms.Label
        Me.TargetEscButton = New System.Windows.Forms.Button
        Me.NumBackFRAMIterationsTextBox = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.StartIterationsButton = New System.Windows.Forms.Button
        Me.ExitButton = New System.Windows.Forms.Button
        Me.MSMRecsButton = New System.Windows.Forms.Button
        Me.SaveScalersButton = New System.Windows.Forms.Button
        Me.IterProgressLabel = New System.Windows.Forms.Label
        Me.IterProgressTextBox = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.RecordSetNameLabel = New System.Windows.Forms.Label
        Me.DatabaseNameLabel = New System.Windows.Forms.Label
        Me.RecordSetTextLabel = New System.Windows.Forms.Label
        Me.DatabaseTextLabel = New System.Windows.Forms.Label
        Me.NoMSFBiasCorrection = New System.Windows.Forms.CheckBox
        Me.chk2from3 = New System.Windows.Forms.CheckBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(305, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(277, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Backwards FRAM Run Menu"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(106, 67)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(674, 24)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Post-Season Stock Abundance using Observed Catch and Escapement"
        '
        'TargetEscButton
        '
        Me.TargetEscButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TargetEscButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TargetEscButton.Location = New System.Drawing.Point(319, 140)
        Me.TargetEscButton.Name = "TargetEscButton"
        Me.TargetEscButton.Size = New System.Drawing.Size(248, 55)
        Me.TargetEscButton.TabIndex = 2
        Me.TargetEscButton.Text = "Target Escapements"
        Me.TargetEscButton.UseVisualStyleBackColor = False
        '
        'NumBackFRAMIterationsTextBox
        '
        Me.NumBackFRAMIterationsTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NumBackFRAMIterationsTextBox.Location = New System.Drawing.Point(319, 213)
        Me.NumBackFRAMIterationsTextBox.Name = "NumBackFRAMIterationsTextBox"
        Me.NumBackFRAMIterationsTextBox.Size = New System.Drawing.Size(46, 26)
        Me.NumBackFRAMIterationsTextBox.TabIndex = 3
        Me.NumBackFRAMIterationsTextBox.Text = "99"
        Me.NumBackFRAMIterationsTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(371, 216)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(174, 20)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "Number of Iterations"
        '
        'StartIterationsButton
        '
        Me.StartIterationsButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StartIterationsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StartIterationsButton.Location = New System.Drawing.Point(319, 266)
        Me.StartIterationsButton.Name = "StartIterationsButton"
        Me.StartIterationsButton.Size = New System.Drawing.Size(248, 55)
        Me.StartIterationsButton.TabIndex = 5
        Me.StartIterationsButton.Text = "Start Iterations"
        Me.StartIterationsButton.UseVisualStyleBackColor = False
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ExitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExitButton.Location = New System.Drawing.Point(319, 368)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.Size = New System.Drawing.Size(248, 55)
        Me.ExitButton.TabIndex = 6
        Me.ExitButton.Text = "EXIT"
        Me.ExitButton.UseVisualStyleBackColor = False
        '
        'MSMRecsButton
        '
        Me.MSMRecsButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.MSMRecsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MSMRecsButton.Location = New System.Drawing.Point(202, 429)
        Me.MSMRecsButton.Name = "MSMRecsButton"
        Me.MSMRecsButton.Size = New System.Drawing.Size(483, 55)
        Me.MSMRecsButton.TabIndex = 7
        Me.MSMRecsButton.Text = "Create MSM Imputed CWT Recoveries"
        Me.MSMRecsButton.UseVisualStyleBackColor = False
        '
        'SaveScalersButton
        '
        Me.SaveScalersButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.SaveScalersButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SaveScalersButton.Location = New System.Drawing.Point(319, 533)
        Me.SaveScalersButton.Name = "SaveScalersButton"
        Me.SaveScalersButton.Size = New System.Drawing.Size(248, 55)
        Me.SaveScalersButton.TabIndex = 8
        Me.SaveScalersButton.Text = " Save BkFRAM targets and new Recruit Scalars"
        Me.SaveScalersButton.UseVisualStyleBackColor = False
        '
        'IterProgressLabel
        '
        Me.IterProgressLabel.AutoSize = True
        Me.IterProgressLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IterProgressLabel.Location = New System.Drawing.Point(314, 495)
        Me.IterProgressLabel.Name = "IterProgressLabel"
        Me.IterProgressLabel.Size = New System.Drawing.Size(177, 20)
        Me.IterProgressLabel.TabIndex = 9
        Me.IterProgressLabel.Text = "Working on Iteration "
        '
        'IterProgressTextBox
        '
        Me.IterProgressTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.IterProgressTextBox.Location = New System.Drawing.Point(531, 490)
        Me.IterProgressTextBox.Name = "IterProgressTextBox"
        Me.IterProgressTextBox.Size = New System.Drawing.Size(46, 26)
        Me.IterProgressTextBox.TabIndex = 10
        Me.IterProgressTextBox.Text = "99"
        Me.IterProgressTextBox.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
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
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(128, 649)
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
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(128, 615)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 18
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(45, 649)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 17
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(45, 615)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 16
        Me.DatabaseTextLabel.Text = "Database"
        '
        'NoMSFBiasCorrection
        '
        Me.NoMSFBiasCorrection.AutoSize = True
        Me.NoMSFBiasCorrection.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.NoMSFBiasCorrection.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NoMSFBiasCorrection.Location = New System.Drawing.Point(319, 327)
        Me.NoMSFBiasCorrection.Name = "NoMSFBiasCorrection"
        Me.NoMSFBiasCorrection.Size = New System.Drawing.Size(285, 17)
        Me.NoMSFBiasCorrection.TabIndex = 20
        Me.NoMSFBiasCorrection.Text = "Run without MSF Bias Correction (if checked)"
        Me.NoMSFBiasCorrection.UseVisualStyleBackColor = False
        '
        'chk2from3
        '
        Me.chk2from3.AutoSize = True
        Me.chk2from3.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.chk2from3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chk2from3.Location = New System.Drawing.Point(319, 338)
        Me.chk2from3.Name = "chk2from3"
        Me.chk2from3.Size = New System.Drawing.Size(131, 24)
        Me.chk2from3.TabIndex = 21
        Me.chk2from3.Text = "Age 2 from 3"
        Me.chk2from3.UseVisualStyleBackColor = False
        '
        'FVS_BackwardsFram
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(887, 795)
        Me.Controls.Add(Me.chk2from3)
        Me.Controls.Add(Me.NoMSFBiasCorrection)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.IterProgressTextBox)
        Me.Controls.Add(Me.IterProgressLabel)
        Me.Controls.Add(Me.SaveScalersButton)
        Me.Controls.Add(Me.MSMRecsButton)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.StartIterationsButton)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.NumBackFRAMIterationsTextBox)
        Me.Controls.Add(Me.TargetEscButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "FVS_BackwardsFram"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FVS_BackwardsFram"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents TargetEscButton As System.Windows.Forms.Button
   Friend WithEvents NumBackFRAMIterationsTextBox As System.Windows.Forms.TextBox
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents StartIterationsButton As System.Windows.Forms.Button
   Friend WithEvents ExitButton As System.Windows.Forms.Button
   Friend WithEvents MSMRecsButton As System.Windows.Forms.Button
   Friend WithEvents SaveScalersButton As System.Windows.Forms.Button
   Friend WithEvents IterProgressLabel As System.Windows.Forms.Label
   Friend WithEvents IterProgressTextBox As System.Windows.Forms.TextBox
   Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
    Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
    Friend WithEvents NoMSFBiasCorrection As System.Windows.Forms.CheckBox
    Friend WithEvents chk2from3 As System.Windows.Forms.CheckBox
End Class
