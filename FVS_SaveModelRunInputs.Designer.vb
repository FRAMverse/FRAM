<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_SaveModelRunInputs
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
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.ModelRunNameLabel = New System.Windows.Forms.Label()
      Me.Label2 = New System.Windows.Forms.Label()
      Me.SMRReplaceButton = New System.Windows.Forms.Button()
      Me.SMRSaveButton = New System.Windows.Forms.Button()
      Me.Label3 = New System.Windows.Forms.Label()
      Me.Label4 = New System.Windows.Forms.Label()
      Me.SMRCancelButton = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(295, 63)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(226, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Save Model Run Inputs"
      '
      'DatabaseNameLabel
      '
      Me.DatabaseNameLabel.AutoSize = True
      Me.DatabaseNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.DatabaseNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(300, 168)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 17
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(217, 202)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(68, 13)
      Me.RecordSetTextLabel.TabIndex = 16
      Me.RecordSetTextLabel.Text = "Model Run"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(217, 168)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 15
      Me.DatabaseTextLabel.Text = "Database"
      '
      'ModelRunNameLabel
      '
      Me.ModelRunNameLabel.AutoSize = True
      Me.ModelRunNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.ModelRunNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ModelRunNameLabel.Location = New System.Drawing.Point(300, 202)
      Me.ModelRunNameLabel.Name = "ModelRunNameLabel"
      Me.ModelRunNameLabel.Size = New System.Drawing.Size(129, 17)
      Me.ModelRunNameLabel.TabIndex = 18
      Me.ModelRunNameLabel.Text = "Model Run name"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label2.Location = New System.Drawing.Point(100, 301)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(542, 24)
      Me.Label2.TabIndex = 19
      Me.Label2.Text = "Replace Current Model Run or Save New Model Run ???"
      '
      'SMRReplaceButton
      '
      Me.SMRReplaceButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SMRReplaceButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SMRReplaceButton.Location = New System.Drawing.Point(94, 444)
      Me.SMRReplaceButton.Name = "SMRReplaceButton"
      Me.SMRReplaceButton.Size = New System.Drawing.Size(297, 53)
      Me.SMRReplaceButton.TabIndex = 20
      Me.SMRReplaceButton.Text = "Replace Current Model Run "
      Me.SMRReplaceButton.UseVisualStyleBackColor = False
      '
      'SMRSaveButton
      '
      Me.SMRSaveButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SMRSaveButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SMRSaveButton.Location = New System.Drawing.Point(426, 444)
      Me.SMRSaveButton.Name = "SMRSaveButton"
      Me.SMRSaveButton.Size = New System.Drawing.Size(297, 53)
      Me.SMRSaveButton.TabIndex = 21
      Me.SMRSaveButton.Text = "Save NEW Model Run "
      Me.SMRSaveButton.UseVisualStyleBackColor = False
      '
      'Label3
      '
      Me.Label3.AutoSize = True
      Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label3.Location = New System.Drawing.Point(101, 421)
      Me.Label3.Name = "Label3"
      Me.Label3.Size = New System.Drawing.Size(232, 17)
      Me.Label3.TabIndex = 22
      Me.Label3.Text = "Current Model Run Overwritten"
      '
      'Label4
      '
      Me.Label4.AutoSize = True
      Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label4.Location = New System.Drawing.Point(432, 421)
      Me.Label4.Name = "Label4"
      Me.Label4.Size = New System.Drawing.Size(182, 17)
      Me.Label4.TabIndex = 23
      Me.Label4.Text = "New Model Run Created"
      '
      'SMRCancelButton
      '
      Me.SMRCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SMRCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SMRCancelButton.Location = New System.Drawing.Point(279, 560)
      Me.SMRCancelButton.Name = "SMRCancelButton"
      Me.SMRCancelButton.Size = New System.Drawing.Size(297, 53)
      Me.SMRCancelButton.TabIndex = 24
      Me.SMRCancelButton.Text = "Cancel Save"
      Me.SMRCancelButton.UseVisualStyleBackColor = False
      '
      'FVS_SaveModelRunInputs
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(871, 747)
      Me.Controls.Add(Me.SMRCancelButton)
      Me.Controls.Add(Me.Label4)
      Me.Controls.Add(Me.Label3)
      Me.Controls.Add(Me.SMRSaveButton)
      Me.Controls.Add(Me.SMRReplaceButton)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.ModelRunNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.Label1)
      Me.Name = "FVS_SaveModelRunInputs"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_SaveModelRunInputs"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
   Friend WithEvents ModelRunNameLabel As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents SMRReplaceButton As System.Windows.Forms.Button
   Friend WithEvents SMRSaveButton As System.Windows.Forms.Button
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents SMRCancelButton As System.Windows.Forms.Button
End Class
