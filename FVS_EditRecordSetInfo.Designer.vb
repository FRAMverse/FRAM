<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_EditRecordSetInfo
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
        Me.RSETitle = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.RunIDLabel = New System.Windows.Forms.Label
        Me.SpeciesNameLabel = New System.Windows.Forms.Label
        Me.RunNameTextBox = New System.Windows.Forms.TextBox
        Me.RunTitleTextBox = New System.Windows.Forms.TextBox
        Me.CommentsRichTextBox = New System.Windows.Forms.RichTextBox
        Me.CreationDateLabel = New System.Windows.Forms.Label
        Me.ModifyInputDateLabel = New System.Windows.Forms.Label
        Me.RunTimeDateLabel = New System.Windows.Forms.Label
        Me.REDoneButton = New System.Windows.Forms.Button
        Me.RECancelButton = New System.Windows.Forms.Button
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.BasePeriodIDLabel = New System.Windows.Forms.Label
        Me.BasePeriodNameLabel = New System.Windows.Forms.Label
        Me.lblRunYear = New System.Windows.Forms.Label
        Me.RunYearTextBox = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'RSETitle
        '
        Me.RSETitle.AutoSize = True
        Me.RSETitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RSETitle.Location = New System.Drawing.Point(304, 20)
        Me.RSETitle.Name = "RSETitle"
        Me.RSETitle.Size = New System.Drawing.Size(214, 24)
        Me.RSETitle.TabIndex = 0
        Me.RSETitle.Text = "Recordset Information"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(98, 74)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(52, 17)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "RunID"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(98, 103)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(65, 17)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Species"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(98, 199)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(78, 17)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "RunName"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(98, 235)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(69, 17)
        Me.Label4.TabIndex = 4
        Me.Label4.Text = "RunTitle"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(98, 315)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(82, 17)
        Me.Label5.TabIndex = 5
        Me.Label5.Text = "Comments"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(26, 550)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(103, 17)
        Me.Label6.TabIndex = 6
        Me.Label6.Text = "CreationDate"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(26, 579)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(125, 17)
        Me.Label7.TabIndex = 7
        Me.Label7.Text = "ModifyInputDate"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(26, 607)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(106, 17)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "RunTimeDate"
        '
        'RunIDLabel
        '
        Me.RunIDLabel.AutoSize = True
        Me.RunIDLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunIDLabel.Location = New System.Drawing.Point(242, 74)
        Me.RunIDLabel.Name = "RunIDLabel"
        Me.RunIDLabel.Size = New System.Drawing.Size(52, 17)
        Me.RunIDLabel.TabIndex = 9
        Me.RunIDLabel.Text = "RunID"
        '
        'SpeciesNameLabel
        '
        Me.SpeciesNameLabel.AutoSize = True
        Me.SpeciesNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SpeciesNameLabel.Location = New System.Drawing.Point(242, 103)
        Me.SpeciesNameLabel.Name = "SpeciesNameLabel"
        Me.SpeciesNameLabel.Size = New System.Drawing.Size(52, 17)
        Me.SpeciesNameLabel.TabIndex = 10
        Me.SpeciesNameLabel.Text = "RunID"
        '
        'RunNameTextBox
        '
        Me.RunNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunNameTextBox.Location = New System.Drawing.Point(246, 199)
        Me.RunNameTextBox.Name = "RunNameTextBox"
        Me.RunNameTextBox.Size = New System.Drawing.Size(285, 23)
        Me.RunNameTextBox.TabIndex = 11
        '
        'RunTitleTextBox
        '
        Me.RunTitleTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunTitleTextBox.Location = New System.Drawing.Point(246, 235)
        Me.RunTitleTextBox.Name = "RunTitleTextBox"
        Me.RunTitleTextBox.Size = New System.Drawing.Size(585, 23)
        Me.RunTitleTextBox.TabIndex = 12
        '
        'CommentsRichTextBox
        '
        Me.CommentsRichTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CommentsRichTextBox.Location = New System.Drawing.Point(246, 315)
        Me.CommentsRichTextBox.Name = "CommentsRichTextBox"
        Me.CommentsRichTextBox.Size = New System.Drawing.Size(585, 204)
        Me.CommentsRichTextBox.TabIndex = 13
        Me.CommentsRichTextBox.Text = ""
        '
        'CreationDateLabel
        '
        Me.CreationDateLabel.AutoSize = True
        Me.CreationDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CreationDateLabel.Location = New System.Drawing.Point(183, 550)
        Me.CreationDateLabel.Name = "CreationDateLabel"
        Me.CreationDateLabel.Size = New System.Drawing.Size(103, 17)
        Me.CreationDateLabel.TabIndex = 14
        Me.CreationDateLabel.Text = "CreationDate"
        '
        'ModifyInputDateLabel
        '
        Me.ModifyInputDateLabel.AutoSize = True
        Me.ModifyInputDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ModifyInputDateLabel.Location = New System.Drawing.Point(183, 579)
        Me.ModifyInputDateLabel.Name = "ModifyInputDateLabel"
        Me.ModifyInputDateLabel.Size = New System.Drawing.Size(103, 17)
        Me.ModifyInputDateLabel.TabIndex = 15
        Me.ModifyInputDateLabel.Text = "CreationDate"
        '
        'RunTimeDateLabel
        '
        Me.RunTimeDateLabel.AutoSize = True
        Me.RunTimeDateLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunTimeDateLabel.Location = New System.Drawing.Point(183, 607)
        Me.RunTimeDateLabel.Name = "RunTimeDateLabel"
        Me.RunTimeDateLabel.Size = New System.Drawing.Size(103, 17)
        Me.RunTimeDateLabel.TabIndex = 16
        Me.RunTimeDateLabel.Text = "CreationDate"
        '
        'REDoneButton
        '
        Me.REDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.REDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.REDoneButton.Location = New System.Drawing.Point(458, 576)
        Me.REDoneButton.Name = "REDoneButton"
        Me.REDoneButton.Size = New System.Drawing.Size(147, 51)
        Me.REDoneButton.TabIndex = 17
        Me.REDoneButton.Text = "OK - Done"
        Me.REDoneButton.UseVisualStyleBackColor = False
        '
        'RECancelButton
        '
        Me.RECancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.RECancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RECancelButton.Location = New System.Drawing.Point(663, 576)
        Me.RECancelButton.Name = "RECancelButton"
        Me.RECancelButton.Size = New System.Drawing.Size(147, 51)
        Me.RECancelButton.TabIndex = 18
        Me.RECancelButton.Text = "Cancel"
        Me.RECancelButton.UseVisualStyleBackColor = False
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(98, 134)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(106, 17)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "BasePeriodID"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(63, 170)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(137, 17)
        Me.Label10.TabIndex = 20
        Me.Label10.Text = "BasePeriod Name"
        '
        'BasePeriodIDLabel
        '
        Me.BasePeriodIDLabel.AutoSize = True
        Me.BasePeriodIDLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BasePeriodIDLabel.Location = New System.Drawing.Point(242, 134)
        Me.BasePeriodIDLabel.Name = "BasePeriodIDLabel"
        Me.BasePeriodIDLabel.Size = New System.Drawing.Size(65, 17)
        Me.BasePeriodIDLabel.TabIndex = 21
        Me.BasePeriodIDLabel.Text = "Species"
        '
        'BasePeriodNameLabel
        '
        Me.BasePeriodNameLabel.AutoSize = True
        Me.BasePeriodNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BasePeriodNameLabel.Location = New System.Drawing.Point(242, 170)
        Me.BasePeriodNameLabel.Name = "BasePeriodNameLabel"
        Me.BasePeriodNameLabel.Size = New System.Drawing.Size(65, 17)
        Me.BasePeriodNameLabel.TabIndex = 22
        Me.BasePeriodNameLabel.Text = "Species"
        '
        'lblRunYear
        '
        Me.lblRunYear.AutoSize = True
        Me.lblRunYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRunYear.Location = New System.Drawing.Point(98, 270)
        Me.lblRunYear.Name = "lblRunYear"
        Me.lblRunYear.Size = New System.Drawing.Size(71, 17)
        Me.lblRunYear.TabIndex = 23
        Me.lblRunYear.Text = "RunYear"
        '
        'RunYearTextBox
        '
        Me.RunYearTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RunYearTextBox.Location = New System.Drawing.Point(246, 270)
        Me.RunYearTextBox.Name = "RunYearTextBox"
        Me.RunYearTextBox.Size = New System.Drawing.Size(188, 23)
        Me.RunYearTextBox.TabIndex = 24
        '
        'FVS_EditRecordSetInfo
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(874, 684)
        Me.Controls.Add(Me.RunYearTextBox)
        Me.Controls.Add(Me.lblRunYear)
        Me.Controls.Add(Me.BasePeriodNameLabel)
        Me.Controls.Add(Me.BasePeriodIDLabel)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.RECancelButton)
        Me.Controls.Add(Me.REDoneButton)
        Me.Controls.Add(Me.RunTimeDateLabel)
        Me.Controls.Add(Me.ModifyInputDateLabel)
        Me.Controls.Add(Me.CreationDateLabel)
        Me.Controls.Add(Me.CommentsRichTextBox)
        Me.Controls.Add(Me.RunTitleTextBox)
        Me.Controls.Add(Me.RunNameTextBox)
        Me.Controls.Add(Me.SpeciesNameLabel)
        Me.Controls.Add(Me.RunIDLabel)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.RSETitle)
        Me.Name = "FVS_EditRecordSetInfo"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Edit Recordset Information"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
   Friend WithEvents RSETitle As System.Windows.Forms.Label
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents Label3 As System.Windows.Forms.Label
   Friend WithEvents Label4 As System.Windows.Forms.Label
   Friend WithEvents Label5 As System.Windows.Forms.Label
   Friend WithEvents Label6 As System.Windows.Forms.Label
   Friend WithEvents Label7 As System.Windows.Forms.Label
   Friend WithEvents Label8 As System.Windows.Forms.Label
   Friend WithEvents RunIDLabel As System.Windows.Forms.Label
   Friend WithEvents SpeciesNameLabel As System.Windows.Forms.Label
   Friend WithEvents RunNameTextBox As System.Windows.Forms.TextBox
   Friend WithEvents RunTitleTextBox As System.Windows.Forms.TextBox
   Friend WithEvents CommentsRichTextBox As System.Windows.Forms.RichTextBox
   Friend WithEvents CreationDateLabel As System.Windows.Forms.Label
   Friend WithEvents ModifyInputDateLabel As System.Windows.Forms.Label
   Friend WithEvents RunTimeDateLabel As System.Windows.Forms.Label
   Friend WithEvents REDoneButton As System.Windows.Forms.Button
   Friend WithEvents RECancelButton As System.Windows.Forms.Button
   Friend WithEvents Label9 As System.Windows.Forms.Label
   Friend WithEvents Label10 As System.Windows.Forms.Label
   Friend WithEvents BasePeriodIDLabel As System.Windows.Forms.Label
    Friend WithEvents BasePeriodNameLabel As System.Windows.Forms.Label
    Friend WithEvents lblRunYear As System.Windows.Forms.Label
    Friend WithEvents RunYearTextBox As System.Windows.Forms.TextBox
End Class
