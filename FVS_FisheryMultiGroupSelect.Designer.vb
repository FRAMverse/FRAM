<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_FisheryMultiGroupSelect
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
      Me.SelectTypeLabel = New System.Windows.Forms.Label()
      Me.FisheryGroupNameTextBox = New System.Windows.Forms.TextBox()
      Me.FisheryGroupLabel = New System.Windows.Forms.Label()
      Me.FGCancelButton = New System.Windows.Forms.Button()
      Me.FGDoneButton = New System.Windows.Forms.Button()
      Me.FisheryListBox = New System.Windows.Forms.CheckedListBox()
      Me.FisherySelectionTitle = New System.Windows.Forms.Label()
      Me.FisherySelectedListBox = New System.Windows.Forms.CheckedListBox()
      Me.Label1 = New System.Windows.Forms.Label()
      Me.FGNextGrpButton = New System.Windows.Forms.Button()
      Me.FGReviewButton = New System.Windows.Forms.Button()
      Me.FGNextReviewButton = New System.Windows.Forms.Button()
      Me.FGExitReviewButton = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'SelectTypeLabel
      '
      Me.SelectTypeLabel.AutoSize = True
      Me.SelectTypeLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SelectTypeLabel.Location = New System.Drawing.Point(60, 67)
      Me.SelectTypeLabel.Name = "SelectTypeLabel"
      Me.SelectTypeLabel.Size = New System.Drawing.Size(292, 17)
      Me.SelectTypeLabel.TabIndex = 20
      Me.SelectTypeLabel.Text = "Fisheries Available for Group Selection"
      '
      'FisheryGroupNameTextBox
      '
      Me.FisheryGroupNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisheryGroupNameTextBox.Location = New System.Drawing.Point(289, 709)
      Me.FisheryGroupNameTextBox.Name = "FisheryGroupNameTextBox"
      Me.FisheryGroupNameTextBox.Size = New System.Drawing.Size(330, 23)
      Me.FisheryGroupNameTextBox.TabIndex = 19
      '
      'FisheryGroupLabel
      '
      Me.FisheryGroupLabel.AutoSize = True
      Me.FisheryGroupLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisheryGroupLabel.Location = New System.Drawing.Point(126, 712)
      Me.FisheryGroupLabel.Name = "FisheryGroupLabel"
      Me.FisheryGroupLabel.Size = New System.Drawing.Size(121, 13)
      Me.FisheryGroupLabel.TabIndex = 18
      Me.FisheryGroupLabel.Text = "Fishery Group Name"
      '
      'FGCancelButton
      '
      Me.FGCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGCancelButton.Location = New System.Drawing.Point(696, 755)
      Me.FGCancelButton.Name = "FGCancelButton"
      Me.FGCancelButton.Size = New System.Drawing.Size(136, 39)
      Me.FGCancelButton.TabIndex = 17
      Me.FGCancelButton.Text = "Cancel"
      Me.FGCancelButton.UseVisualStyleBackColor = False
      '
      'FGDoneButton
      '
      Me.FGDoneButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGDoneButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGDoneButton.Location = New System.Drawing.Point(76, 755)
      Me.FGDoneButton.Name = "FGDoneButton"
      Me.FGDoneButton.Size = New System.Drawing.Size(139, 39)
      Me.FGDoneButton.TabIndex = 16
      Me.FGDoneButton.Text = "OK - Done"
      Me.FGDoneButton.UseVisualStyleBackColor = False
      '
      'FisheryListBox
      '
      Me.FisheryListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FisheryListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisheryListBox.FormattingEnabled = True
      Me.FisheryListBox.Location = New System.Drawing.Point(23, 90)
      Me.FisheryListBox.Name = "FisheryListBox"
      Me.FisheryListBox.Size = New System.Drawing.Size(440, 598)
      Me.FisheryListBox.TabIndex = 15
      '
      'FisherySelectionTitle
      '
      Me.FisherySelectionTitle.AutoSize = True
      Me.FisherySelectionTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisherySelectionTitle.Location = New System.Drawing.Point(339, 20)
      Me.FisherySelectionTitle.Name = "FisherySelectionTitle"
      Me.FisherySelectionTitle.Size = New System.Drawing.Size(236, 24)
      Me.FisherySelectionTitle.TabIndex = 14
      Me.FisherySelectionTitle.Text = "Fishery Group Selection"
      '
      'FisherySelectedListBox
      '
      Me.FisherySelectedListBox.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FisherySelectedListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FisherySelectedListBox.FormattingEnabled = True
      Me.FisherySelectedListBox.Location = New System.Drawing.Point(520, 90)
      Me.FisherySelectedListBox.Name = "FisherySelectedListBox"
      Me.FisherySelectedListBox.Size = New System.Drawing.Size(440, 598)
      Me.FisherySelectedListBox.TabIndex = 21
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(553, 67)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(248, 17)
      Me.Label1.TabIndex = 22
      Me.Label1.Text = "Fisheries Selected for this Group"
      '
      'FGNextGrpButton
      '
      Me.FGNextGrpButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGNextGrpButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGNextGrpButton.Location = New System.Drawing.Point(276, 755)
      Me.FGNextGrpButton.Name = "FGNextGrpButton"
      Me.FGNextGrpButton.Size = New System.Drawing.Size(152, 39)
      Me.FGNextGrpButton.TabIndex = 23
      Me.FGNextGrpButton.Text = "Next Group"
      Me.FGNextGrpButton.UseVisualStyleBackColor = False
      '
      'FGReviewButton
      '
      Me.FGReviewButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGReviewButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGReviewButton.Location = New System.Drawing.Point(480, 755)
      Me.FGReviewButton.Name = "FGReviewButton"
      Me.FGReviewButton.Size = New System.Drawing.Size(154, 39)
      Me.FGReviewButton.TabIndex = 24
      Me.FGReviewButton.Text = "Review Groups"
      Me.FGReviewButton.UseVisualStyleBackColor = False
      '
      'FGNextReviewButton
      '
      Me.FGNextReviewButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGNextReviewButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGNextReviewButton.Location = New System.Drawing.Point(647, 700)
      Me.FGNextReviewButton.Name = "FGNextReviewButton"
      Me.FGNextReviewButton.Size = New System.Drawing.Size(139, 39)
      Me.FGNextReviewButton.TabIndex = 25
      Me.FGNextReviewButton.Text = "Next Review"
      Me.FGNextReviewButton.UseVisualStyleBackColor = False
      '
      'FGExitReviewButton
      '
      Me.FGExitReviewButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FGExitReviewButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FGExitReviewButton.Location = New System.Drawing.Point(808, 700)
      Me.FGExitReviewButton.Name = "FGExitReviewButton"
      Me.FGExitReviewButton.Size = New System.Drawing.Size(139, 39)
      Me.FGExitReviewButton.TabIndex = 26
      Me.FGExitReviewButton.Text = "Exit Review"
      Me.FGExitReviewButton.UseVisualStyleBackColor = False
      '
      'FVS_FisheryMultiGroupSelect
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(1003, 806)
      Me.Controls.Add(Me.FGExitReviewButton)
      Me.Controls.Add(Me.FGNextReviewButton)
      Me.Controls.Add(Me.FGReviewButton)
      Me.Controls.Add(Me.FGNextGrpButton)
      Me.Controls.Add(Me.Label1)
      Me.Controls.Add(Me.FisherySelectedListBox)
      Me.Controls.Add(Me.SelectTypeLabel)
      Me.Controls.Add(Me.FisheryGroupNameTextBox)
      Me.Controls.Add(Me.FisheryGroupLabel)
      Me.Controls.Add(Me.FGCancelButton)
      Me.Controls.Add(Me.FGDoneButton)
      Me.Controls.Add(Me.FisheryListBox)
      Me.Controls.Add(Me.FisherySelectionTitle)
      Me.Name = "FVS_FisheryMultiGroupSelect"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FVS_FisheryMultiGroupSelect"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents SelectTypeLabel As System.Windows.Forms.Label
   Friend WithEvents FisheryGroupNameTextBox As System.Windows.Forms.TextBox
   Friend WithEvents FisheryGroupLabel As System.Windows.Forms.Label
   Friend WithEvents FGCancelButton As System.Windows.Forms.Button
   Friend WithEvents FGDoneButton As System.Windows.Forms.Button
   Friend WithEvents FisheryListBox As System.Windows.Forms.CheckedListBox
   Friend WithEvents FisherySelectionTitle As System.Windows.Forms.Label
   Friend WithEvents FisherySelectedListBox As System.Windows.Forms.CheckedListBox
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents FGNextGrpButton As System.Windows.Forms.Button
   Friend WithEvents FGReviewButton As System.Windows.Forms.Button
   Friend WithEvents FGNextReviewButton As System.Windows.Forms.Button
   Friend WithEvents FGExitReviewButton As System.Windows.Forms.Button
End Class
