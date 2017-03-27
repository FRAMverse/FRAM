<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_VersionChanges
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
      Me.VersionChangeListBox = New System.Windows.Forms.ListBox()
      Me.VersionTitleLabel = New System.Windows.Forms.Label()
      Me.FVS_Done = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'VersionChangeListBox
      '
      Me.VersionChangeListBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.VersionChangeListBox.FormattingEnabled = True
      Me.VersionChangeListBox.ItemHeight = 16
      Me.VersionChangeListBox.Location = New System.Drawing.Point(12, 64)
      Me.VersionChangeListBox.Name = "VersionChangeListBox"
      Me.VersionChangeListBox.Size = New System.Drawing.Size(1004, 580)
      Me.VersionChangeListBox.TabIndex = 0
      '
      'VersionTitleLabel
      '
      Me.VersionTitleLabel.AutoSize = True
      Me.VersionTitleLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.VersionTitleLabel.Location = New System.Drawing.Point(345, 19)
      Me.VersionTitleLabel.Name = "VersionTitleLabel"
      Me.VersionTitleLabel.Size = New System.Drawing.Size(339, 29)
      Me.VersionTitleLabel.TabIndex = 1
      Me.VersionTitleLabel.Text = "Version Change Description"
      '
      'FVS_Done
      '
      Me.FVS_Done.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.FVS_Done.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.FVS_Done.Location = New System.Drawing.Point(433, 676)
      Me.FVS_Done.Name = "FVS_Done"
      Me.FVS_Done.Size = New System.Drawing.Size(162, 54)
      Me.FVS_Done.TabIndex = 3
      Me.FVS_Done.Text = "OK - Done"
      Me.FVS_Done.UseVisualStyleBackColor = False
      '
      'FVS_VersionChanges
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(1028, 761)
      Me.Controls.Add(Me.FVS_Done)
      Me.Controls.Add(Me.VersionTitleLabel)
      Me.Controls.Add(Me.VersionChangeListBox)
      Me.Name = "FVS_VersionChanges"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "FramVS Version Changes"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents VersionChangeListBox As System.Windows.Forms.ListBox
   Friend WithEvents VersionTitleLabel As System.Windows.Forms.Label
   Friend WithEvents FVS_Done As System.Windows.Forms.Button
End Class
