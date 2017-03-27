<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_MultipleRunDeletion
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
      Me.list_MultiDelete = New System.Windows.Forms.CheckedListBox()
      Me.CmdCancel = New System.Windows.Forms.Button()
      Me.DeleteSelection = New System.Windows.Forms.Button()
      Me.SuspendLayout()
      '
      'list_MultiDelete
      '
      Me.list_MultiDelete.CheckOnClick = True
      Me.list_MultiDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.list_MultiDelete.FormattingEnabled = True
      Me.list_MultiDelete.Location = New System.Drawing.Point(25, 21)
      Me.list_MultiDelete.Name = "list_MultiDelete"
      Me.list_MultiDelete.ScrollAlwaysVisible = True
      Me.list_MultiDelete.Size = New System.Drawing.Size(693, 439)
      Me.list_MultiDelete.TabIndex = 0
      '
      'CmdCancel
      '
      Me.CmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.CmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.CmdCancel.Location = New System.Drawing.Point(384, 485)
      Me.CmdCancel.Name = "CmdCancel"
      Me.CmdCancel.Size = New System.Drawing.Size(171, 49)
      Me.CmdCancel.TabIndex = 3
      Me.CmdCancel.Text = "CANCEL"
      Me.CmdCancel.UseVisualStyleBackColor = False
      '
      'DeleteSelection
      '
      Me.DeleteSelection.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.DeleteSelection.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DeleteSelection.Location = New System.Drawing.Point(133, 485)
      Me.DeleteSelection.Name = "DeleteSelection"
      Me.DeleteSelection.Size = New System.Drawing.Size(219, 49)
      Me.DeleteSelection.TabIndex = 6
      Me.DeleteSelection.Text = "Delete Selection Done"
      Me.DeleteSelection.UseVisualStyleBackColor = False
      '
      'FVS_MultipleRunDeletion
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.AutoScroll = True
      Me.ClientSize = New System.Drawing.Size(750, 546)
      Me.Controls.Add(Me.DeleteSelection)
      Me.Controls.Add(Me.CmdCancel)
      Me.Controls.Add(Me.list_MultiDelete)
      Me.Name = "FVS_MultipleRunDeletion"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Delete Multiple Model Runs"
      Me.ResumeLayout(False)

   End Sub
   Friend WithEvents list_MultiDelete As System.Windows.Forms.CheckedListBox
   Friend WithEvents CmdCancel As System.Windows.Forms.Button
   Friend WithEvents DeleteSelection As System.Windows.Forms.Button
End Class
