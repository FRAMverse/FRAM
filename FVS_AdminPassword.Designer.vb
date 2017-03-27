<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_AdminPassword
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
      Me.Button1 = New System.Windows.Forms.Button()
      Me.Button2 = New System.Windows.Forms.Button()
      Me.txt_pwentry = New System.Windows.Forms.TextBox()
      Me.lbl_PW = New System.Windows.Forms.Label()
      Me.chk_LoadSLRatio = New System.Windows.Forms.CheckBox()
      Me.SuspendLayout()
      '
      'Button1
      '
      Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Button1.Location = New System.Drawing.Point(33, 79)
      Me.Button1.Name = "Button1"
      Me.Button1.Size = New System.Drawing.Size(110, 34)
      Me.Button1.TabIndex = 0
      Me.Button1.Text = "Initialize"
      Me.Button1.UseVisualStyleBackColor = True
      '
      'Button2
      '
      Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Button2.Location = New System.Drawing.Point(158, 80)
      Me.Button2.Name = "Button2"
      Me.Button2.Size = New System.Drawing.Size(110, 34)
      Me.Button2.TabIndex = 1
      Me.Button2.Text = "Cancel"
      Me.Button2.UseVisualStyleBackColor = True
      '
      'txt_pwentry
      '
      Me.txt_pwentry.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.txt_pwentry.Location = New System.Drawing.Point(31, 42)
      Me.txt_pwentry.Name = "txt_pwentry"
      Me.txt_pwentry.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
      Me.txt_pwentry.Size = New System.Drawing.Size(244, 26)
      Me.txt_pwentry.TabIndex = 0
      '
      'lbl_PW
      '
      Me.lbl_PW.AutoSize = True
      Me.lbl_PW.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.lbl_PW.Location = New System.Drawing.Point(27, 11)
      Me.lbl_PW.Name = "lbl_PW"
      Me.lbl_PW.Size = New System.Drawing.Size(263, 20)
      Me.lbl_PW.TabIndex = 3
      Me.lbl_PW.Text = "Enter secret nuclear code to update"
      '
      'chk_LoadSLRatio
      '
      Me.chk_LoadSLRatio.AutoSize = True
      Me.chk_LoadSLRatio.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.chk_LoadSLRatio.Location = New System.Drawing.Point(33, 126)
      Me.chk_LoadSLRatio.Name = "chk_LoadSLRatio"
      Me.chk_LoadSLRatio.Size = New System.Drawing.Size(282, 20)
      Me.chk_LoadSLRatio.TabIndex = 4
      Me.chk_LoadSLRatio.Text = "Load TargetRatio from spreadsheet?"
      Me.chk_LoadSLRatio.UseVisualStyleBackColor = True
      '
      'FVS_AdminPassword
      '
      Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
      Me.ClientSize = New System.Drawing.Size(348, 160)
      Me.Controls.Add(Me.chk_LoadSLRatio)
      Me.Controls.Add(Me.lbl_PW)
      Me.Controls.Add(Me.txt_pwentry)
      Me.Controls.Add(Me.Button2)
      Me.Controls.Add(Me.Button1)
      Me.Name = "FVS_AdminPassword"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "S:L Ratio Updater"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Button1 As System.Windows.Forms.Button
   Friend WithEvents Button2 As System.Windows.Forms.Button
   Friend WithEvents txt_pwentry As System.Windows.Forms.TextBox
   Friend WithEvents lbl_PW As System.Windows.Forms.Label
   Friend WithEvents chk_LoadSLRatio As System.Windows.Forms.CheckBox
End Class
