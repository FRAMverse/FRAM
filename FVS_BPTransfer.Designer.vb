<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_BPTransfer
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
        Me.BPSizeLimitchk = New System.Windows.Forms.CheckBox
        Me.Stockschk = New System.Windows.Forms.CheckBox
        Me.Fisherieschk = New System.Windows.Forms.CheckBox
        Me.TimeStepschk = New System.Windows.Forms.CheckBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BPSizeLimitchk
        '
        Me.BPSizeLimitchk.AutoSize = True
        Me.BPSizeLimitchk.Location = New System.Drawing.Point(24, 32)
        Me.BPSizeLimitchk.Name = "BPSizeLimitchk"
        Me.BPSizeLimitchk.Size = New System.Drawing.Size(207, 22)
        Me.BPSizeLimitchk.TabIndex = 0
        Me.BPSizeLimitchk.Text = "Base Period Size Limits"
        Me.BPSizeLimitchk.UseVisualStyleBackColor = True
        '
        'Stockschk
        '
        Me.Stockschk.AutoSize = True
        Me.Stockschk.Location = New System.Drawing.Point(24, 75)
        Me.Stockschk.Name = "Stockschk"
        Me.Stockschk.Size = New System.Drawing.Size(80, 22)
        Me.Stockschk.TabIndex = 1
        Me.Stockschk.Text = "Stocks"
        Me.Stockschk.UseVisualStyleBackColor = True
        '
        'Fisherieschk
        '
        Me.Fisherieschk.AutoSize = True
        Me.Fisherieschk.Location = New System.Drawing.Point(24, 119)
        Me.Fisherieschk.Name = "Fisherieschk"
        Me.Fisherieschk.Size = New System.Drawing.Size(96, 22)
        Me.Fisherieschk.TabIndex = 2
        Me.Fisherieschk.Text = "Fisheries"
        Me.Fisherieschk.UseVisualStyleBackColor = True
        '
        'TimeStepschk
        '
        Me.TimeStepschk.AutoSize = True
        Me.TimeStepschk.Location = New System.Drawing.Point(24, 164)
        Me.TimeStepschk.Name = "TimeStepschk"
        Me.TimeStepschk.Size = New System.Drawing.Size(112, 22)
        Me.TimeStepschk.TabIndex = 3
        Me.TimeStepschk.Text = "Time Steps"
        Me.TimeStepschk.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.White
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.ForeColor = System.Drawing.Color.SkyBlue
        Me.TextBox1.Location = New System.Drawing.Point(32, 105)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(488, 65)
        Me.TextBox1.TabIndex = 4
        Me.TextBox1.Text = "Tables should only be selected when importing a base period resulting from a new " & _
            "calibration."
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.PapayaWhip
        Me.GroupBox1.Controls.Add(Me.TimeStepschk)
        Me.GroupBox1.Controls.Add(Me.Fisherieschk)
        Me.GroupBox1.Controls.Add(Me.Stockschk)
        Me.GroupBox1.Controls.Add(Me.BPSizeLimitchk)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.GroupBox1.Location = New System.Drawing.Point(32, 176)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(338, 201)
        Me.GroupBox1.TabIndex = 5
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Please Make Your Selection"
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.White
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Font = New System.Drawing.Font("Verdana", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.ForeColor = System.Drawing.Color.Chocolate
        Me.TextBox2.Location = New System.Drawing.Point(32, 69)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(488, 20)
        Me.TextBox2.TabIndex = 6
        Me.TextBox2.Text = "Warning: This will replace the existing FRAM tables!!!"
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Tan
        Me.Button1.Font = New System.Drawing.Font("Verdana", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(409, 340)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(95, 36)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Done"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.Color.White
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(32, 12)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(411, 28)
        Me.TextBox3.TabIndex = 8
        Me.TextBox3.Text = "Select Additional Tables for Transfer"
        '
        'FVS_BPTransfer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(532, 398)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "FVS_BPTransfer"
        Me.Text = "FVS_BPTransfer"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents BPSizeLimitchk As System.Windows.Forms.CheckBox
    Friend WithEvents Stockschk As System.Windows.Forms.CheckBox
    Friend WithEvents Fisherieschk As System.Windows.Forms.CheckBox
    Friend WithEvents TimeStepschk As System.Windows.Forms.CheckBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
End Class
