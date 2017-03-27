<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_ModelRunSelection
   Inherits System.Windows.Forms.Form

   'Form overrides dispose to clean up the component list.
   <System.Diagnostics.DebuggerNonUserCode()> _
   Protected Overrides Sub Dispose(ByVal disposing As Boolean)
      If disposing AndAlso components IsNot Nothing Then
         components.Dispose()
      End If
      MyBase.Dispose(disposing)
   End Sub

   'Required by the Windows Form Designer
   Private components As System.ComponentModel.IContainer

   'NOTE: The following procedure is required by the Windows Form Designer
   'It can be modified using the Windows Form Designer.  
   'Do not modify it using the code editor.
   <System.Diagnostics.DebuggerStepThrough()> _
   Private Sub InitializeComponent()
        Me.RSTitle = New System.Windows.Forms.Label
        Me.CmdCancel = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.TransferButton = New System.Windows.Forms.Button
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.btn_DeleteMulti = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'RSTitle
        '
        Me.RSTitle.AutoSize = True
        Me.RSTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 16.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RSTitle.Location = New System.Drawing.Point(359, 9)
        Me.RSTitle.Name = "RSTitle"
        Me.RSTitle.Size = New System.Drawing.Size(305, 26)
        Me.RSTitle.TabIndex = 0
        Me.RSTitle.Text = "FRAM Model Run Selection"
        '
        'CmdCancel
        '
        Me.CmdCancel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.CmdCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CmdCancel.Location = New System.Drawing.Point(474, 690)
        Me.CmdCancel.Name = "CmdCancel"
        Me.CmdCancel.Size = New System.Drawing.Size(171, 49)
        Me.CmdCancel.TabIndex = 2
        Me.CmdCancel.Text = "CANCEL"
        Me.CmdCancel.UseVisualStyleBackColor = False
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Courier New", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(30, 77)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(432, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Index Species  Title                      Description"
        '
        'TransferButton
        '
        Me.TransferButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.TransferButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TransferButton.Location = New System.Drawing.Point(173, 690)
        Me.TransferButton.Name = "TransferButton"
        Me.TransferButton.Size = New System.Drawing.Size(264, 49)
        Me.TransferButton.TabIndex = 5
        Me.TransferButton.Text = "Transfer Selection Done"
        Me.TransferButton.UseVisualStyleBackColor = False
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ListBox1.Font = New System.Drawing.Font("Courier New", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.ItemHeight = 17
        Me.ListBox1.Location = New System.Drawing.Point(23, 96)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(1069, 565)
        Me.ListBox1.TabIndex = 6
        '
        'btn_DeleteMulti
        '
        Me.btn_DeleteMulti.BackColor = System.Drawing.Color.HotPink
        Me.btn_DeleteMulti.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_DeleteMulti.Location = New System.Drawing.Point(681, 690)
        Me.btn_DeleteMulti.Name = "btn_DeleteMulti"
        Me.btn_DeleteMulti.Size = New System.Drawing.Size(196, 49)
        Me.btn_DeleteMulti.TabIndex = 7
        Me.btn_DeleteMulti.Text = "Delete multiple runs"
        Me.btn_DeleteMulti.UseVisualStyleBackColor = False
        Me.btn_DeleteMulti.Visible = False
        '
        'FVS_ModelRunSelection
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(1135, 770)
        Me.Controls.Add(Me.btn_DeleteMulti)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.TransferButton)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.CmdCancel)
        Me.Controls.Add(Me.RSTitle)
        Me.Name = "FVS_ModelRunSelection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FVS_RecordSet"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
   Friend WithEvents RSTitle As System.Windows.Forms.Label
   Friend WithEvents CmdCancel As System.Windows.Forms.Button
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents TransferButton As System.Windows.Forms.Button
   Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
   Friend WithEvents btn_DeleteMulti As System.Windows.Forms.Button
End Class
