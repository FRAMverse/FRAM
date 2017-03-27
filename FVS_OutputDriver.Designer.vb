<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_OutputDriver
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
      Me.Label2 = New System.Windows.Forms.Label()
      Me.SelectDriverButton = New System.Windows.Forms.Button()
      Me.CreateRDButton = New System.Windows.Forms.Button()
      Me.EditRDButton = New System.Windows.Forms.Button()
      Me.DeleteRDButton = New System.Windows.Forms.Button()
      Me.RunRDButton = New System.Windows.Forms.Button()
      Me.RDCancelButton = New System.Windows.Forms.Button()
      Me.ReportDriverSelectionLabel = New System.Windows.Forms.Label()
      Me.SaveDriverFileDialog = New System.Windows.Forms.SaveFileDialog()
      Me.ReadDriverButton = New System.Windows.Forms.Button()
      Me.OpenDriverFileDialog = New System.Windows.Forms.OpenFileDialog()
      Me.ReportSaveFileLabel = New System.Windows.Forms.Label()
      Me.ReportSaveFileTitle = New System.Windows.Forms.Label()
      Me.RecordSetNameLabel = New System.Windows.Forms.Label()
      Me.DatabaseNameLabel = New System.Windows.Forms.Label()
      Me.RecordSetTextLabel = New System.Windows.Forms.Label()
      Me.DatabaseTextLabel = New System.Windows.Forms.Label()
      Me.SuspendLayout()
      '
      'Label1
      '
      Me.Label1.AutoSize = True
      Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label1.Location = New System.Drawing.Point(279, 46)
      Me.Label1.Name = "Label1"
      Me.Label1.Size = New System.Drawing.Size(252, 24)
      Me.Label1.TabIndex = 0
      Me.Label1.Text = "Report Driver File Options"
      '
      'Label2
      '
      Me.Label2.AutoSize = True
      Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.Label2.Location = New System.Drawing.Point(269, 95)
      Me.Label2.Name = "Label2"
      Me.Label2.Size = New System.Drawing.Size(282, 17)
      Me.Label2.TabIndex = 1
      Me.Label2.Text = "Select Driver File then Choose Option"
      '
      'SelectDriverButton
      '
      Me.SelectDriverButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.SelectDriverButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.SelectDriverButton.Location = New System.Drawing.Point(74, 174)
      Me.SelectDriverButton.Name = "SelectDriverButton"
      Me.SelectDriverButton.Size = New System.Drawing.Size(162, 45)
      Me.SelectDriverButton.TabIndex = 2
      Me.SelectDriverButton.Text = "Select Driver"
      Me.SelectDriverButton.UseVisualStyleBackColor = False
      '
      'CreateRDButton
      '
      Me.CreateRDButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.CreateRDButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.CreateRDButton.Location = New System.Drawing.Point(328, 249)
      Me.CreateRDButton.Name = "CreateRDButton"
      Me.CreateRDButton.Size = New System.Drawing.Size(224, 45)
      Me.CreateRDButton.TabIndex = 3
      Me.CreateRDButton.Text = "CREATE Report Driver"
      Me.CreateRDButton.UseVisualStyleBackColor = False
      '
      'EditRDButton
      '
      Me.EditRDButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.EditRDButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.EditRDButton.Location = New System.Drawing.Point(328, 332)
      Me.EditRDButton.Name = "EditRDButton"
      Me.EditRDButton.Size = New System.Drawing.Size(224, 45)
      Me.EditRDButton.TabIndex = 4
      Me.EditRDButton.Text = "EDIT Report Driver"
      Me.EditRDButton.UseVisualStyleBackColor = False
      '
      'DeleteRDButton
      '
      Me.DeleteRDButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.DeleteRDButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DeleteRDButton.Location = New System.Drawing.Point(328, 408)
      Me.DeleteRDButton.Name = "DeleteRDButton"
      Me.DeleteRDButton.Size = New System.Drawing.Size(224, 45)
      Me.DeleteRDButton.TabIndex = 5
      Me.DeleteRDButton.Text = "DELETE Report Driver"
      Me.DeleteRDButton.UseVisualStyleBackColor = False
      '
      'RunRDButton
      '
      Me.RunRDButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.RunRDButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RunRDButton.Location = New System.Drawing.Point(328, 488)
      Me.RunRDButton.Name = "RunRDButton"
      Me.RunRDButton.Size = New System.Drawing.Size(224, 45)
      Me.RunRDButton.TabIndex = 6
      Me.RunRDButton.Text = "RUN Reports"
      Me.RunRDButton.UseVisualStyleBackColor = False
      '
      'RDCancelButton
      '
      Me.RDCancelButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.RDCancelButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RDCancelButton.Location = New System.Drawing.Point(375, 574)
      Me.RDCancelButton.Name = "RDCancelButton"
      Me.RDCancelButton.Size = New System.Drawing.Size(136, 45)
      Me.RDCancelButton.TabIndex = 7
      Me.RDCancelButton.Text = "EXIT"
      Me.RDCancelButton.UseVisualStyleBackColor = False
      '
      'ReportDriverSelectionLabel
      '
      Me.ReportDriverSelectionLabel.AutoSize = True
      Me.ReportDriverSelectionLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReportDriverSelectionLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReportDriverSelectionLabel.Location = New System.Drawing.Point(280, 186)
      Me.ReportDriverSelectionLabel.Name = "ReportDriverSelectionLabel"
      Me.ReportDriverSelectionLabel.Size = New System.Drawing.Size(199, 17)
      Me.ReportDriverSelectionLabel.TabIndex = 8
      Me.ReportDriverSelectionLabel.Text = "No Report Driver Selected"
      '
      'ReadDriverButton
      '
      Me.ReadDriverButton.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReadDriverButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReadDriverButton.Location = New System.Drawing.Point(595, 249)
      Me.ReadDriverButton.Name = "ReadDriverButton"
      Me.ReadDriverButton.Size = New System.Drawing.Size(242, 45)
      Me.ReadDriverButton.TabIndex = 9
      Me.ReadDriverButton.Text = "READ Old Report Driver"
      Me.ReadDriverButton.UseVisualStyleBackColor = False
      '
      'OpenDriverFileDialog
      '
      Me.OpenDriverFileDialog.FileName = "OpenFileDialog1"
      '
      'ReportSaveFileLabel
      '
      Me.ReportSaveFileLabel.AutoSize = True
      Me.ReportSaveFileLabel.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
      Me.ReportSaveFileLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReportSaveFileLabel.Location = New System.Drawing.Point(70, 547)
      Me.ReportSaveFileLabel.Name = "ReportSaveFileLabel"
      Me.ReportSaveFileLabel.Size = New System.Drawing.Size(199, 17)
      Me.ReportSaveFileLabel.TabIndex = 10
      Me.ReportSaveFileLabel.Text = "No Report Driver Selected"
      '
      'ReportSaveFileTitle
      '
      Me.ReportSaveFileTitle.AutoSize = True
      Me.ReportSaveFileTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.ReportSaveFileTitle.Location = New System.Drawing.Point(71, 530)
      Me.ReportSaveFileTitle.Name = "ReportSaveFileTitle"
      Me.ReportSaveFileTitle.Size = New System.Drawing.Size(138, 13)
      Me.ReportSaveFileTitle.TabIndex = 11
      Me.ReportSaveFileTitle.Text = "Report Save File Name"
      '
      'RecordSetNameLabel
      '
      Me.RecordSetNameLabel.AutoSize = True
      Me.RecordSetNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetNameLabel.Location = New System.Drawing.Point(103, 671)
      Me.RecordSetNameLabel.Name = "RecordSetNameLabel"
      Me.RecordSetNameLabel.Size = New System.Drawing.Size(121, 17)
      Me.RecordSetNameLabel.TabIndex = 30
      Me.RecordSetNameLabel.Text = "recordset name"
      '
      'DatabaseNameLabel
      '
      Me.DatabaseNameLabel.AutoSize = True
      Me.DatabaseNameLabel.BackColor = System.Drawing.Color.Yellow
      Me.DatabaseNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseNameLabel.Location = New System.Drawing.Point(103, 637)
      Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
      Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
      Me.DatabaseNameLabel.TabIndex = 29
      Me.DatabaseNameLabel.Text = "database name"
      '
      'RecordSetTextLabel
      '
      Me.RecordSetTextLabel.AutoSize = True
      Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.RecordSetTextLabel.Location = New System.Drawing.Point(20, 671)
      Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
      Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
      Me.RecordSetTextLabel.TabIndex = 28
      Me.RecordSetTextLabel.Text = "RecordSet"
      '
      'DatabaseTextLabel
      '
      Me.DatabaseTextLabel.AutoSize = True
      Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
      Me.DatabaseTextLabel.Location = New System.Drawing.Point(20, 637)
      Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
      Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
      Me.DatabaseTextLabel.TabIndex = 27
      Me.DatabaseTextLabel.Text = "Database"
      '
      'FVS_OutputDriver
      '
      Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
      Me.AutoScroll = True
      Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
      Me.ClientSize = New System.Drawing.Size(877, 700)
      Me.Controls.Add(Me.RecordSetNameLabel)
      Me.Controls.Add(Me.DatabaseNameLabel)
      Me.Controls.Add(Me.RecordSetTextLabel)
      Me.Controls.Add(Me.DatabaseTextLabel)
      Me.Controls.Add(Me.ReportSaveFileTitle)
      Me.Controls.Add(Me.ReportSaveFileLabel)
      Me.Controls.Add(Me.ReadDriverButton)
      Me.Controls.Add(Me.ReportDriverSelectionLabel)
      Me.Controls.Add(Me.RDCancelButton)
      Me.Controls.Add(Me.RunRDButton)
      Me.Controls.Add(Me.DeleteRDButton)
      Me.Controls.Add(Me.EditRDButton)
      Me.Controls.Add(Me.CreateRDButton)
      Me.Controls.Add(Me.SelectDriverButton)
      Me.Controls.Add(Me.Label2)
      Me.Controls.Add(Me.Label1)
      Me.Name = "FVS_OutputDriver"
      Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
      Me.Text = "Report Driver File Options"
      Me.ResumeLayout(False)
      Me.PerformLayout()

   End Sub
   Friend WithEvents Label1 As System.Windows.Forms.Label
   Friend WithEvents Label2 As System.Windows.Forms.Label
   Friend WithEvents SelectDriverButton As System.Windows.Forms.Button
   Friend WithEvents CreateRDButton As System.Windows.Forms.Button
   Friend WithEvents EditRDButton As System.Windows.Forms.Button
   Friend WithEvents DeleteRDButton As System.Windows.Forms.Button
   Friend WithEvents RunRDButton As System.Windows.Forms.Button
   Friend WithEvents RDCancelButton As System.Windows.Forms.Button
   Friend WithEvents ReportDriverSelectionLabel As System.Windows.Forms.Label
   Friend WithEvents SaveDriverFileDialog As System.Windows.Forms.SaveFileDialog
   Friend WithEvents ReadDriverButton As System.Windows.Forms.Button
   Friend WithEvents OpenDriverFileDialog As System.Windows.Forms.OpenFileDialog
   Friend WithEvents ReportSaveFileLabel As System.Windows.Forms.Label
   Friend WithEvents ReportSaveFileTitle As System.Windows.Forms.Label
   Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
   Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
   Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
End Class
