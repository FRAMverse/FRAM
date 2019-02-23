<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FVS_FramUtils
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
        Me.FUTitle = New System.Windows.Forms.Label
        Me.RecSetInfoButton = New System.Windows.Forms.Button
        Me.ReadCmdButton = New System.Windows.Forms.Button
        Me.ReadOUTFileButton = New System.Windows.Forms.Button
        Me.FUExitButton = New System.Windows.Forms.Button
        Me.CMDFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.DelRecSetButton = New System.Windows.Forms.Button
        Me.DeleteBPButton = New System.Windows.Forms.Button
        Me.CopyRecordsetButton = New System.Windows.Forms.Button
        Me.ReadTaaEtrsButton = New System.Windows.Forms.Button
        Me.TransferModelRunButton = New System.Windows.Forms.Button
        Me.GetModelRunButton = New System.Windows.Forms.Button
        Me.RecordSetNameLabel = New System.Windows.Forms.Label
        Me.DatabaseNameLabel = New System.Windows.Forms.Label
        Me.RecordSetTextLabel = New System.Windows.Forms.Label
        Me.DatabaseTextLabel = New System.Windows.Forms.Label
        Me.MDBSaveFileDialog = New System.Windows.Forms.SaveFileDialog
        Me.CoweemanButton = New System.Windows.Forms.Button
        Me.btn_Chin2s3s = New System.Windows.Forms.Button
        Me.GetBPTransferBtn = New System.Windows.Forms.Button
        Me.TransferBPBtn = New System.Windows.Forms.Button
        Me.OpenTransferModelRunFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.PassonePasstwoBtn = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'FUTitle
        '
        Me.FUTitle.AutoSize = True
        Me.FUTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 13.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FUTitle.Location = New System.Drawing.Point(344, 18)
        Me.FUTitle.Name = "FUTitle"
        Me.FUTitle.Size = New System.Drawing.Size(140, 24)
        Me.FUTitle.TabIndex = 0
        Me.FUTitle.Text = "FRAM Utilities"
        '
        'RecSetInfoButton
        '
        Me.RecSetInfoButton.BackColor = System.Drawing.Color.SandyBrown
        Me.RecSetInfoButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecSetInfoButton.Location = New System.Drawing.Point(78, 78)
        Me.RecSetInfoButton.Name = "RecSetInfoButton"
        Me.RecSetInfoButton.Size = New System.Drawing.Size(265, 53)
        Me.RecSetInfoButton.TabIndex = 1
        Me.RecSetInfoButton.Text = "Edit Model Run Info"
        Me.RecSetInfoButton.UseVisualStyleBackColor = False
        '
        'ReadCmdButton
        '
        Me.ReadCmdButton.BackColor = System.Drawing.Color.Tan
        Me.ReadCmdButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReadCmdButton.Location = New System.Drawing.Point(78, 381)
        Me.ReadCmdButton.Name = "ReadCmdButton"
        Me.ReadCmdButton.Size = New System.Drawing.Size(265, 53)
        Me.ReadCmdButton.TabIndex = 3
        Me.ReadCmdButton.Text = "Read Old CMD File"
        Me.ReadCmdButton.UseVisualStyleBackColor = False
        '
        'ReadOUTFileButton
        '
        Me.ReadOUTFileButton.BackColor = System.Drawing.Color.CadetBlue
        Me.ReadOUTFileButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReadOUTFileButton.Location = New System.Drawing.Point(503, 78)
        Me.ReadOUTFileButton.Name = "ReadOUTFileButton"
        Me.ReadOUTFileButton.Size = New System.Drawing.Size(265, 53)
        Me.ReadOUTFileButton.TabIndex = 4
        Me.ReadOUTFileButton.Text = "Read Old Base Period"
        Me.ReadOUTFileButton.UseVisualStyleBackColor = False
        '
        'FUExitButton
        '
        Me.FUExitButton.BackColor = System.Drawing.Color.OliveDrab
        Me.FUExitButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FUExitButton.Location = New System.Drawing.Point(307, 618)
        Me.FUExitButton.Name = "FUExitButton"
        Me.FUExitButton.Size = New System.Drawing.Size(265, 53)
        Me.FUExitButton.TabIndex = 5
        Me.FUExitButton.Text = "EXIT"
        Me.FUExitButton.UseVisualStyleBackColor = False
        '
        'CMDFileDialog
        '
        Me.CMDFileDialog.FileName = "OpenFileDialog1"
        '
        'DelRecSetButton
        '
        Me.DelRecSetButton.BackColor = System.Drawing.Color.SandyBrown
        Me.DelRecSetButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DelRecSetButton.Location = New System.Drawing.Point(78, 214)
        Me.DelRecSetButton.Name = "DelRecSetButton"
        Me.DelRecSetButton.Size = New System.Drawing.Size(265, 53)
        Me.DelRecSetButton.TabIndex = 6
        Me.DelRecSetButton.Text = "Delete Model Run"
        Me.DelRecSetButton.UseVisualStyleBackColor = False
        '
        'DeleteBPButton
        '
        Me.DeleteBPButton.BackColor = System.Drawing.Color.CadetBlue
        Me.DeleteBPButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DeleteBPButton.Location = New System.Drawing.Point(503, 137)
        Me.DeleteBPButton.Name = "DeleteBPButton"
        Me.DeleteBPButton.Size = New System.Drawing.Size(265, 53)
        Me.DeleteBPButton.TabIndex = 7
        Me.DeleteBPButton.Text = "Delete Base Period"
        Me.DeleteBPButton.UseVisualStyleBackColor = False
        '
        'CopyRecordsetButton
        '
        Me.CopyRecordsetButton.BackColor = System.Drawing.Color.SandyBrown
        Me.CopyRecordsetButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CopyRecordsetButton.Location = New System.Drawing.Point(78, 146)
        Me.CopyRecordsetButton.Name = "CopyRecordsetButton"
        Me.CopyRecordsetButton.Size = New System.Drawing.Size(265, 53)
        Me.CopyRecordsetButton.TabIndex = 8
        Me.CopyRecordsetButton.Text = "Copy Model Run"
        Me.CopyRecordsetButton.UseVisualStyleBackColor = False
        '
        'ReadTaaEtrsButton
        '
        Me.ReadTaaEtrsButton.BackColor = System.Drawing.Color.BlanchedAlmond
        Me.ReadTaaEtrsButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ReadTaaEtrsButton.Location = New System.Drawing.Point(78, 440)
        Me.ReadTaaEtrsButton.Name = "ReadTaaEtrsButton"
        Me.ReadTaaEtrsButton.Size = New System.Drawing.Size(265, 53)
        Me.ReadTaaEtrsButton.TabIndex = 9
        Me.ReadTaaEtrsButton.Text = "Read TAA/ETRS File"
        Me.ReadTaaEtrsButton.UseVisualStyleBackColor = False
        '
        'TransferModelRunButton
        '
        Me.TransferModelRunButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.TransferModelRunButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TransferModelRunButton.Location = New System.Drawing.Point(503, 440)
        Me.TransferModelRunButton.Name = "TransferModelRunButton"
        Me.TransferModelRunButton.Size = New System.Drawing.Size(265, 53)
        Me.TransferModelRunButton.TabIndex = 10
        Me.TransferModelRunButton.Text = "Transfer Model Runs"
        Me.TransferModelRunButton.UseVisualStyleBackColor = False
        '
        'GetModelRunButton
        '
        Me.GetModelRunButton.BackColor = System.Drawing.SystemColors.GradientInactiveCaption
        Me.GetModelRunButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GetModelRunButton.Location = New System.Drawing.Point(503, 499)
        Me.GetModelRunButton.Name = "GetModelRunButton"
        Me.GetModelRunButton.Size = New System.Drawing.Size(265, 53)
        Me.GetModelRunButton.TabIndex = 11
        Me.GetModelRunButton.Text = "Get Model Run Transfers"
        Me.GetModelRunButton.UseVisualStyleBackColor = False
        '
        'RecordSetNameLabel
        '
        Me.RecordSetNameLabel.AutoSize = True
        Me.RecordSetNameLabel.BackColor = System.Drawing.Color.PaleGoldenrod
        Me.RecordSetNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetNameLabel.Location = New System.Drawing.Point(155, 711)
        Me.RecordSetNameLabel.Name = "RecordSetNameLabel"
        Me.RecordSetNameLabel.Size = New System.Drawing.Size(121, 17)
        Me.RecordSetNameLabel.TabIndex = 30
        Me.RecordSetNameLabel.Text = "recordset name"
        '
        'DatabaseNameLabel
        '
        Me.DatabaseNameLabel.AutoSize = True
        Me.DatabaseNameLabel.BackColor = System.Drawing.Color.PaleGoldenrod
        Me.DatabaseNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseNameLabel.Location = New System.Drawing.Point(155, 677)
        Me.DatabaseNameLabel.Name = "DatabaseNameLabel"
        Me.DatabaseNameLabel.Size = New System.Drawing.Size(119, 17)
        Me.DatabaseNameLabel.TabIndex = 29
        Me.DatabaseNameLabel.Text = "database name"
        '
        'RecordSetTextLabel
        '
        Me.RecordSetTextLabel.AutoSize = True
        Me.RecordSetTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RecordSetTextLabel.Location = New System.Drawing.Point(72, 711)
        Me.RecordSetTextLabel.Name = "RecordSetTextLabel"
        Me.RecordSetTextLabel.Size = New System.Drawing.Size(67, 13)
        Me.RecordSetTextLabel.TabIndex = 28
        Me.RecordSetTextLabel.Text = "RecordSet"
        '
        'DatabaseTextLabel
        '
        Me.DatabaseTextLabel.AutoSize = True
        Me.DatabaseTextLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.DatabaseTextLabel.Location = New System.Drawing.Point(72, 677)
        Me.DatabaseTextLabel.Name = "DatabaseTextLabel"
        Me.DatabaseTextLabel.Size = New System.Drawing.Size(61, 13)
        Me.DatabaseTextLabel.TabIndex = 27
        Me.DatabaseTextLabel.Text = "Database"
        '
        'CoweemanButton
        '
        Me.CoweemanButton.BackColor = System.Drawing.Color.DarkKhaki
        Me.CoweemanButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CoweemanButton.Location = New System.Drawing.Point(503, 366)
        Me.CoweemanButton.Name = "CoweemanButton"
        Me.CoweemanButton.Size = New System.Drawing.Size(265, 53)
        Me.CoweemanButton.TabIndex = 31
        Me.CoweemanButton.Text = "Update COWEEMAN Sheets"
        Me.CoweemanButton.UseVisualStyleBackColor = False
        '
        'btn_Chin2s3s
        '
        Me.btn_Chin2s3s.BackColor = System.Drawing.Color.DarkSeaGreen
        Me.btn_Chin2s3s.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_Chin2s3s.Location = New System.Drawing.Point(78, 322)
        Me.btn_Chin2s3s.Name = "btn_Chin2s3s"
        Me.btn_Chin2s3s.Size = New System.Drawing.Size(265, 53)
        Me.btn_Chin2s3s.TabIndex = 32
        Me.btn_Chin2s3s.Text = "Compute 2s From 3s (Chin)"
        Me.btn_Chin2s3s.UseVisualStyleBackColor = False
        '
        'GetBPTransferBtn
        '
        Me.GetBPTransferBtn.BackColor = System.Drawing.Color.CadetBlue
        Me.GetBPTransferBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GetBPTransferBtn.Location = New System.Drawing.Point(503, 255)
        Me.GetBPTransferBtn.Name = "GetBPTransferBtn"
        Me.GetBPTransferBtn.Size = New System.Drawing.Size(265, 53)
        Me.GetBPTransferBtn.TabIndex = 33
        Me.GetBPTransferBtn.Text = "Get Base Period Transfer"
        Me.GetBPTransferBtn.UseVisualStyleBackColor = False
        '
        'TransferBPBtn
        '
        Me.TransferBPBtn.BackColor = System.Drawing.Color.CadetBlue
        Me.TransferBPBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TransferBPBtn.Location = New System.Drawing.Point(503, 196)
        Me.TransferBPBtn.Name = "TransferBPBtn"
        Me.TransferBPBtn.Size = New System.Drawing.Size(265, 53)
        Me.TransferBPBtn.TabIndex = 34
        Me.TransferBPBtn.Text = "Transfer Base Period"
        Me.TransferBPBtn.UseVisualStyleBackColor = False
        '
        'OpenTransferModelRunFileDialog
        '
        Me.OpenTransferModelRunFileDialog.FileName = "OpenFileDialog1"
        '
        'PassonePasstwoBtn
        '
        Me.PassonePasstwoBtn.BackColor = System.Drawing.Color.DarkSalmon
        Me.PassonePasstwoBtn.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.PassonePasstwoBtn.Location = New System.Drawing.Point(78, 499)
        Me.PassonePasstwoBtn.Name = "PassonePasstwoBtn"
        Me.PassonePasstwoBtn.Size = New System.Drawing.Size(265, 53)
        Me.PassonePasstwoBtn.TabIndex = 35
        Me.PassonePasstwoBtn.Text = "Automate Pass 1 Pass 2 (Chin)"
        Me.PassonePasstwoBtn.UseVisualStyleBackColor = False
        '
        'FVS_FramUtils
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.AutoScroll = True
        Me.BackColor = System.Drawing.Color.PaleGoldenrod
        Me.ClientSize = New System.Drawing.Size(874, 701)
        Me.Controls.Add(Me.PassonePasstwoBtn)
        Me.Controls.Add(Me.TransferBPBtn)
        Me.Controls.Add(Me.GetBPTransferBtn)
        Me.Controls.Add(Me.btn_Chin2s3s)
        Me.Controls.Add(Me.CoweemanButton)
        Me.Controls.Add(Me.RecordSetNameLabel)
        Me.Controls.Add(Me.DatabaseNameLabel)
        Me.Controls.Add(Me.RecordSetTextLabel)
        Me.Controls.Add(Me.DatabaseTextLabel)
        Me.Controls.Add(Me.GetModelRunButton)
        Me.Controls.Add(Me.TransferModelRunButton)
        Me.Controls.Add(Me.ReadTaaEtrsButton)
        Me.Controls.Add(Me.CopyRecordsetButton)
        Me.Controls.Add(Me.DeleteBPButton)
        Me.Controls.Add(Me.DelRecSetButton)
        Me.Controls.Add(Me.FUExitButton)
        Me.Controls.Add(Me.ReadOUTFileButton)
        Me.Controls.Add(Me.ReadCmdButton)
        Me.Controls.Add(Me.RecSetInfoButton)
        Me.Controls.Add(Me.FUTitle)
        Me.Name = "FVS_FramUtils"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FRAM Utilities"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents FUTitle As System.Windows.Forms.Label
    Friend WithEvents RecSetInfoButton As System.Windows.Forms.Button
    Friend WithEvents ReadCmdButton As System.Windows.Forms.Button
    Friend WithEvents ReadOUTFileButton As System.Windows.Forms.Button
    Friend WithEvents FUExitButton As System.Windows.Forms.Button
    Friend WithEvents CMDFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DelRecSetButton As System.Windows.Forms.Button
    Friend WithEvents DeleteBPButton As System.Windows.Forms.Button
    Friend WithEvents CopyRecordsetButton As System.Windows.Forms.Button
    Friend WithEvents ReadTaaEtrsButton As System.Windows.Forms.Button
    Friend WithEvents TransferModelRunButton As System.Windows.Forms.Button
    Friend WithEvents GetModelRunButton As System.Windows.Forms.Button
    Friend WithEvents RecordSetNameLabel As System.Windows.Forms.Label
    Friend WithEvents DatabaseNameLabel As System.Windows.Forms.Label
    Friend WithEvents RecordSetTextLabel As System.Windows.Forms.Label
    Friend WithEvents DatabaseTextLabel As System.Windows.Forms.Label
    Friend WithEvents MDBSaveFileDialog As System.Windows.Forms.SaveFileDialog
    Friend WithEvents CoweemanButton As System.Windows.Forms.Button
    Friend WithEvents btn_Chin2s3s As System.Windows.Forms.Button
    Friend WithEvents GetBPTransferBtn As System.Windows.Forms.Button
    Friend WithEvents TransferBPBtn As System.Windows.Forms.Button
    Friend WithEvents OpenTransferModelRunFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents PassonePasstwoBtn As System.Windows.Forms.Button
End Class
