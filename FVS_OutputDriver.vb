Public Class FVS_OutputDriver

   Private Sub ReadDriverButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadDriverButton.Click
      Dim OpenDriverFileDialog As New OpenFileDialog()

      OldDriverFileName = ""
      OpenDriverFileDialog.Filter = "Driver Files (*.drv)|*.drv|All files (*.*)|*.*"
      OpenDriverFileDialog.FilterIndex = 1
      OpenDriverFileDialog.RestoreDirectory = True
      If OpenDriverFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            OldDriverFileName = OpenDriverFileDialog.FileName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If OldDriverFileName = "" Then Exit Sub
      Call ReadOldDriverFile()
      ReportDriverSelectionLabel.Text = ReportDriverName

   End Sub

   Private Sub RDCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RDCancelButton.Click
      Me.Close()
      FVS_Output.Visible = True
   End Sub

   Private Sub RunRDButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RunRDButton.Click

      If ReportDriverName = "" Then
         MsgBox("Report Driver Must Be Selected First", MsgBoxStyle.OkOnly)
         Exit Sub
      End If

      Dim ExtPos As Integer
      Dim SaveDriverFileDialog As New SaveFileDialog()

      ExtPos = InStr(ReportDriverName, ".")
      If ExtPos <> 0 Then
         SaveDriverFileDialog.FileName = Mid(ReportDriverName, 1, ExtPos - 1) & ".PRN"
      Else
         SaveDriverFileDialog.FileName = ReportDriverName & ".PRN"
      End If
      ReportFileName = ""
      SaveDriverFileDialog.Filter = "Report Print Files (*.prn)|*.prn|All files (*.*)|*.*"
      SaveDriverFileDialog.FilterIndex = 1
      SaveDriverFileDialog.RestoreDirectory = True
      If SaveDriverFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            ReportFileName = SaveDriverFileDialog.FileName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If ReportFileName = "" Then Exit Sub
      ReportSaveFileTitle.Visible = True
      ReportSaveFileLabel.Visible = True
      ReportSaveFileLabel.Text = ReportFileName

      Me.Cursor = Cursors.WaitCursor
      Call RunReportDriver()
      Me.Cursor = Cursors.Default
      Me.BringToFront()

   End Sub

   Private Sub FVS_OutputDriver_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 743
      FormWidth = 895
      If FVSdatabasename.Length > 50 Then
         DatabaseNameLabel.Text = FVSshortname
      Else
         DatabaseNameLabel.Text = FVSdatabasename
      End If
      RecordSetNameLabel.Text = RunIDNameSelect
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_OutputDriver_ReSize = False Then
            Resize_Form(Me)
            FVS_OutputDriver_ReSize = True
         End If
      End If

      ReportSaveFileTitle.Visible = False
      ReportSaveFileLabel.Visible = False
   End Sub

   Private Sub EditRDButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles EditRDButton.Click
      If ReportDriverName = "" Then
         MsgBox("Report Driver Must Be Selected First", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      MsgBox("This Function has not been Implemented yet" & vbCrLf & "Create a New Report Driver to make changes", MsgBoxStyle.OkOnly)
   End Sub

   Private Sub SelectDriverButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectDriverButton.Click
      Me.Visible = False
      DriverSelectionType = 1
      FVS_OutputDriverSelection.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub CreateRDButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CreateRDButton.Click
      Me.Visible = False
      FVS_ReportSelection.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub DeleteRDButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteRDButton.Click
      Me.Visible = False
      DriverSelectionType = 2
      FVS_OutputDriverSelection.ShowDialog()
      Me.BringToFront()
   End Sub
End Class