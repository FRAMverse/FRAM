Public Class FVS_Output

   Private Sub OTSCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OTSCancelButton.Click
      Me.Close()
      FVS_MainMenu.Visible = True
   End Sub

   Private Sub DriverButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DriverButton.Click
      Me.Visible = False
      FVS_OutputDriver.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub ScreenButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ScreenButton.Click
      Me.Visible = False
      FVS_ScreenReports.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub FVS_Output_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      FormHeight = 606
      FormWidth = 796
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
         If FVS_Output_ReSize = False Then
            Resize_Form(Me)
            FVS_Output_ReSize = True
         End If
      End If
      '- These Reports are very specific for the Database and Spreadsheet ... Used for PSC Coho Tech Comm
      '- Change to TRUE when PSC CoTC Periodic and MU Reports need to be updated
      PSCCohoReportButton.Visible = False
      PSCCohoReportButton.Enabled = False
   End Sub

   Private Sub PSCCohoReportButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PSCCohoReportButton.Click
      If SpeciesName = "CHINOOK" Then Exit Sub
      Me.Cursor = Cursors.WaitCursor
        ' PSCCohoSpreadsheet()
      Me.Cursor = Cursors.Default
   End Sub
End Class