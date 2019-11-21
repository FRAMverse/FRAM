Public Class FVS_ScreenReports

   Private Sub RepCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RepCancelButton.Click
      Me.Close()
      FVS_Output.Visible = True
   End Sub

   Private Sub FisheryMortalityCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FisheryMortalityCheckBox.Click
      Me.Visible = False
      ScreenReportType = 1
      FVS_MortalityReport.ShowDialog()
      FisheryMortalityCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub StockCatchCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StockCatchCheckBox.Click
      Me.Visible = False
      ScreenReportType = 2
      StockSelectionType = 1
      FVS_StockSelect.ShowDialog()
      If NumSelectedStocks = 0 Then Exit Sub
      FVS_MortalityReport.ShowDialog()
      StockCatchCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub FVS_ScreenReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 819
      FormWidth = 963
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
         If FVS_ScreenReports_ReSize = False Then
            Resize_Form(Me)
            FVS_ScreenReports_ReSize = True
         End If
      End If

      '- Set Return Point for Stock, Fishery, Mortality SubRoutines 2=Screen Reports
      CallingRoutine = 2
   End Sub

   Private Sub FisheryScalerCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FisheryScalerCheckBox.Click
      Me.Visible = False
      FVS_FisheryScalerScreen.ShowDialog()
      FisheryScalerCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub MSFCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSFCheckBox.Click
      Me.Visible = False
      FVS_SelectiveFisheryScreen.ShowDialog()
      MSFCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub FishStkCompCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FishStkCompCheckBox.Click
      Me.Visible = False
      FVS_FishStkCompScreen.ShowDialog()
      FishStkCompCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub PSCCohoERCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PSCCohoERCheckBox.Click
      If SpeciesName = "CHINOOK" Then
         PSCCohoERCheckBox.CheckState = CheckState.Unchecked
         Exit Sub
      End If
      Me.Visible = False
      FVS_PSCCohoERScreen.ShowDialog()
      PSCCohoERCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub StockPer1000CheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StockPer1000CheckBox.Click
      Me.Visible = False
      FVS_StockImpactsPer1000Screen.ShowDialog()
      StockPer1000CheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

   Private Sub PopStatCheckBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PopStatCheckBox.Click
      Me.Visible = False
      FVS_PopStatScreen.ShowDialog()
      PopStatCheckBox.CheckState = CheckState.Unchecked
      Me.BringToFront()
      Exit Sub
   End Sub

Private Sub StockPer1000CheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles StockPer1000CheckBox.CheckedChanged

End Sub

   Private Sub PopStatCheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles PopStatCheckBox.CheckedChanged

   End Sub

   Private Sub StockCatchCheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles StockCatchCheckBox.CheckedChanged

   End Sub

   Private Sub FisheryMortalityCheckBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles FisheryMortalityCheckBox.CheckedChanged

   End Sub

    Private Sub FisheryScalerCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FisheryScalerCheckBox.CheckedChanged

    End Sub
End Class