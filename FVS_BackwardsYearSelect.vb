Public Class FVS_BackwardsYearSelect

   Private Sub BFYCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFYCancelButton.Click
      Me.Close()
      BFYearSelection = 0
      FVS_BackwardsTarget.Visible = True
      Exit Sub
   End Sub

   Private Sub FVS_BackwardsYearSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 788
      FormWidth = 565
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
         If FVS_BackwardsYearSelect_ReSize = False Then
            Resize_Form(Me)
            FVS_BackwardsYearSelect_ReSize = True
         End If
      End If

      BFYearCheckedListBox.Items.Clear()
      If BFYearSelectType = 1 Then
         For Stk As Integer = 1 To 50
            If BFEscYears(Stk) = 0 Then Exit Sub
            BFYearCheckedListBox.Items.Add(BFEscYears(Stk))
         Next
      ElseIf BFYearSelectType = 2 Then
         For Stk As Integer = 1 To 50
            If BFCatchYears(Stk) = 0 Then Exit Sub
            BFYearCheckedListBox.Items.Add(BFCatchYears(Stk))
         Next
      End If
   End Sub

   Private Sub BFYearCheckedListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BFYearCheckedListBox.Click
      BFYearSelection = CInt(BFYearCheckedListBox.SelectedItem.ToString)
      Me.Close()
      FVS_BackwardsTarget.Enabled = True
   End Sub
End Class