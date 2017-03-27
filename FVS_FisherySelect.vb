Public Class FVS_FisherySelect

   Private Sub FVS_FisherySelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 827
      FormWidth = 665
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_FisherySelect_ReSize = False Then
            Resize_Form(Me)
            FVS_FisherySelect_ReSize = True
         End If
      End If

      FisheryListBox.Items.Clear()
      For Fish As Integer = 1 To NumFish
         FisheryListBox.Items.Add(FisheryTitle(Fish))
      Next
      If FisherySelectionType = 1 Then
         FisherySelectionTitle.Text = "Fishery Selection"
         FisheryGroupNameTextBox.Visible = False
         FisheryGroupLabel.Visible = False
         FSAllButton.Visible = False
         SelectTypeLabel.Text = "Single Fishery Selection"
      ElseIf FisherySelectionType = 2 Then
         FisherySelectionTitle.Text = "Fishery Selection"
         FisheryGroupNameTextBox.Visible = False
         FisheryGroupLabel.Visible = False
         FSAllButton.Visible = True
         SelectTypeLabel.Text = "Multi-Fishery Selection Allowed"
      ElseIf FisherySelectionType = 3 Then
         FisherySelectionTitle.Text = "Fishery Group Selection"
         FisheryGroupNameTextBox.Visible = True
         FisheryGroupLabel.Visible = True
         FSAllButton.Visible = True
         SelectTypeLabel.Text = "Multi-Fishery Selection Allowed"
      End If
   End Sub

   Private Sub FisheryListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs)
      If FisherySelectionType = 1 Then
         FisheryEditSelection = FisheryListBox.SelectedIndex + 1
         Me.Close()
         If CallingRoutine = 1 Then
            FVS_ReportSelection.Visible = True
         ElseIf CallingRoutine = 2 Then
            FVS_ScreenReports.Visible = True
         ElseIf CallingRoutine = 3 Then
            FisheryEditSelection = FisherySelection(1)
            FVS_StockFisheryScalerEdit.Visible = True
         End If
      Else
         Exit Sub
      End If
   End Sub

   Private Sub FSDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSDoneButton.Click
      NumSelectedFisheries = FisheryListBox.SelectedItems.Count
      ReDim FisherySelection(NumSelectedFisheries)
      For Fish As Integer = 1 To NumSelectedFisheries
         FisherySelection(Fish) = FisheryListBox.SelectedIndices.Item(Fish - 1) + 1
      Next
      If NumSelectedFisheries > 1 And FisherySelectionType = 3 Then
         If FisheryGroupNameTextBox.Text <> "" Then
            FisheryGroupName = FisheryGroupNameTextBox.Text
         Else
            MsgBox("Fishery Group Name Required for this Selection!!" & vbCrLf & _
            "Please Enter Name in TextBox Below", MsgBoxStyle.OkOnly)
            Exit Sub
         End If
      ElseIf NumSelectedFisheries = 1 And FisherySelectionType = 3 Then
         If FisheryGroupNameTextBox.Text <> "" Then
            FisheryGroupName = FisheryGroupNameTextBox.Text
         Else
            FisheryGroupName = FisheryName(FisherySelection(1))
         End If
      End If
      Me.Close()

      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FisheryEditSelection = FisherySelection(1)
         FVS_StockFisheryScalerEdit.Visible = True
      End If

   End Sub

   Private Sub FSCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSCancelButton.Click

      NumSelectedFisheries = 0
      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FisheryEditSelection = 0
         FVS_StockFisheryScalerEdit.Visible = True
      End If
      Me.Close()

   End Sub

   Private Sub FSAllButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSAllButton.Click

      NumSelectedFisheries = NumFish
      ReDim FisherySelection(NumSelectedFisheries)
      For Fish As Integer = 1 To NumSelectedFisheries
         FisherySelection(Fish) = Fish
      Next
      If NumSelectedFisheries > 1 And FisherySelectionType = 3 Then
         If FisheryGroupNameTextBox.Text <> "" Then
            FisheryGroupName = FisheryGroupNameTextBox.Text
         Else
            MsgBox("Fishery Group Name Required for this Selection!!" & vbCrLf & _
            "Please Enter Name in TextBox Below", MsgBoxStyle.OkOnly)
            Exit Sub
         End If
      ElseIf NumSelectedFisheries = 1 And FisherySelectionType = 3 Then
         If FisheryGroupNameTextBox.Text <> "" Then
            FisheryGroupName = FisheryGroupNameTextBox.Text
         Else
            FisheryGroupName = FisheryName(FisherySelection(1))
         End If
      End If
      Me.Close()

      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FisheryEditSelection = FisherySelection(1)
         FVS_StockFisheryScalerEdit.Visible = True
      End If

   End Sub
End Class