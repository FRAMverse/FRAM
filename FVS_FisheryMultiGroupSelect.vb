Public Class FVS_FisheryMultiGroupSelect

   Private Sub FVS_FisheryMultiGroupSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 849
      FormWidth = 1021
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_FisheryMultiGroupSelect_ReSize = False Then
            Resize_Form(Me)
            FVS_FisheryMultiGroupSelect_ReSize = True
         End If
      End If

      ReDim SelectFisheryList(NumFish, NumFish)
      ReDim FisheryCheckList(NumFish)
      ReDim FisheryGroupNames(NumFish)
      NumGroupFisheries = 0
      NumFisheryGroups = 0
      FisheryListBox.Items.Clear()
      '= Fill Fishery List with leading Fishery Number plus Title
      For Fish = 1 To NumFish
         FisheryListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
      Next
      '- Reset Form Buttons
      FisheryGroupNameTextBox.Visible = True
      FisheryGroupNameTextBox.Text = "Group-1"
      FisheryGroupLabel.Visible = True
      SelectTypeLabel.Text = "Multi-Fishery Selection Allowed"
      FGDoneButton.Visible = True
      FGNextGrpButton.Visible = True
      FGReviewButton.Visible = True
      FGNextReviewButton.Visible = False
      FGExitReviewButton.Visible = False
   End Sub

   Private Sub FisheryListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FisheryListBox.Click
      '- Get Fishery Number from ListBox Item (Number Added when Item Added to ListBox)
      SelectFisheryName = FisheryListBox.Items(FisheryListBox.SelectedIndex)
      SelectFishery = CInt(SelectFisheryName.Substring(0, 3))
      NumGroupFisheries += 1
      If NumFisheryGroups = 0 Then NumFisheryGroups = 1
      '- Add Fishery Selection to current Fishery Group
      SelectFisheryList(NumFisheryGroups, NumGroupFisheries) = SelectFishery
      '- Add Fishery Selection to Selection List & CheckList
      FisherySelectedListBox.Items.Add(FisheryListBox.SelectedItem)
      FisheryCheckList(SelectFishery) = NumFisheryGroups
      '- Remove Fishery Selection from Available List
      FisheryListBox.Items.Clear()
      '= Fill Fishery List without Selected Fisheries
      For Fish = 1 To NumFish
         If FisheryCheckList(Fish) = 0 Then
            FisheryListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
         End If
      Next

   End Sub

   Private Sub FGDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGDoneButton.Click
      '- Check if last Fishery Group saved to arrays
      If NumGroupFisheries <> 0 Then
         If FisheryGroupNameTextBox.Text = "" Then
            MsgBox("Please Enter Fishery Group Name", MsgBoxStyle.OkOnly)
            Exit Sub
         End If
         SelectFisheryList(NumFisheryGroups, 0) = NumGroupFisheries
         FisheryGroupNames(NumFisheryGroups) = FisheryGroupNameTextBox.Text
      Else
         NumFisheryGroups -= 1
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

   Private Sub FGCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGCancelButton.Click

      NumSelectedFisheries = 0
      '- Reset Form Buttons
      FGDoneButton.Visible = True
      FGNextGrpButton.Visible = True
      FGReviewButton.Visible = True
      FGNextReviewButton.Visible = False
      FGExitReviewButton.Visible = False
      FisheryListBox.Visible = True
      FisherySelectedListBox.Items.Clear()
      Me.Close()
      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FisheryEditSelection = 0
         FVS_StockFisheryScalerEdit.Visible = True
      End If

   End Sub

   Private Sub FGNextGrpButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGNextGrpButton.Click
      '- ReSet Form for Next Fishery Group
      If FisheryGroupNameTextBox.Text = "" Then
         MsgBox("Please Enter Fishery Group Name", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      If FisheryGroupNameTextBox.Text.Contains(Chr(34)) Then
         MsgBox("Please Enter Fishery Group Name WITHOUT Quotation Marks!!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      SelectFisheryList(NumFisheryGroups, 0) = NumGroupFisheries
      FisheryGroupNames(NumFisheryGroups) = FisheryGroupNameTextBox.Text
      NumGroupFisheries = 0
      NumFisheryGroups += 1
      FisheryGroupNameTextBox.Text = "Group-" & CStr(NumFisheryGroups)
      FisherySelectedListBox.Items.Clear()
   End Sub

   Private Sub FGReviewButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGReviewButton.Click
      '- Change status of Buttons
      FGDoneButton.Visible = False
      FGNextGrpButton.Visible = False
      FGReviewButton.Visible = False
      FGNextReviewButton.Visible = True
      FGExitReviewButton.Visible = True
      FisheryListBox.Visible = False
      '- Display First Fishery Group
      FisherySelectedListBox.Items.Clear()
      For Fish = 1 To NumFish
         If FisheryCheckList(Fish) = 1 Then
            FisherySelectedListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
         End If
      Next
      FisheryGroupNameTextBox.Text = FisheryGroupNames(1)
      '- Turn control over to review buttons
      SelectFishery = 1
      Exit Sub
   End Sub

   Private Sub FGNextReviewButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGNextReviewButton.Click
      '- Display Next Fishery Group
      If FisheryGroupNameTextBox.Text <> FisheryGroupNames(SelectFishery) Then FisheryGroupNames(SelectFishery) = FisheryGroupNameTextBox.Text
      SelectFishery += 1
      FisherySelectedListBox.Items.Clear()
      For Fish = 1 To NumFish
         If FisheryCheckList(Fish) = SelectFishery Then
            FisherySelectedListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
         End If
      Next
      FisheryGroupNameTextBox.Text = FisheryGroupNames(SelectFishery)
   End Sub

   Private Sub FGExitReviewButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FGExitReviewButton.Click
      '- Return Control to other Form Buttons
      FGDoneButton.Visible = True
      FGNextGrpButton.Visible = True
      FGReviewButton.Visible = True
      FGNextReviewButton.Visible = False
      FGExitReviewButton.Visible = False
      FisheryListBox.Visible = True
      FisherySelectedListBox.Items.Clear()
      FisheryGroupNameTextBox.Text = "Group-" & CStr(NumFisheryGroups)
      Exit Sub
   End Sub

   Private Sub FisherySelectedListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FisherySelectedListBox.Click
      '- Remove a Fishery from Fishery Selected List and Storage Arrays
      SelectFisheryName = FisherySelectedListBox.Items(FisherySelectedListBox.SelectedIndex)
      SelectFishery = CInt(SelectFisheryName.Substring(0, 3))
      FisheryCheckList(SelectFishery) = 0
      SelectFisheryList(NumFisheryGroups, NumGroupFisheries) = 0
      NumGroupFisheries -= 1
      '- Update Selected Fishery List
      FisherySelectedListBox.Items.Clear()
      For Fish = 1 To NumFish
         If FisheryCheckList(Fish) = NumFisheryGroups Then
            FisherySelectedListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
         End If
      Next
      '- Update Available Fishery List
      FisheryListBox.Items.Clear()
      For Fish = 1 To NumFish
         If FisheryCheckList(Fish) = 0 Then
            FisheryListBox.Items.Add(String.Format("{0,3}", Fish.ToString("##0")) & "-" & FisheryTitle(Fish))
         End If
      Next
   End Sub

End Class