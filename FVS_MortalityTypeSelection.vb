
Public Class FVS_MortalityTypeSelection

   Private Sub MortalityTypeListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MortalityTypeListBox.Click
      MortalityType = MortalityTypeListBox.SelectedIndex + 1
      If MortalityType = 7 Then
         If Not (ReportNumber = 2 And SpeciesName = "CHINOOK") Then
            MsgBox("Brood Year Style for CHINOOK Only!" & vbCrLf & "Please Re-Do Your Choice", MsgBoxStyle.OkOnly)
            Exit Sub
         End If
      End If
      Me.Close()
      FVS_ReportSelection.Visible = True
   End Sub

   Private Sub FVS_MortalityTypeSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 789
      FormWidth = 692
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
         If FVS_MortalityTypeSelection_ReSize = False Then
            Resize_Form(Me)
            FVS_MortalityTypeSelection_ReSize = True
         End If
      End If

      MortalityTypeListBox.Items.Clear()
      Select Case ReportNumber
         Case 1, 2
            Dim MortalityTypeList(7) As String
            MortalityTypeList(1) = "Landed Catch"
            MortalityTypeList(2) = "Shaker + DropOff Mortality"
            MortalityTypeList(3) = "Non-Retention Mortality"
            MortalityTypeList(4) = "Landed Catch plus Non-Retention"
            MortalityTypeList(5) = "Total Mortality"
            MortalityTypeList(6) = "All Reports Selected"
            For Fish As Integer = 1 To 6
               MortalityTypeListBox.Items.Add(MortalityTypeList(Fish))
            Next
            If ReportNumber = 2 Then
               MortalityTypeList(7) = "Brood Year Style (CHINOOK)"
               MortalityTypeListBox.Items.Add(MortalityTypeList(7))
            Else
               MortalityTypeList(7) = ""
            End If
         Case 13
            '- ER Distribution Report
            Dim MortalityTypeList(7) As String
            MortalityTypeList(1) = "Landed Catch"
            MortalityTypeList(2) = "Total Mortality"
            MortalityTypeList(3) = "AEQ Total Mortality"
            MortalityTypeList(4) = "Landed Catch Plus Escapement"
            MortalityTypeList(5) = "Total Mortality Plus Escapement"
            MortalityTypeList(6) = "AEQ Total Mortality Plus Escapement"
            MortalityTypeList(7) = ""
            For Fish As Integer = 1 To 6
               MortalityTypeListBox.Items.Add(MortalityTypeList(Fish))
            Next
         Case 14
            '- Mortality by Age Report
            Dim MortalityTypeList(7) As String
            MortalityTypeList(1) = "Landed Catch"
            MortalityTypeList(2) = "Total Mortality"
            MortalityTypeList(3) = "AEQ Total Mortality"
            MortalityTypeList(4) = ""
            MortalityTypeList(5) = ""
            MortalityTypeList(6) = ""
            MortalityTypeList(7) = ""
            For Fish As Integer = 1 To 3
               MortalityTypeListBox.Items.Add(MortalityTypeList(Fish))
            Next
      End Select
   End Sub

   Private Sub CancelMortButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelMortButton.Click
      Me.Close()
      MortalityType = 0
      FVS_ReportSelection.Visible = True
   End Sub
End Class