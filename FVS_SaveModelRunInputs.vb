Public Class FVS_SaveModelRunInputs

   Private Sub FVS_SaveModelRunInputs_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 790
      FormWidth = 889
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_SaveModelRunInputs_ReSize = False Then
            Resize_Form(Me)
            FVS_SaveModelRunInputs_ReSize = True
         End If
      End If

      If FVSdatabasename.Length > 50 Then
         DatabaseNameLabel.Text = FVSshortname
      Else
         DatabaseNameLabel.Text = FVSdatabasename
      End If
      ModelRunNameLabel.Text = RunIDNameSelect
   End Sub

   Private Sub SMRCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SMRCancelButton.Click
      Me.Close()
      FVS_MainMenu.Visible = True
   End Sub

   Private Sub SMRSaveButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SMRSaveButton.Click
      RecordsetSelectionType = 5
      FVS_EditRecordSetInfo.ShowDialog()
      If RecordsetSelectionType = -5 Then
         MsgBox("Recordset COPY & SAVE Cancelled", MsgBoxStyle.OkOnly)
         Me.Close()
         FVS_MainMenu.Visible = True
         Exit Sub
      End If
      Me.Cursor = Cursors.WaitCursor
      Call CopyNewRecordset()
      RecordsetSelectionType = 0
      Call SaveModelRunInputs()
      Me.Close()
      If BackFramSave = True Then
         FVS_BackwardsFram.Visible = True
         BackFramSave = False
      Else
         FVS_MainMenu.Visible = True
      End If
      Me.Cursor = Cursors.Default

   End Sub

   Private Sub SMRReplaceButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SMRReplaceButton.Click
      Call SaveModelRunInputs()
      Me.Close()
      If BackFramSave = True Then
         FVS_BackwardsFram.Visible = True
         BackFramSave = False
      Else
         FVS_MainMenu.Visible = True
      End If
   End Sub

End Class