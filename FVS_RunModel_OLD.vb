Public Class FVS_RunModel

   Private Sub FVS_ModelRun_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 721
      FormWidth = 900
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_RunModel_ReSize = False Then
            Resize_Form(Me)
            FVS_RunModel_ReSize = True
         End If
      End If

      If FVSdatabasename.Length > 50 Then
         DatabaseNameLabel.Text = FVSshortname
      Else
         DatabaseNameLabel.Text = FVSdatabasename
      End If
      RecordSetNameLabel.Text = RunIDNameSelect
      TAMMSpreadSheet = ""
      TammNameLabel.Text = TAMMSpreadSheet
      RunProgressLabel.Visible = False
      OptionReplaceQuota = False
      OptionOldTAMMformat = False
      OptionUseTAMMfws = False
      OptionChinookBYAEQ = False
      MRProgressBar.Visible = False
      If SpeciesName = "COHO" Then
         ChinookBYCheck.Visible = False
         ChinookBYCheck.Enabled = False
         OldTammCheck.Visible = False
         OldTammCheck.Enabled = False
         TammFwsCheck.Visible = False
         TammFwsCheck.Enabled = False
      ElseIf SpeciesName = "CHINOOK" Then
         ChinookBYCheck.Visible = True
         ChinookBYCheck.Enabled = True
         OldTammCheck.Visible = True
         OldTammCheck.Enabled = True
         TammFwsCheck.Visible = True
         TammFwsCheck.Enabled = True
      End If

      '- Not Supported for now- Feb 2011
      ReplaceQuotaCheck.Enabled = False
      ReplaceQuotaCheck.Visible = False

   End Sub

   Private Sub SelectTAMMButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectTAMMButton.Click

      Dim OpenTAMMspreadsheet As New OpenFileDialog()
      Dim TAMMSpreadSheetName As String

      TAMMSpreadSheet = ""
      OpenTAMMspreadsheet.Filter = "TAMM Spreadsheets (*.xls*)|*.xls*|All files (*.*)|*.*"
      OpenTAMMspreadsheet.FilterIndex = 1
      OpenTAMMspreadsheet.RestoreDirectory = True

      If OpenTAMMspreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            TAMMSpreadSheet = OpenTAMMspreadsheet.FileName
            TAMMSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(TAMMSpreadSheet).DirectoryName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file selected. Original error: " & Ex.Message)
         End Try
      End If

      If TAMMSpreadSheet = "" Then Exit Sub

      TAMMSpreadSheetName = My.Computer.FileSystem.GetFileInfo(TAMMSpreadSheet).Name
      TammNameLabel.Text = TAMMSpreadSheetName

   End Sub

   Private Sub RunModelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RunModelButton.Click
      Dim result

      '- Set Chinook Tamm Run Option
      TammChinookRunFlag = 0
      If SpeciesName = "CHINOOK" Then
         If OptionOldTAMMformat = True And OptionUseTAMMfws = False Then
            TammChinookRunFlag = 1
         ElseIf OptionOldTAMMformat = False And OptionUseTAMMfws = True Then
            TammChinookRunFlag = 2
         ElseIf OptionOldTAMMformat = True And OptionUseTAMMfws = True Then
            TammChinookRunFlag = 3
         End If
      End If

      '- Check for TAMM Selection
      If TAMMSpreadSheet <> "" Then
         RunTAMMIter = 1
         result = MsgBox("Do You Want to SAVE TAMM Tranfer Values into TAMM SpreadSheet?", MsgBoxStyle.YesNo)
         If result = vbYes Then
            TammTransferSave = True
         Else
            TammTransferSave = False
         End If
      End If
      MRProgressBar.Visible = True

      '- All RunCalcs Routine for all processing
      Call RunCalcs()

      Me.Close()
      FVS_MainMenu.Visible = True

   End Sub

   Private Sub CancelRunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelRunButton.Click
      Me.Close()
      FVS_MainMenu.Visible = True
   End Sub

   Private Sub ReplaceQuotaCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReplaceQuotaCheck.CheckedChanged
      If ReplaceQuotaCheck.Checked = True Then
         OptionReplaceQuota = True
      Else
         OptionReplaceQuota = False
      End If
   End Sub

   Private Sub ChinookBYCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChinookBYCheck.CheckedChanged
      If ChinookBYCheck.Checked = True Then
         OptionChinookBYAEQ = 1
      Else
         OptionChinookBYAEQ = 0
      End If
   End Sub

   Private Sub OldTammCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles OldTammCheck.CheckedChanged
      If OldTammCheck.Checked = True Then
         OptionOldTAMMformat = True
      Else
         OptionOldTAMMformat = False
      End If
   End Sub

   Private Sub TammFwsCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TammFwsCheck.CheckedChanged
      If TammFwsCheck.Checked = True Then
         OptionUseTAMMfws = True
      Else
         OptionUseTAMMfws = False
      End If
   End Sub



   Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles MSFBias.CheckedChanged
      If MSFBias.Checked = True Then
         SelBiasVersion = 5
      Else
         SelBiasVersion = 0
      End If
   End Sub
End Class