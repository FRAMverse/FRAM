Public Class FVS_RunModel

   Private Sub FVS_ModelRun_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      '*********************************************
      'Pete Jan 2015; fix to ensure without a shadow of a doubt that the sublegal ratio process enters into
      ' the SaveDat() subroutine upon completion
      RunTAMMIter = 0
      '*********************************************

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
        'OptionReplaceQuota = False
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
         ChinookSizeLimitCheck.Visible = False
         TammFwsCheck.Enabled = False
         MSFBiasCorrectionCheckBox.Visible = True
         MSFBiasCorrectionCheckBox.Enabled = True
            'MSFBiasCorrectionCheckBox.Checked = True
            OldCohort.Visible = False
            OldCohort.Enabled = False
            chkCoastalIterations.Visible = True
            'MSFBiasFlag = True
            'GetBP.Visible = False

      ElseIf SpeciesName = "CHINOOK" Then
         ChinookBYCheck.Visible = True
         ChinookBYCheck.Enabled = True
         OldTammCheck.Visible = True
         OldTammCheck.Enabled = True
         ChinookSizeLimitCheck.Visible = True
         TammFwsCheck.Visible = True
         TammFwsCheck.Enabled = True
         MSFBiasCorrectionCheckBox.Visible = False
            MSFBiasCorrectionCheckBox.Enabled = False
            OldCohort.Visible = True
            OldCohort.Enabled = True
            MSFBiasFlag = False
            chkCoastalIterations.Visible = False
      End If

      '- Not Supported for now- Feb 2011
        'ReplaceQuotaCheck.Enabled = False
        'ReplaceQuotaCheck.Visible = False

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

      'PPPPPP------------------------------------------------------------------------------------------------------------
      '- Pete 12/13 Code for Executing an integrated update system for external S:L Ratio based EncounterRateAdjustments
      '- Outer flank to original RunModelButtonClick code...
      FinalUpdatePass = False 'This should always be false unless set to true during S:L Ratio Update
      Dim iters As Integer = 1
      Dim c As Integer = 1 'Allows RunModelButton_Click to execute as normal (for coho or non-update Chinook runs)

      If UpdateRunEncounterRateAdjustment = True Then
         'set limit on outer loop iterations (3 is plenty for convergence, but we'll do 4 to be overachievers)
         iters = 4
      End If

      Do While c <= iters
         'PPPPPP--(end of leading Pete 12/13 Block, more at end of loop)-------------------------------------------------


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
                If SizeLimitFix = True Then
                    If RunIDNameSelect.Substring(0, 4) <> "SLC-" Then
                        RunIDNameSelect = "SLC-" & RunIDNameSelect
                    End If
                ElseIf SizeLimitFix = False Then
                    If RunIDNameSelect.Substring(0, 4) = "SLC-" Then
                        RunIDNameSelect = RunIDNameSelect.Substring(4, RunIDNameSelect.Length - 4)
                    End If
                End If
                FVS_MainMenu.RecordSetNameLabel.Text = RunIDNameSelect
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


            '****************Begin PETE-2/27/13-Code for adding Delineation to Model Run Name if Bias Correction Is Applied
            If SpeciesName = "COHO" Then


                If MSFBiasCorrectionCheckBox.Checked = True Then
                    MSFBiasFlag = False
                Else
                    MSFBiasFlag = True
                End If


                If MSFBiasFlag = True Then
                    If RunIDNameSelect.Substring(0, 3) <> "bc-" Then
                        RunIDNameSelect = "bc-" & RunIDNameSelect
                    End If
                ElseIf MSFBiasFlag = False Then
                    If RunIDNameSelect.Substring(0, 3) = "bc-" Then
                        RunIDNameSelect = RunIDNameSelect.Substring(3, RunIDNameSelect.Length - 3)
                    End If
                End If
            End If

            FVS_MainMenu.RecordSetNameLabel.Text = RunIDNameSelect
            '****************End PETE-2/27/13-Code for adding Delineation to Model Run Name if Bias Correction Is Applied


            Call RunCalcs()


            'PPPPPP------------------------------------------------------------------------------------------------------------
            '- Closing flank of Pete 12/13 SL Ratio Code 
            If UpdateRunEncounterRateAdjustment = True And c < iters Then 'don't enter ExternalSubCalcs on last pass
                RunProgressLabel.Text = " Loading Kfat for SLratio update pass #" & c & " ..."
                RunProgressLabel.Refresh()
                Call ExternalSubCalcs(c, iters)
            End If
            c = c + 1
        Loop

      '- Set the UpdateRunEncounterRateAdjustment back to False
      '(should always be false except when set to true during update runs)
      UpdateRunEncounterRateAdjustment = False
      RunTAMMIter = 0 'This Needs to be zero OR things will get goofy on sequential runs.
      'PPPPPP---(end of closing Pete 12/13 Block)------------------------------------------------------------------------


      Me.Close()
      FVS_MainMenu.Visible = True

   End Sub

   Private Sub CancelRunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CancelRunButton.Click
      UpdateRunEncounterRateAdjustment = False
      Me.Close()
      FVS_MainMenu.Visible = True
   End Sub

    'Private Sub ReplaceQuotaCheck_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReplaceQuotaCheck.CheckedChanged
    '    If ReplaceQuotaCheck.Checked = True Then
    '        OptionReplaceQuota = True
    '    Else
    '        OptionReplaceQuota = False
    '    End If
    'End Sub

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

   

    Private Sub ChinookSizeLimitCheck_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChinookSizeLimitCheck.CheckedChanged
        If ChinookSizeLimitCheck.Checked = True Then
            SizeLimitFix = True
        Else
            SizeLimitFix = False
        End If
    End Sub

   Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
      If SpeciesName = "COHO" Then
         MessageBox.Show("The S:L ratio update procedure does not apply to coho")
         Exit Sub
      End If
      Me.Close()
      FVS_AdminPassword.Visible = True
   End Sub

   '-///////////////////////////(*_*)\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ begin
   '- Pete 12/13 Subroutine required to iterate and find the right kFAT values, update RunEncounterRateAdjustment

   Sub ExternalSubCalcs(ByVal c As Integer, ByVal iters As Integer)

      'Now that the run is done, calculate the encounter rate adjustments needed to achieve targets
      'get the external ratio and enc rate adjustment tables for calculations
      Dim dbconn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FVSdatabasename)
      Dim sql As String       'SQL Query text string
      Dim oledbAdapter As OleDb.OleDbDataAdapter
      'start clean
      dsSLquery.Clear()
      'Try
      dbconn.Open()

      'This is creates a simple table of legal and sublegal mortalities for computing SLRatio below.
      sql = "SELECT Mortality.RunID, Mortality.FisheryID, Mortality.TimeStep, Sum(Mortality.MSFShaker) AS MSFSub, " & _
            "Sum(Mortality.MSFEncounter) AS MSFLeg, Sum(Mortality.LandedCatch) AS NSLeg, Sum(Mortality.Shaker) " & _
            "AS NSSub " & _
            "FROM Mortality " & _
            "GROUP BY Mortality.RunID, Mortality.FisheryID, Mortality.TimeStep " & _
            "HAVING (((Mortality.RunID)=" & RunIDSelect & "));"

      oledbAdapter = New OleDb.OleDbDataAdapter(sql, dbconn)
      oledbAdapter.Fill(dsSLquery, "TheMeat")
      oledbAdapter.Dispose()
      dbconn.Close()

      'Catch ex As Exception
      'MsgBox("S:L Ratio Calc Query Bombed!" & vbCr & "Verify that your database contains this correct tables and try again.")
      'End Try

      'Now compute the new Kfats and modify the EncounterRateAdjustment for the next run...
      For F = 1 To NumFish
         For T = 1 To NumSteps
            Dim str As String = "FisheryID = " & F.ToString & " AND TimeStep = " & T.ToString
            Dim dr() As DataRow
            Dim kfatold As Double
            Dim leg, subleg, subrat As Double

            dr = dsSLquery.Tables("TheMeat").Select(str) 'Gets query results for fishery and time step
            'If F = 53 And T = 3 Then
            '   F = 53
            'End If

            If dr.Length = 1 Then
               leg = dr(0)("MSFLeg") + dr(0)("NSLeg")
               subleg = dr(0)("MSFSub") + dr(0)("NSSub")
               For A = MinAge To MaxAge
                  kfatold = Kfat(F, A, T) 'debugging variable
                  If leg = 0 Or subleg = 0 Then
                     Kfat(F, A, T) = 1 'Leave it at 1.00 = no adjustment.
                  Else
                     If TargetRatio(F, A, T) <> -1 Then 'Only compute new adjustments for fisheries providing an estimate of SL ratio 
                        subrat = (subleg / ShakerMortRate(F, T)) / leg '<-FRAM SL Ratio
                        Kfat(F, A, T) = TargetRatio(F, A, T) / subrat
                        RunEncounterRateAdjustment(F, A, T) = RunEncounterRateAdjustment(F, A, T) * Kfat(F, A, T) 'Put it here for correct update/storage for saving
                        EncounterRateAdjustment(A, F, T) = EncounterRateAdjustment(A, F, T) * Kfat(F, A, T) 'Put it here for correct execution during iterations
                        'If (F = 16 Or F = 17) And T = 3 Then
                        'Debug.Print("Fishery =, " & F & ",iteration = ," & c.ToString & " ,Age =," & A.ToString & " ,subrat =," & subrat.ToString & " ,Target =," & TargetRatio(F, A, T).ToString & " ,OldKfat =," & kfatold.ToString & " ,NewKfat =," & Kfat(F, A, T).ToString & " ,EncounterRateAdj =," & EncounterRateAdjustment(A, F, T).ToString & " ,RUNEncounterRateAdj =," & RunEncounterRateAdjustment(F, A, T).ToString)
                        'End If
                     End If
                  End If
               Next
            End If
         Next
      Next

      'Set the boolean to true once FRAM has made all update passes; the last one will just be a calculation pass
      If c = iters - 1 Then
         FinalUpdatePass = True
      End If

   End Sub

   '-///////////////////////////(*_*)\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ end

    Private Sub OldCohort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OldCohort.CheckedChanged
        'will not place time 1 cohort into time 4 for stocks with a missing abundance of age-1 - time 4 age will be zero

        If OldCohort.Checked = True Then
            T4CohortFlag = True
        Else
            T4CohortFlag = False
        End If
    End Sub

    'Private Sub MSFBiasCorrectionCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSFBiasCorrectionCheckBox.CheckedChanged
    '    If MSFBiasCorrectionCheckBox.Checked = True Then
    '        NoMSFBiasCalcs = True
    '    End If


    'End Sub

    Private Sub MSFBiasCorrectionCheckBox_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MSFBiasCorrectionCheckBox.CheckedChanged

    End Sub

    Private Sub chkCoastalIterations_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCoastalIterations.CheckedChanged
        If chkCoastalIterations.Checked = True Then
            CoastalIterations = True
            ReDim FisheryQuotaCompare(NumFish, NumSteps)
        Else
            CoastalIterations = False
        End If
    End Sub
End Class