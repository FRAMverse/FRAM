Imports System.IO
Imports System.IO.File
Public Class FVS_BackwardsFram

   Public PrnLine As String
   Public bfsw As StreamWriter

   Private Sub FVS_BackwardsFram_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      'Me.AutoScaleMode = Windows.Forms.AutoScaleMode.Dpi
      IterProgressLabel.Visible = False
      IterProgressTextBox.Visible = False
      MSMRecsButton.Visible = False
      MSMRecsButton.Enabled = False
      SaveScalersButton.Visible = False
      SaveScalersButton.Enabled = False
      FormHeight = 724
      FormWidth = 903
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
         If FVS_MainMenu_ReSize = False Then
            Resize_Form(Me)
            FVS_MainMenu_ReSize = True
         End If
        End If
        If SpeciesName = "CHINOOK" Then
            NoMSFBiasCorrection.Visible = False
        End If
   End Sub

   Private Sub TargetEscButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TargetEscButton.Click
      Me.Visible = False
      FVS_BackwardsTarget.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub StartIterationsButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StartIterationsButton.Click
      Dim BackFRAMIteration, NumBackFRAMIterations As Integer

      If Not IsNumeric(NumBackFRAMIterationsTextBox.Text) Then
         MsgBox("Number of Iterations must be Numeric!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      NumBackFRAMIterations = CInt(NumBackFRAMIterationsTextBox.Text)

      '- CHINOOK is done separately with Terminal Runs as Targets 
      If SpeciesName = "CHINOOK" Then
         Call BackChinookFram(BackFRAMIteration, NumBackFRAMIterations)
         Exit Sub
      End If

        If NoMSFBiasCorrection.Checked = True Then
            MSFBiasFlag = False
        Else
            MSFBiasFlag = True
        End If

      ReDim BackScaler(NumStk, NumBackFRAMIterations)
        ReDim BackEsc(NumStk, NumBackFRAMIterations)
        'ReDim BackwardsTarget(NumStk)

      '- Process any Marked/UnMarked Splits (Flag=2)
      Dim SumScalers As Double
      For Stk As Integer = 1 To NumStk
         If BackwardsFlag(Stk) = 2 Then
            If (Stk Mod 2) = 0 Then
               '- Marked Target ... calculate UnMarked Target
               If BackwardsFlag(Stk - 1) = 0 And BackwardsTarget(Stk - 1) = 0 Then
                  SumScalers = StockRecruit(Stk, 3, 1) + StockRecruit(Stk - 1, 3, 1)
                  If SumScalers = 0 Then
                     MsgBox("Error - Backwards Stock FLAG = 2 points to Stock Scalers = ZERO" & vbCrLf & "Stock Name = " & StockTitle(Stk), MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
                  BackwardsTarget(Stk - 1) = BackwardsTarget(Stk) * (StockRecruit(Stk - 1, 3, 1) / SumScalers)
                  BackwardsTarget(Stk) = BackwardsTarget(Stk) * (StockRecruit(Stk, 3, 1) / SumScalers)
                  BackwardsFlag(Stk) = 1
                  BackwardsFlag(Stk - 1) = 1
               Else
                  MsgBox("FLAG = 2 - Error for Backwards FRAM Target Esc" & vbCrLf & "Stock# " & Stk.ToString & " Name = " & StockTitle(Stk), MsgBoxStyle.OkOnly)
                  Exit Sub
               End If
            Else
               '- UnMarked Target ... calculate Marked Target
               If BackwardsFlag(Stk + 1) = 0 And BackwardsTarget(Stk + 1) = 0 Then
                  SumScalers = StockRecruit(Stk, 3, 1) + StockRecruit(Stk + 1, 3, 1)
                  If SumScalers = 0 Then
                     MsgBox("Error - Backwards Stock FLAG = 2 points to Stock Scalers = ZERO" & vbCrLf & "Stock Name = " & StockTitle(Stk), MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
                  BackwardsTarget(Stk) = BackwardsFlag(Stk) * (StockRecruit(Stk, 3, 1) / SumScalers)
                  BackwardsTarget(Stk + 1) = BackwardsFlag(Stk) * (StockRecruit(Stk + 1, 3, 1) / SumScalers)
                  BackwardsFlag(Stk) = 1
                  BackwardsFlag(Stk + 1) = 1
               Else
                  MsgBox("FLAG = 2 - Error for Backwards FRAM Target Esc" & vbCrLf & "Stock# " & Stk.ToString & " Name = " & StockTitle(Stk), MsgBoxStyle.OkOnly)
                  Exit Sub
               End If
            End If
         End If
      Next Stk

      RunBackFramFlag = 1
      Me.Cursor = Cursors.WaitCursor

      '- Open Backwards FRAM Report Text File ... Used for DeBugging Errors
      File_Name = FVSdatabasepath & "\BackFramCheck.Txt"
      If Exists(File_Name) Then Delete(File_Name)
      bfsw = CreateText(File_Name)
      PrnLine = "Backwards FRAM Iteration Calculations " + FVSdatabasepath + "\" & RunIDNameSelect.ToString & " " & Date.Today.ToString
      bfsw.WriteLine(PrnLine)
      PrnLine = RunIDNameSelect.ToString & " -Date- " & Date.Today.ToString
      bfsw.WriteLine(PrnLine)
      bfsw.WriteLine(" ")

      IterProgressLabel.Visible = True
      IterProgressLabel.Refresh()
      IterProgressTextBox.Visible = True
      IterProgressTextBox.BringToFront()

      '- Open FramChk.Txt for RunCalcs (RunBackFRAM)
      File_Name = FVSdatabasepath & "\FramCheck.Txt"
      If Exists(File_Name) Then Delete(File_Name)
      sw = CreateText(File_Name)
      PrnLine = "Command File =" & RunIDNameSelect.ToString & "     " & Date.Today.ToString
      sw.WriteLine(PrnLine)
      sw.WriteLine(" ")

      Dim StartTime, Endtime As Date
      Dim DiffSpan1, DiffSpan2 As TimeSpan

        For BackFRAMIteration = 1 To NumBackFRAMIterations

            StartTime = Date.Now
            '- Update Iteration Label
            IterProgressTextBox.Text = BackFRAMIteration.ToString
            IterProgressTextBox.Refresh()
            '- Print Title for BackFRAM.Prn Report
            PrnLine = "Iteration #" & CStr(BackFRAMIteration) & " "
            bfsw.WriteLine(PrnLine)
            bfsw.WriteLine(" ")
            PrnLine = "Stk# FRAM-Esc Target-Esc ScaleFactor Old-Scalar New-Scalar   Cohort"
            bfsw.WriteLine(PrnLine)

            '- Call RunBackFRAM with RunBackFramFlag ON (=1)
            Call RunBackFRAM()
            Endtime = Date.Now
            DiffSpan1 = Endtime - StartTime
            'StartTime = Endtime

            '- Check RunCalcs Values against Target Escapements
            Call Check_BackwardsTarget(BackFRAMIteration, NumBackFRAMIterations)
            PrnLine = "Iteration " & BackFRAMIteration.ToString & " - BackFram Secs=" & DiffSpan1.Seconds
            bfsw.WriteLine(PrnLine)
            Endtime = Date.Now
            DiffSpan2 = Endtime - StartTime
            PrnLine = "Iteration " & BackFRAMIteration.ToString & " - CheckBFm Secs=" & DiffSpan2.Seconds
            bfsw.WriteLine(PrnLine)

        Next BackFRAMIteration

      IterProgressTextBox.Text = "Save"
      IterProgressTextBox.Refresh()
      SaveDat()

      bfsw.Close()
      sw.Close()

      Me.Cursor = Cursors.Default
      IterProgressLabel.Visible = False
      IterProgressTextBox.Visible = False
      MSMRecsButton.Visible = True
      MSMRecsButton.Enabled = True
      SaveScalersButton.Visible = True
      SaveScalersButton.Enabled = True
      'BackwardsCMDFlag = 1
      RunBackFramFlag = 0

      Me.Visible = False
      FVS_BackwardsResults.ShowDialog()
      Me.BringToFront()
      Exit Sub

   End Sub

   Sub Check_BackwardsTarget(ByVal IterNum As Integer, ByVal BackFRAMIteration As Integer)

      Dim EscDiff, ERTotal As Double
      Dim InitialCohort As Double

      '- Compare FRAM Escapements to Target Escapements
      '  Recalculate Stock Scalars for Next Iteration
      '  Exit if Convergence Criteria is met ... do this later

      Age = 3
      TStep = 5
      For Stk As Integer = 1 To NumStk

         '- Output Report
         PrnLine = String.Format("{0,4}", Stk.ToString("###0"))
         PrnLine &= String.Format("{0,8}", Escape(Stk, Age, TStep).ToString("#######0"))
         PrnLine &= String.Format("{0,8}", BackwardsTarget(Stk).ToString("#######0"))
         If Escape(Stk, Age, TStep) <> 0 Then
            PrnLine &= String.Format("{0,10}", (BackwardsTarget(Stk) / Escape(Stk, Age, TStep)).ToString("####0.0000"))
         Else
            PrnLine &= "         -"
         End If
         PrnLine &= String.Format("{0,11}", StockRecruit(Stk, Age, 1).ToString("###0.0000  "))

         '----------
         BackScaler(Stk, IterNum) = StockRecruit(Stk, Age, 1)
         BackEsc(Stk, IterNum) = Escape(Stk, Age, TStep)
         If BackwardsFlag(Stk) = 0 Then
            'If Escape(Stk, Age, TStep) > 0 And BackwardsTarget(Stk) = 0 Then
            '   '- Target Esc = zero set StkSclr to zero
            '   StockRecruit(Stk, Age, 1) = 0
            'End If
            GoTo NextStockRecruitr
         End If

         InitialCohort = BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)

         If InitialCohort = 0 Then
            StockRecruit(Stk, Age, 1) = 0
         Else
            If IterNum > 1 Then
               '- Reset Zero Stocks to Zero (TAMM Effects)
               If BackScaler(Stk, IterNum - 1) = 0 Then
                  StockRecruit(Stk, Age, 1) = 0
                  GoTo NextStockRecruitr
               End If
            End If
            If Escape(Stk, Age, TStep) < 0 Then
               '- Increase Scalar when Escapement is negative
               If IterNum = 1 Then
                  StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
               Else
                  If BackEsc(Stk, IterNum - 1) < 0 Then
                     StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
                  Else
                     StockRecruit(Stk, Age, 1) = (StockRecruit(Stk, Age, 1) + BackScaler(Stk, IterNum - 1)) / 2
                  End If
               End If
            Else
               '- Increase Cohort Size by Escapement Difference times Survival Rate
               If StockRecruit(Stk, Age, 1) <> 0 And BackwardsTarget(Stk) <> 0 Then
                  EscDiff = BackwardsTarget(Stk) - Escape(Stk, Age, TStep)
                  ERTotal = Escape(Stk, Age, TStep) / InitialCohort * 1.33571
                  If InitialCohort < (Math.Abs(EscDiff) * (1.33571 + ERTotal)) Then
                     '- Check for Negative Scaler
                     If IterNum = 1 Then
                        If EscDiff > 0 Then
                           StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.5
                        Else
                           StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) / 2
                        End If
                     Else
                        If EscDiff > 0 Then
                           If StockRecruit(Stk, Age, 1) < 1 Then
                              StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * (BackwardsTarget(Stk) / Escape(Stk, Age, TStep))
                           Else
                              StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
                           End If
                        Else
                           If StockRecruit(Stk, Age, 1) > 2 Then
                              StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * (BackwardsTarget(Stk) / Escape(Stk, Age, TStep))
                           Else
                              StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) / 1.1
                           End If
                        End If
                     End If
                  Else
                     '- Normal Scaling
                     StockRecruit(Stk, Age, 1) = (InitialCohort + (EscDiff * (1.33571 + ERTotal))) / BaseCohortSize(Stk, Age)
                  End If
               ElseIf BackwardsTarget(Stk) <> 0 And StockRecruit(Stk, Age, 1) = 0 Then
                  '- Target Esc > zero and StkSclr = 0 change SS to one
                  StockRecruit(Stk, Age, 1) = 1
               ElseIf BackwardsTarget(Stk) = 0 And StockRecruit(Stk, Age, 1) <> 0 Then
                  '- Target Esc = zero and StkSclr <> 0 change SS to zero
                  StockRecruit(Stk, Age, 1) = 0
               End If
               'NewStkSize = ((Cohort(Stk, Age, 0, 1) / 0.882496) + ((BackwardsTarget(Stk) - Escape(Stk, Age, TStep)) * 1.22892))
               'OldStkSize = Cohort(Stk, Age, 0, 1) / 0.882496
            End If
         End If

NextStockRecruitr:
         PrnLine &= String.Format("{0,11}", StockRecruit(Stk, Age, 1).ToString("###0.0000  "))
         If BackwardsFlag(Stk) = 0 Then
            PrnLine &= "        *        *"
         Else
            PrnLine &= String.Format("{0,9}", InitialCohort.ToString("########0"))
            PrnLine &= String.Format("{0,9}", (InitialCohort * StockRecruit(Stk, Age, 1)).ToString("########0"))
         End If

         PrnLine &= " - " & StockName(Stk)
         bfsw.WriteLine(PrnLine)

      Next Stk

      'bfsw.Close()

   End Sub

   Private Sub BackChinookFram(ByVal BackFRAMIteration As Integer, ByVal NumBackFRAMIterations As Integer)

      'Dim BackIter, NumBackIterations As Integer
      Dim Result As Integer

      '- ReDim Terminal Stock Arrays
      Call BackChinArrays()

      '- Open Backwards FRAM Report Text File ... Used for DeBugging Errors
      File_Name = FVSdatabasepath & "\BackFramCheck.Txt"
      If Exists(File_Name) Then Delete(File_Name)
      bfsw = CreateText(File_Name)
      PrnLine = "Backwards FRAM Iteration Calculations " + FVSdatabasepath + "\" & RunIDNameSelect.ToString & " " & Date.Today.ToString
      bfsw.WriteLine(PrnLine)
      bfsw.WriteLine(" ")

      'Dim TermChinRun(NumStk + NumChinTermRuns, 5) As Double

      Result = MsgBox("Do You Want to Use TAMI Catches ???", MsgBoxStyle.YesNo)

      If Result = vbYes Then

         Dim OpenTAMMspreadsheet As New OpenFileDialog()

            OpenTAMMspreadsheet.Filter = "All Excel Files (*.xls; *.xlsx; *xlsm)|*.xls; *.xlsx; *xlsm|All files (*.*)|*.*"
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

         TammChinookRunFlag = 1
         If TAMMSpreadSheet <> "" Then Call ReadChinookTAMM()

      End If

      '- Backwards CHINOOK FRAM

      ReDim BackChinScaler(NumStk + NumChinTermRuns, 5, NumBackFRAMIterations)
      ReDim BackChinEsc(NumStk + NumChinTermRuns, 5, NumBackFRAMIterations)

      RunBackFramFlag = 1
      Me.Cursor = Cursors.WaitCursor

      ''- Open Backwards FRAM Report Text File ... Used for DeBugging Errors
      'File_Name = FVSdatabasepath & "\BackFramCheck.Txt"
      'If Exists(File_Name) Then Delete(File_Name)
      'bfsw = CreateText(File_Name)
      'PrnLine = "Backwards FRAM Iteration Calculations " + FVSdatabasepath + "\" & RunIDNameSelect.ToString & " " & Date.Today.ToString
      'bfsw.WriteLine(PrnLine)
      'bfsw.WriteLine(" ")

      IterProgressLabel.Visible = True
      IterProgressLabel.Refresh()
      IterProgressTextBox.Visible = True
      IterProgressTextBox.BringToFront()




      For BackFRAMIteration = 1 To NumBackFRAMIterations
         '- Update Iteration Label
         IterProgressTextBox.Text = BackFRAMIteration.ToString
         IterProgressTextBox.Refresh()
         '- Print Title for BackFRAM.Prn Report
         PrnLine = "Iteration #" & CStr(BackFRAMIteration) & " "
         bfsw.WriteLine(PrnLine)
         bfsw.WriteLine(" ")
         PrnLine = "Stk# Age FRAM-TRS Target-TRS ScaleFactor Old-Scalar New-Scalar   Cohort"
         bfsw.WriteLine(PrnLine)

    '********************************************************
    '**Pete-Jul 2014** See full description of coding changes under Check_CHINOOK_TerminalRun()
    '   The goal wiht this block is simply to start the scalar at 1.0 for all stocks for the first pass.
            If BackFRAMIteration = 1 Then

                If NumStk = 38 Or NumStk = 76 Then
                    NumChinTermRuns = 37
                ElseIf NumStk = 33 Or NumStk = 66 Then
                    NumChinTermRuns = 32
                Else
                    NumChinTermRuns = NumStk / 2 - 1
                End If
                '- Sum Terminal Runs for Flagged Stocks (Combined or Individual)
                For TRun = 1 To NumStk + NumChinTermRuns '(76 Stocks plus 37 Terminal Runs = 113)
                    '- Check if Combined Terminal Run
                    If TermStockNum(TRun) < 0 Then  '- Term Runs nums are negative
                        '- Divide Combined Target into Stock Components for evaluation on First Pass
                        If TermStockNum(TRun) = -2 Then '- Only Nooksack Spring has 4 stocks
                            For Age As Integer = 3 To 5
                                If BackwardsFlag(TRun + 1) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 1), Age, 1) = 1
                                End If
                                If BackwardsFlag(TRun + 2) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 2), Age, 1) = 1
                                End If
                                If BackwardsFlag(TRun + 3) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 3), Age, 1) = 1
                                End If
                                If BackwardsFlag(TRun + 4) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 4), Age, 1) = 1
                                End If
                            Next Age
                        Else
                            '- All Other Stocks
                            For Age As Integer = 3 To 5
                                If BackwardsFlag(TRun + 1) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 1), Age, 1) = 1
                                End If
                                If BackwardsFlag(TRun + 2) = 3 Then
                                    StockRecruit(TermStockNum(TRun + 2), Age, 1) = 1
                                End If
                            Next Age
                        End If
                    End If
                Next TRun
            End If
    '********************************************************



         Call RunCalcs()
         Call Check_CHINOOK_TerminalRun(BackFRAMIteration, NumBackFRAMIterations)

         If BackFRAMIteration = NumBackFRAMIterations Then Exit For
         ReDim LandedCatch(NumStk, MaxAge, NumFish, NumSteps)
         ReDim NonRetention(NumStk, MaxAge, NumFish, NumSteps)
         ReDim Shakers(NumStk, MaxAge, NumFish, NumSteps)
         ReDim DropOff(NumStk, MaxAge, NumFish, NumSteps)
         ReDim Encounters(NumStk, MaxAge, NumFish, NumSteps)
         ReDim MSFLandedCatch(NumStk, MaxAge, NumFish, NumSteps)
         ReDim MSFNonRetention(NumStk, MaxAge, NumFish, NumSteps)
         ReDim MSFShakers(NumStk, MaxAge, NumFish, NumSteps)
         ReDim MSFDropOff(NumStk, MaxAge, NumFish, NumSteps)
         ReDim MSFEncounters(NumStk, MaxAge, NumFish, NumSteps)

      Next BackFRAMIteration

      ChangeStockRecruit = True

      '- Check for Negative Escapements
      'If AnyNeg = 1 Then
      '   MsgBox("Negative Escapements were Detected for this Run" & vbCrLf & "Please check 'FramChk' file for Details")
      'End If

      bfsw.Close()

      Me.Cursor = Cursors.Default
      IterProgressLabel.Visible = False
      IterProgressTextBox.Visible = False
      MSMRecsButton.Visible = True
      MSMRecsButton.Enabled = True
      SaveScalersButton.Visible = True
      SaveScalersButton.Enabled = True
      'BackwardsCMDFlag = 1
      RunBackFramFlag = 0

      Me.Visible = False
      FVS_BackwardsResults.ShowDialog()
      Me.BringToFront()
      Exit Sub

   End Sub

   Private Sub Check_CHINOOK_TerminalRun(ByVal IterNum As Integer, ByVal BackFRAMIteration As Integer)

      Dim EscScaler, StartCohort, EscDiff, ERTotal, TRunSum As Double
      Dim ChinSurvMultTemp As Double 'Temporary survival multiplier to accomdate spring and fall stock multiplier differences
      Dim TRun As Integer
      Dim Stk1, Stk2, Stk3, Stk4 As Integer

      'Pete Jul 2014 -- Code added to allow for exact terminal run target achievement for ALL stocks
      ' Requires a flag = 3 in the terminal run loading template and a value where necessary
      ' The jist of the approach is essentially the same as the old algorithm, i.e., successive forward passes with
      ' adjustment to scalars after each to reduce deviation between target and FRAM t-run  
      ' The major difference is that it now rescales based on the ratio of Target/FRAM between each pass rather
      ' than based on the difference expanded for natural mort and maturation.
      ' This required a few tweaks, all of which are flagged with a note and/or comment that includes '**Pete-Jul 2014**'
      Dim PeteScale, PeteSclTemp '**Pete-Jul 2014** variables for rescaling

      On Error GoTo 0

      '- Compare FRAM Escapements to Target Escapements
      '  Recalculate Stock Scalars for Next Iteration
      '  Exit if Convergence Criteria is met ... do this later

      'BackFRAMStkFlag = 0  '- Save for later

        If NumStk = 38 Or NumStk = 76 Then
            NumChinTermRuns = 37
        ElseIf NumStk = 33 Or NumStk = 66 Then
            NumChinTermRuns = 32
        Else
            NumChinTermRuns = NumStk / 2 - 1
        End If

      ReDim TermChinRun(NumStk + NumChinTermRuns, 5)

      '- Sum Terminal Runs for Flagged Stocks (Combined or Individual)
        For TRun = 1 To NumStk + NumChinTermRuns '(76 Stocks plus 37 Terminal Runs = 113)
            If TRun = 61 Then
                TRun = 61
            End If

            If BackwardsFlag(TRun) = 1 Or BackwardsFlag(TRun) = 3 Then '**Pete-Jul 2014** Also enter this block of code if flag = 3
                '- Check if Combined Terminal Run
                If TermStockNum(TRun) < 0 Then  '- Term Runs nums are negative
                    Call SumChinTermRun(TRun, -1, IterNum)
                    '- Divide Combined Target into Stock Components for evaluation on First Pass
                    If IterNum = 1 Then
                        If TermStockNum(TRun) = -2 Then '- Only Nooksack Spring has 4 stocks
                            Stk1 = TermStockNum(TRun + 1)
                            Stk2 = TermStockNum(TRun + 2)
                            Stk3 = TermStockNum(TRun + 3)
                            Stk4 = TermStockNum(TRun + 4)
                            For Age As Integer = 3 To 5
                                If BackwardsChinook(TRun, Age) = 0 Then
                                    GoTo NextTSprAge
                                Else
                                    If StockRecruit(Stk1, Age, 1) = 0 And StockRecruit(Stk2, Age, 1) = 0 And StockRecruit(Stk3, Age, 1) = 0 And StockRecruit(Stk4, Age, 1) = 0 Then
                                        '- StockRecruitrs Zero but Target is Not
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age) / 4
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age) / 4
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age) / 4
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age) / 4
                                        StockRecruit(Stk1, Age, 1) = 1
                                        StockRecruit(Stk2, Age, 1) = 1
                                        StockRecruit(Stk3, Age, 1) = 1
                                        StockRecruit(Stk4, Age, 1) = 1
                                    Else
                                        '- Divide Target by Ratio of StockRecruitrs ... fix this 2 stocks 2 marks
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk1, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1) + StockRecruit(Stk3, Age, 1) + StockRecruit(Stk4, Age, 1)))
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk2, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1) + StockRecruit(Stk3, Age, 1) + StockRecruit(Stk4, Age, 1)))
                                        BackwardsChinook(TRun + 3, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk3, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1) + StockRecruit(Stk3, Age, 1) + StockRecruit(Stk4, Age, 1)))
                                        BackwardsChinook(TRun + 4, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk4, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1) + StockRecruit(Stk3, Age, 1) + StockRecruit(Stk4, Age, 1)))
                                    End If
                                End If
NextTSprAge:
                            Next Age
                        Else
                            '- All Other Stocks
                            Stk1 = TermStockNum(TRun + 1)
                            Stk2 = TermStockNum(TRun + 2)
                            For Age As Integer = 3 To 5
                                If BackwardsChinook(TRun, Age) = 0 Then
                                    GoTo NextTRunAge
                                Else
                                    If StockRecruit(Stk1, Age, 1) = 0 And StockRecruit(Stk2, Age, 1) = 0 Then
                                        '- StockRecruitrs Zero but Target is Not
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age) / 2
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age) / 2
                                        StockRecruit(Stk1, Age, 1) = 1
                                        StockRecruit(Stk2, Age, 1) = 1
                                    ElseIf StockRecruit(Stk1, Age, 1) = 0 Then
                                        '- One StockRecruitr is Zero
                                        BackwardsChinook(TRun + 1, Age) = 0
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age)
                                    ElseIf StockRecruit(Stk2, Age, 1) = 0 Then
                                        '- Other StockRecruitr is Zero
                                        BackwardsChinook(TRun + 2, Age) = 0
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age)
                                    Else
                                        '- Divide Target by Ratio of StockRecruitrs
                                        BackwardsChinook(TRun + 1, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk1, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1)))
                                        BackwardsChinook(TRun + 2, Age) = BackwardsChinook(TRun, Age) * (StockRecruit(Stk2, Age, 1) / (StockRecruit(Stk1, Age, 1) + StockRecruit(Stk2, Age, 1)))
                                    End If
                                End If
NextTRunAge:
                            Next Age
                        End If
                    End If
                    '- Skip over component stocks when terminal run is flagged (= 1)
                    If TermStockNum(TRun) = -2 Then '- Only Nooksack Spring has 4 stocks
                        TRun = TRun + 4
                    Else
                        TRun = TRun + 2
                    End If
                Else
                    Call SumChinTermRun(TRun, TermStockNum(TRun), IterNum)
                End If
            ElseIf BackwardsFlag(TRun) <> 3 Or BackwardsFlag(TRun) <> 1 Then
                For Age As Integer = 3 To 5
                    TermChinRun(TRun, Age) = 0
                Next Age
            End If
        Next TRun

      For TRun = 1 To NumStk + NumChinTermRuns
         If TermStockNum(TRun) < 0 Then GoTo NextTRun '- skip combined Term Runs
         TRunSum = TermChinRun(TRun, 3) + TermChinRun(TRun, 4) + TermChinRun(TRun, 5)
         If TRunSum <> 0 Then
                Stk = TermStockNum(TRun)
                If Stk = 38 Then
                    Stk = 38
                End If
                
            '- Check Terminal Runs against Target Values and ReSet Stock Recruit Scalers
            For Age As Integer = 3 To 5
               '- Output Report
               If Age <> 2 Then
                  PrnLine = String.Format("{0,4}", Stk.ToString("###0"))
                  PrnLine &= String.Format("{0,3}", Age.ToString("##0"))
                  PrnLine &= String.Format("{0,8}", CLng(TermChinRun(TRun, Age)).ToString("#######0"))
                  PrnLine &= String.Format("{0,8}", CLng(BackwardsChinook(TRun, Age)).ToString("#######0"))
                  If TermChinRun(TRun, Age) <> 0 Then
                     PrnLine &= String.Format("{0,10}", (BackwardsChinook(TRun, Age) / TermChinRun(TRun, Age)).ToString("####0.0000"))
                  Else
                     PrnLine &= "         - "
                  End If
                  PrnLine &= String.Format("{0,11}", StockRecruit(Stk, Age, 1).ToString("###0.0000  "))
               End If

               '- Save old scalers and Term Runs
               BackChinScaler(TRun, Age, IterNum) = StockRecruit(Stk, Age, 1)
               BackChinEsc(TRun, Age, IterNum) = TermChinRun(TRun, Age)

               'InitialCohort = BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)
               '- Reset Stock Scalers
               If BaseCohortSize(Stk, Age) = 0 Then
                  StockRecruit(Stk, Age, 1) = 0
               Else
                  If IterNum > 1 Then
                     '- Reset Zero Stocks to Zero (TAMM Effects)
                     If BackChinScaler(TRun, Age, IterNum - 1) = 0 Then
                        StockRecruit(Stk, Age, 1) = 0
                        GoTo NextChinScaler
                     End If
                  End If

                  If TermChinRun(TRun, Age) < 0 Then
                     '- Increase Scalar when Terminal Run is negative
                     If IterNum = 1 Then
                        StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
                            Else
                                If Stk = 38 Then
                                    Stk = 38
                                End If
                                If BackChinEsc(TRun, Age, IterNum - 1) < 0 Then
                                    StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
                                Else
                                    If StockRecruit(Stk, Age, 1) < BackChinScaler(TRun, Age, IterNum - 1) Then
                                        StockRecruit(Stk, Age, 1) = (StockRecruit(Stk, Age, 1) + BackChinScaler(TRun, Age, IterNum - 1)) / 2
                                    Else
                                        StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.1
                                    End If
                                End If
                     End If
                  Else
                     '- Increase Cohort Size by Term Run Difference times Survival Rate
                     If StockRecruit(Stk, Age, 1) <> 0 And BackwardsChinook(TRun, Age) <> 0 Then
                        StartCohort = BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)
                        EscDiff = BackwardsChinook(TRun, Age) - TermChinRun(TRun, Age)
                        PeteScale = BackwardsChinook(TRun, Age) / TermChinRun(TRun, Age) '**Pete-Jul 2014** Compute adjustment factor
                        If TermChinRun(TRun, Age) < 0 Then 'Special case where negative escapement/T-run is achieved.
                            PeteScale = 1.5
                        End If
                        If Stk = 34 And Age = 5 Then
                            Dim whazzup As Integer = 5
                        End If
                        PeteSclTemp = StockRecruit(Stk, Age, 1) '**Pete-Jul 2014** Temporarily store Recruit Scalar from last pass

                        '- Non-Selective Base Period MatRates
                        If NumStk = 33 Or NumStk = 38 Then
                           If Stk >= 25 And Stk <= 26 Then  '- Spring Stocks mature in Time 1
                              TStep = 1
                              ChinSurvMultTemp = ChinSurvMult(Age) 'Pete Jul 2014 - fix to address different NM for spring vs. fall stocks (mature timestep-driven)
                           Else
                              TStep = 3
                              ChinSurvMultTemp = ChinSurvMult(Age) 'Pete Jul 2014 - fix to address different NM for spring vs. fall stocks (mature timestep-driven)
                           End If
                        Else
                           If Stk >= 49 And Stk <= 52 Then  '- Spring Stocks mature in Time 1
                              TStep = 1
                              ChinSurvMultTemp = ChinSurvMult(Age) 'Pete Jul 2014 - fix to address different NM for spring vs. fall stocks (mature timestep-driven)
                           Else
                              TStep = 3
                              ChinSurvMultTemp = ChinSurvMult(Age) 'Pete Jul 2014 - fix to address different NM for spring vs. fall stocks (mature timestep-driven)
                           End If
                        End If
                        ERTotal = (TermChinRun(TRun, Age) / MaturationRate(Stk, Age, TStep)) / StartCohort
                        '- Check if estimates are completely out of expected range
                        If StartCohort < Math.Abs(EscDiff / MaturationRate(Stk, Age, TStep) * (ChinSurvMultTemp + ERTotal)) Then
                           '- Check for Negative Scaler
                           If IterNum = 1 Then
                              If EscDiff > 0 Then
                                 StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.5
                              Else
                                 StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) / 2
                              End If
                           Else
                              If EscDiff > 0 Then
                                 If StockRecruit(Stk, Age, 1) < 1 Then
                                    StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * (BackwardsChinook(TRun, Age) / TermChinRun(TRun, Age))
                                 Else
                                    StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * 1.3
                                 End If
                              Else
                                 If StockRecruit(Stk, Age, 1) > 2 Then
                                    StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) * (BackwardsChinook(TRun, Age) / TermChinRun(TRun, Age))
                                 Else
                                    StockRecruit(Stk, Age, 1) = StockRecruit(Stk, Age, 1) / 1.8
                                 End If
                              End If
                           End If
                        Else
                                    '- Normal Scaling  ********** Main Algorithm of Original Packer Method
                           StockRecruit(Stk, Age, 1) = (StartCohort + (EscDiff / MaturationRate(Stk, Age, TStep) * (ChinSurvMultTemp + ERTotal))) / BaseCohortSize(Stk, Age)

                            '**Pete-Jul 2014** Enter here and compute using 'brute force' method if flag = 3
                            If BackwardsFlag(TRun) = 3 Then
                                StockRecruit(Stk, Age, 1) = PeteScale * PeteSclTemp
                            End If


                        End If



                     ElseIf BackwardsChinook(TRun, Age) <> 0 And StockRecruit(Stk, Age, 1) = 0 Then
                        '- Target Esc > zero and StkSclr = 0 change SS to one
                        StockRecruit(Stk, Age, 1) = 1
                     ElseIf BackwardsChinook(TRun, Age) = 0 And StockRecruit(Stk, Age, 1) <> 0 Then
                        '- Target Esc = zero and StkSclr <> 0 change SS to zero
                        StockRecruit(Stk, Age, 1) = 0
                     End If
                     'NewStkSize = ((Cohort(Stk, Age, 0, 1) / 0.882496) + ((BackwardsChinook(TRun, Age) - TermChinRun(TRun, Age)) * 1.22892))
                     'OldStkSize = Cohort(Stk, Age, 0, 1) / 0.882496
                  End If
               End If
NextChinScaler:
               PrnLine &= String.Format("{0,11}", StockRecruit(Stk, Age, 1).ToString("###0.0000  "))
               'Print #22, Format(Format(StartCohort, "########0"), "@@@@@@@@@  ");
               PrnLine &= String.Format("{0,9}", (BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)).ToString("########0"))
               PrnLine &= StockName(Stk)
               bfsw.WriteLine(PrnLine)
            Next Age
         End If
NextTRun:
      Next TRun

      Exit Sub

   End Sub

   Sub SumChinTermRun(ByVal TermRun As Integer, ByVal Stock As Integer, ByVal IterNumbr As Integer)

      Dim StartNum, EndNum, TSum, I, J As Integer
      '   On Error GoTo BackChinSumErr
      On Error GoTo 0
      If Stock = -1 Then  '- Combined Terminal Run
         StartNum = TermRun + 1
         '- Non-Selective Type Base File
         If NumStk = 33 Or NumStk = 38 Then
            If TermRun = 2 Then
               EndNum = TermRun + 2
            Else
               EndNum = TermRun + 1
            End If
         Else
            If TermRun = 2 Then
               EndNum = TermRun + 4
            Else
               EndNum = TermRun + 2
            End If
         End If
         '- Loop through component stocks of this combined terminal run
         For J = StartNum To EndNum
            Stk = TermStockNum(J)
            TSum = TermRunStock(Stk)
            For Age As Integer = 3 To 5
               '- Sum Escapement
               For TStep As Integer = 1 To 3
                  TermChinRun(J, Age) = TermChinRun(J, Age) + Escape(Stk, Age, TStep)
               Next TStep
                    '- Sum Terminal Fishery Catches
                    For I = 2 To TFish(TSum, 1) + 1
                        '- Loop through stock specific fisheries
                        Fish = TFish(TSum, I)
                        For TStep As Integer = TTime(TSum, 1) To TTime(TSum, 2)
                            TermChinRun(J, Age) = TermChinRun(J, Age) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                        Next TStep
                    Next I
            Next Age
         Next J
      Else  '- Individual Stock Terminal Run
            Stk = Stock
            
         TSum = TermRunStock(Stock)
         For Age As Integer = 3 To 5
            '- Sum Escapement
            For TStep As Integer = 1 To 3
               TermChinRun(TermRun, Age) = TermChinRun(TermRun, Age) + Escape(Stk, Age, TStep)
            Next TStep
                ' - Sum Terminal Fishery Catches
                For I = 2 To TFish(TSum, 1)
                    '- Loop through stock specific fisheries
                    Fish = TFish(TSum, I)
                    For TStep As Integer = TTime(TSum, 1) To TTime(TSum, 2)
                        TermChinRun(TermRun, Age) = TermChinRun(TermRun, Age) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                    Next TStep
                Next I
         Next Age
      End If

        If TermChinRun(TermRun, Age) < 0 Then
            Dim What As Integer
            What = 1
        End If

      Exit Sub

   End Sub

   Public Sub BackChinArrays()

      '- Backwards Chinook Number of Terminal Runs
        If NumStk = 38 Or NumStk = 76 Then
            NumChinTermRuns = 37
        ElseIf NumStk = 33 Or NumStk = 66 Then
            NumChinTermRuns = 32
        Else
            NumChinTermRuns = NumStk / 2 - 1
        End If
        '- Backwards Chinook Terminal Run Names
        Select Case NumStk
            Case 38, 76
                ReDim TermRunName(37)
                TermRunName(1) = "Nook/Samish Fall"
                TermRunName(2) = "Nook Spring"
                TermRunName(3) = "Skagit Su/Fl Fing"
                TermRunName(4) = "Skagit Su/Fl Year"
                TermRunName(5) = "Skagit Sprng Year"
                TermRunName(6) = "Snohom Fall Fing"
                TermRunName(7) = "Snohom Fall Year"
                TermRunName(8) = "Stilla Fall Fing"
                TermRunName(9) = "Tulalip Fall Fing"
                TermRunName(10) = "Mid PS Fall Fing"
                TermRunName(11) = "UW Accelerated"
                TermRunName(12) = "SPS Fall Fing"
                TermRunName(13) = "SPS Fall Year"
                TermRunName(14) = "White Spr Fing"
                TermRunName(15) = "HC Fall Fing"
                TermRunName(16) = "HC Fall Year"
                TermRunName(17) = "JDF Tribs. Fall"
                TermRunName(18) = "White Spr Year"
                TermRunName(19) = "Hoko River"
                TermRunName(20) = "OREGON Tule"
                TermRunName(21) = "WASHINGTON Tule"
                TermRunName(22) = "Lower Col Wild"
                TermRunName(23) = "Bonn. Pool Hat"
                TermRunName(24) = "Col Rvr Summer"
                TermRunName(25) = "Col Rvr UR-Bright"
                TermRunName(26) = "Cowlitz Spring"
                TermRunName(27) = "Willamette Sprg"
                TermRunName(28) = "Snake River Fall"
                TermRunName(29) = "OR No Coast Fall"
                TermRunName(30) = "West Cst Vanc Isl"
                TermRunName(31) = "Fraser Rvr Late"
                TermRunName(32) = "Fraser Rvr Early"
                TermRunName(33) = "Lower Georgia Str"
                TermRunName(34) = "Lower Col Tule Nat"
                TermRunName(35) = "Central Valley"
                TermRunName(36) = "WA North Coast"
                TermRunName(37) = "Willapa Bay"
            Case 33, 66

                ReDim TermRunName(32)
                TermRunName(1) = "Nook/Samish Fall"
                TermRunName(2) = "Nook Spring"
                TermRunName(3) = "Skagit Su/Fl Fing"
                TermRunName(4) = "Skagit Su/Fl Year"
                TermRunName(5) = "Skagit Sprng Year"
                TermRunName(6) = "Snohom Fall Fing"
                TermRunName(7) = "Snohom Fall Year"
                TermRunName(8) = "Stilla Fall Fing"
                TermRunName(9) = "Tulalip Fall Fing"
                TermRunName(10) = "Mid PS Fall Fing"
                TermRunName(11) = "UW Accelerated"
                TermRunName(12) = "SPS Fall Fing"
                TermRunName(13) = "SPS Fall Year"
                TermRunName(14) = "White Spr Fing"
                TermRunName(15) = "HC Fall Fing"
                TermRunName(16) = "HC Fall Year"
                TermRunName(17) = "JDF Tribs. Fall"
                TermRunName(18) = "White Spr Year"
                TermRunName(19) = "OREGON Tule"
                TermRunName(20) = "WASHINGTON Tule"
                TermRunName(21) = "Lower Col Wild"
                TermRunName(22) = "Bonn. Pool Hat"
                TermRunName(23) = "Col Rvr Summer"
                TermRunName(24) = "Col Rvr UR-Bright"
                TermRunName(25) = "Cowlitz Spring"
                TermRunName(26) = "Willamette Sprg"
                TermRunName(27) = "Snake River Fall"
                TermRunName(28) = "OR No Coast Fall"
                TermRunName(29) = "West Cst Vanc Isl"
                TermRunName(30) = "Fraser Rvr Late"
                TermRunName(31) = "Fraser Rvr Early"
                TermRunName(32) = "Lower Georgia Str"
            Case Else
                ReDim TermRunName(NumChinTermRuns)
                For n = 1 To NumChinTermRuns
                    Select Case n
                        Case 1, 2
                            TermRunName(n) = Mid(StockName(n * 2), 3, 20)
                        Case 3 To 17
                            TermRunName(n) = Mid(StockName(n * 2 + 2), 3, 20)
                        Case 18
                            TermRunName(n) = Mid(StockName(n * 2 + 30), 3, 20)
                        Case 19
                            TermRunName(n) = Mid(StockName(n * 2 + 38), 3, 20)
                        Case 20 To 33
                            TermRunName(n) = Mid(StockName(n * 2 - 2), 3, 20)
                        Case 34 To 37
                            TermRunName(n) = Mid(StockName(n * 2), 3, 20)
                        Case Is > 37
                            TermRunName(n) = Mid(StockName(n * 2 + 2), 3, 20)
                    End Select
                Next n
        End Select

        '--- TermRunStock used for Backwards CHINOOK FRAM
        '--- Used to Index Stock to Terminal Run
        If NumStk = 38 Or NumStk = 76 Then
            ReDim TermRunStock(NumStk + 37)
        ElseIf NumStk = 33 Or NumStk = 66 Then
            ReDim TermRunStock(NumStk + 32)
        Else
            ReDim TermRunStock(NumStk + NumStk / 2 - 1)
        End If
        For Stk As Integer = 1 To NumStk
            Select Case NumStk
                Case 33
                    Select Case Stk
                        Case 1, 2
                            TermRunStock(Stk) = Stk
                        Case 3 To 18
                            TermRunStock(Stk) = Stk - 1
                        Case 19 To 32
                            TermRunStock(Stk) = Stk
                        Case 33
                            TermRunStock(Stk) = 18
                    End Select
                Case 38
                    Select Case Stk
                        Case 1, 2, 34, 35, 36, 37
                            TermRunStock(Stk) = Stk
                        Case 3 To 18
                            TermRunStock(Stk) = Stk - 1
                        Case 19 To 32
                            TermRunStock(Stk) = Stk + 1
                        Case 33
                            TermRunStock(Stk) = 18
                        Case 34 To 37
                            TermRunStock(Stk) = Stk
                        Case 38
                            TermRunStock(Stk) = 19
                    End Select
                Case 66
                    Select Case Stk
                        Case 1, 2
                            TermRunStock(Stk) = 1
                        Case 3 To 6
                            TermRunStock(Stk) = 2
                        Case 7, 8
                            TermRunStock(Stk) = 3
                        Case 9, 10
                            TermRunStock(Stk) = 4
                        Case 11, 12
                            TermRunStock(Stk) = 5
                        Case 13, 14
                            TermRunStock(Stk) = 6
                        Case 15, 16
                            TermRunStock(Stk) = 7
                        Case 17, 18
                            TermRunStock(Stk) = 8
                        Case 19, 20
                            TermRunStock(Stk) = 9
                        Case 21, 22
                            TermRunStock(Stk) = 10
                        Case 23, 24
                            TermRunStock(Stk) = 11
                        Case 25, 26
                            TermRunStock(Stk) = 12
                        Case 27, 28
                            TermRunStock(Stk) = 13
                        Case 29, 30
                            TermRunStock(Stk) = 14
                        Case 31, 32
                            TermRunStock(Stk) = 15
                        Case 33, 34
                            TermRunStock(Stk) = 16
                        Case 35, 36
                            TermRunStock(Stk) = 17
                        Case 37, 38
                            TermRunStock(Stk) = 19
                        Case 39, 40
                            TermRunStock(Stk) = 20
                        Case 41, 42
                            TermRunStock(Stk) = 21
                        Case 43, 44
                            TermRunStock(Stk) = 22
                        Case 45, 46
                            TermRunStock(Stk) = 23
                        Case 47, 48
                            TermRunStock(Stk) = 24
                        Case 49, 50
                            TermRunStock(Stk) = 25
                        Case 51, 52
                            TermRunStock(Stk) = 26
                        Case 53, 54
                            TermRunStock(Stk) = 27
                        Case 55, 56
                            TermRunStock(Stk) = 28
                        Case 57, 58
                            TermRunStock(Stk) = 29
                        Case 59, 60
                            TermRunStock(Stk) = 30
                        Case 61, 62
                            TermRunStock(Stk) = 31
                        Case 63, 64
                            TermRunStock(Stk) = 32
                        Case 65, 66
                            TermRunStock(Stk) = 18
                    End Select
                Case 76
                    Select Case Stk
                        Case 1, 2
                            TermRunStock(Stk) = 1
                        Case 3 To 6
                            TermRunStock(Stk) = 2
                        Case 7, 8
                            TermRunStock(Stk) = 3
                        Case 9, 10
                            TermRunStock(Stk) = 4
                        Case 11, 12
                            TermRunStock(Stk) = 5
                        Case 13, 14
                            TermRunStock(Stk) = 6
                        Case 15, 16
                            TermRunStock(Stk) = 7
                        Case 17, 18
                            TermRunStock(Stk) = 8
                        Case 19, 20
                            TermRunStock(Stk) = 9
                        Case 21, 22
                            TermRunStock(Stk) = 10
                        Case 23, 24
                            TermRunStock(Stk) = 11
                        Case 25, 26
                            TermRunStock(Stk) = 12
                        Case 27, 28
                            TermRunStock(Stk) = 13
                        Case 29, 30
                            TermRunStock(Stk) = 14
                        Case 31, 32
                            TermRunStock(Stk) = 15
                        Case 33, 34
                            TermRunStock(Stk) = 16
                        Case 35, 36
                            TermRunStock(Stk) = 17
                        Case 37, 38
                            TermRunStock(Stk) = 20
                        Case 39, 40
                            TermRunStock(Stk) = 21
                        Case 41, 42
                            TermRunStock(Stk) = 22
                        Case 43, 44
                            TermRunStock(Stk) = 23
                        Case 45, 46
                            TermRunStock(Stk) = 24
                        Case 47, 48
                            TermRunStock(Stk) = 25
                        Case 49, 50
                            TermRunStock(Stk) = 26
                        Case 51, 52
                            TermRunStock(Stk) = 27
                        Case 53, 54
                            TermRunStock(Stk) = 28
                        Case 55, 56
                            TermRunStock(Stk) = 29
                        Case 57, 58
                            TermRunStock(Stk) = 30
                        Case 59, 60
                            TermRunStock(Stk) = 31
                        Case 61, 62
                            TermRunStock(Stk) = 32
                        Case 63, 64
                            TermRunStock(Stk) = 33
                        Case 65, 66
                            TermRunStock(Stk) = 18
                        Case 67, 68
                            TermRunStock(Stk) = 34
                        Case 69, 70
                            TermRunStock(Stk) = 35
                        Case 71, 72
                            TermRunStock(Stk) = 36
                        Case 73, 74
                            TermRunStock(Stk) = 37
                        Case 75, 76
                            TermRunStock(Stk) = 19
                    End Select
                Case Is > 76
                    Select Case Stk
                        Case 1, 2
                            TermRunStock(Stk) = 1
                        Case 3 To 6
                            TermRunStock(Stk) = 2
                        Case 7, 8
                            TermRunStock(Stk) = 3
                        Case 9, 10
                            TermRunStock(Stk) = 4
                        Case 11, 12
                            TermRunStock(Stk) = 5
                        Case 13, 14
                            TermRunStock(Stk) = 6
                        Case 15, 16
                            TermRunStock(Stk) = 7
                        Case 17, 18
                            TermRunStock(Stk) = 8
                        Case 19, 20
                            TermRunStock(Stk) = 9
                        Case 21, 22
                            TermRunStock(Stk) = 10
                        Case 23, 24
                            TermRunStock(Stk) = 11
                        Case 25, 26
                            TermRunStock(Stk) = 12
                        Case 27, 28
                            TermRunStock(Stk) = 13
                        Case 29, 30
                            TermRunStock(Stk) = 14
                        Case 31, 32
                            TermRunStock(Stk) = 15
                        Case 33, 34
                            TermRunStock(Stk) = 16
                        Case 35, 36
                            TermRunStock(Stk) = 17
                        Case 37, 38
                            TermRunStock(Stk) = 20
                        Case 39, 40
                            TermRunStock(Stk) = 21
                        Case 41, 42
                            TermRunStock(Stk) = 22
                        Case 43, 44
                            TermRunStock(Stk) = 23
                        Case 45, 46
                            TermRunStock(Stk) = 24
                        Case 47, 48
                            TermRunStock(Stk) = 25
                        Case 49, 50
                            TermRunStock(Stk) = 26
                        Case 51, 52
                            TermRunStock(Stk) = 27
                        Case 53, 54
                            TermRunStock(Stk) = 28
                        Case 55, 56
                            TermRunStock(Stk) = 29
                        Case 57, 58
                            TermRunStock(Stk) = 30
                        Case 59, 60
                            TermRunStock(Stk) = 31
                        Case 61, 62
                            TermRunStock(Stk) = 32
                        Case 63, 64
                            TermRunStock(Stk) = 33
                        Case 65, 66
                            TermRunStock(Stk) = 18
                        Case 67, 68
                            TermRunStock(Stk) = 34
                        Case 69, 70
                            TermRunStock(Stk) = 35
                        Case 71, 72
                            TermRunStock(Stk) = 36
                        Case 73, 74
                            TermRunStock(Stk) = 37
                        Case 75, 76
                            TermRunStock(Stk) = 19
                        Case Is > 76
                            If Stk Mod 2 = 0 Then
                                TermRunStock(Stk) = Stk / 2 - 0.5
                            Else
                                TermRunStock(Stk) = Stk / 2 - 1
                            End If
                    End Select
            End Select
        Next Stk


        '--- TermStockNum for Backwards CHINOOK FRAM
        '--- Used to Index Stocks with defined Terminal Runs
        Select Case NumStk
            Case 33
                ReDim TermStockNum(NumStk + 32)
                TermStockNum(1) = -1
                TermStockNum(2) = 1
                TermStockNum(3) = -2
                TermStockNum(4) = 2
                TermStockNum(5) = 3
                TermStockNum(6) = -3
                TermStockNum(7) = 4
                TermStockNum(8) = -4
                TermStockNum(9) = 5
                TermStockNum(10) = -5
                TermStockNum(11) = 6
                TermStockNum(12) = -6
                TermStockNum(13) = 7
                TermStockNum(14) = -7
                TermStockNum(15) = 8
                TermStockNum(16) = -8
                TermStockNum(17) = 9
                TermStockNum(18) = -9
                TermStockNum(19) = 10
                TermStockNum(20) = -10
                TermStockNum(21) = 11
                TermStockNum(22) = -11
                TermStockNum(23) = 12
                TermStockNum(24) = -12
                TermStockNum(25) = 13
                TermStockNum(26) = -13
                TermStockNum(27) = 14
                TermStockNum(28) = -14
                TermStockNum(29) = 15
                TermStockNum(30) = -15
                TermStockNum(31) = 16
                TermStockNum(32) = -16
                TermStockNum(33) = 17
                TermStockNum(34) = -17
                TermStockNum(35) = 18
                TermStockNum(36) = -18
                TermStockNum(37) = 33
                TermStockNum(38) = -19
                TermStockNum(39) = 38
                TermStockNum(40) = -20
                TermStockNum(41) = 19
                TermStockNum(42) = -21
                TermStockNum(43) = 20
                TermStockNum(44) = -22
                TermStockNum(45) = 21
                TermStockNum(46) = -23
                TermStockNum(47) = 22
                TermStockNum(48) = -24
                TermStockNum(49) = 23
                TermStockNum(50) = -25
                TermStockNum(51) = 24
                TermStockNum(52) = -26
                TermStockNum(53) = 25
                TermStockNum(54) = -27
                TermStockNum(55) = 26
                TermStockNum(56) = -28
                TermStockNum(57) = 27
                TermStockNum(58) = -29
                TermStockNum(59) = 28
                TermStockNum(60) = -30
                TermStockNum(61) = 29
                TermStockNum(62) = -31
                TermStockNum(63) = 30
                TermStockNum(64) = -32
                TermStockNum(65) = 31
            Case 38
                ReDim TermStockNum(NumStk + 37)
                TermStockNum(1) = -1
                TermStockNum(2) = 1
                TermStockNum(3) = -2
                TermStockNum(4) = 2
                TermStockNum(5) = 3
                TermStockNum(6) = -3
                TermStockNum(7) = 4
                TermStockNum(8) = -4
                TermStockNum(9) = 5
                TermStockNum(10) = -5
                TermStockNum(11) = 6
                TermStockNum(12) = -6
                TermStockNum(13) = 7
                TermStockNum(14) = -7
                TermStockNum(15) = 8
                TermStockNum(16) = -8
                TermStockNum(17) = 9
                TermStockNum(18) = -9
                TermStockNum(19) = 10
                TermStockNum(20) = -10
                TermStockNum(21) = 11
                TermStockNum(22) = -11
                TermStockNum(23) = 12
                TermStockNum(24) = -12
                TermStockNum(25) = 13
                TermStockNum(26) = -13
                TermStockNum(27) = 14
                TermStockNum(28) = -14
                TermStockNum(29) = 15
                TermStockNum(30) = -15
                TermStockNum(31) = 16
                TermStockNum(32) = -16
                TermStockNum(33) = 17
                TermStockNum(34) = -17
                TermStockNum(35) = 18
                TermStockNum(36) = -18
                TermStockNum(37) = 33
                TermStockNum(38) = -19
                TermStockNum(39) = 38
                TermStockNum(40) = -20
                TermStockNum(41) = 19
                TermStockNum(42) = -21
                TermStockNum(43) = 20
                TermStockNum(44) = -22
                TermStockNum(45) = 21
                TermStockNum(46) = -23
                TermStockNum(47) = 22
                TermStockNum(48) = -24
                TermStockNum(49) = 23
                TermStockNum(50) = -25
                TermStockNum(51) = 24
                TermStockNum(52) = -26
                TermStockNum(53) = 25
                TermStockNum(54) = -27
                TermStockNum(55) = 26
                TermStockNum(56) = -28
                TermStockNum(57) = 27
                TermStockNum(58) = -29
                TermStockNum(59) = 28
                TermStockNum(60) = -30
                TermStockNum(61) = 29
                TermStockNum(62) = -31
                TermStockNum(63) = 30
                TermStockNum(64) = -32
                TermStockNum(65) = 31
                TermStockNum(66) = -33
                TermStockNum(67) = 32
                TermStockNum(68) = -34
                TermStockNum(69) = 34
                TermStockNum(70) = -35
                TermStockNum(71) = 35
                TermStockNum(72) = -36
                TermStockNum(73) = 36
                TermStockNum(74) = -37
                TermStockNum(75) = 37
            Case 66
                ReDim TermStockNum(NumStk + 32)
                TermStockNum(1) = -1
                TermStockNum(2) = 1
                TermStockNum(3) = 2
                TermStockNum(4) = -2
                TermStockNum(5) = 3
                TermStockNum(6) = 4
                TermStockNum(7) = 5
                TermStockNum(8) = 6
                TermStockNum(9) = -3
                TermStockNum(10) = 7
                TermStockNum(11) = 8
                TermStockNum(12) = -4
                TermStockNum(13) = 9
                TermStockNum(14) = 10
                TermStockNum(15) = -5
                TermStockNum(16) = 11
                TermStockNum(17) = 12
                TermStockNum(18) = -6
                TermStockNum(19) = 13
                TermStockNum(20) = 14
                TermStockNum(21) = -7
                TermStockNum(22) = 15
                TermStockNum(23) = 16
                TermStockNum(24) = -8
                TermStockNum(25) = 17
                TermStockNum(26) = 18
                TermStockNum(27) = -9
                TermStockNum(28) = 19
                TermStockNum(29) = 20
                TermStockNum(30) = -10
                TermStockNum(31) = 21
                TermStockNum(32) = 22
                TermStockNum(33) = -11
                TermStockNum(34) = 23
                TermStockNum(35) = 24
                TermStockNum(36) = -12
                TermStockNum(37) = 25
                TermStockNum(38) = 26
                TermStockNum(39) = -13
                TermStockNum(40) = 27
                TermStockNum(41) = 28
                TermStockNum(42) = -14
                TermStockNum(43) = 29
                TermStockNum(44) = 30
                TermStockNum(45) = -15
                TermStockNum(46) = 31
                TermStockNum(47) = 32
                TermStockNum(48) = -16
                TermStockNum(49) = 33
                TermStockNum(50) = 34
                TermStockNum(51) = -17
                TermStockNum(52) = 35
                TermStockNum(53) = 36
                TermStockNum(54) = -18
                TermStockNum(55) = 65
                TermStockNum(56) = 66
                TermStockNum(57) = -19
                TermStockNum(58) = 37
                TermStockNum(59) = 38
                TermStockNum(60) = -20
                TermStockNum(61) = 39
                TermStockNum(62) = 40
                TermStockNum(63) = -21
                TermStockNum(64) = 41
                TermStockNum(65) = 42
                TermStockNum(66) = -22
                TermStockNum(67) = 43
                TermStockNum(68) = 44
                TermStockNum(69) = -23
                TermStockNum(70) = 45
                TermStockNum(71) = 46
                TermStockNum(72) = -24
                TermStockNum(73) = 47
                TermStockNum(74) = 48
                TermStockNum(75) = -25
                TermStockNum(76) = 49
                TermStockNum(77) = 50
                TermStockNum(78) = -26
                TermStockNum(79) = 51
                TermStockNum(80) = 52
                TermStockNum(81) = -27
                TermStockNum(82) = 53
                TermStockNum(83) = 54
                TermStockNum(84) = -28
                TermStockNum(85) = 55
                TermStockNum(86) = 56
                TermStockNum(87) = -29
                TermStockNum(88) = 57
                TermStockNum(89) = 58
                TermStockNum(90) = -30
                TermStockNum(91) = 59
                TermStockNum(92) = 60
                TermStockNum(93) = -31
                TermStockNum(94) = 61
                TermStockNum(95) = 62
                TermStockNum(96) = -32
                TermStockNum(97) = 63
                TermStockNum(98) = 64
            Case 76
                ReDim TermStockNum(NumStk + 37)
                TermStockNum(1) = -1
                TermStockNum(2) = 1
                TermStockNum(3) = 2
                TermStockNum(4) = -2
                TermStockNum(5) = 3
                TermStockNum(6) = 4
                TermStockNum(7) = 5
                TermStockNum(8) = 6
                TermStockNum(9) = -3
                TermStockNum(10) = 7
                TermStockNum(11) = 8
                TermStockNum(12) = -4
                TermStockNum(13) = 9
                TermStockNum(14) = 10
                TermStockNum(15) = -5
                TermStockNum(16) = 11
                TermStockNum(17) = 12
                TermStockNum(18) = -6
                TermStockNum(19) = 13
                TermStockNum(20) = 14
                TermStockNum(21) = -7
                TermStockNum(22) = 15
                TermStockNum(23) = 16
                TermStockNum(24) = -8
                TermStockNum(25) = 17
                TermStockNum(26) = 18
                TermStockNum(27) = -9
                TermStockNum(28) = 19
                TermStockNum(29) = 20
                TermStockNum(30) = -10
                TermStockNum(31) = 21
                TermStockNum(32) = 22
                TermStockNum(33) = -11
                TermStockNum(34) = 23
                TermStockNum(35) = 24
                TermStockNum(36) = -12
                TermStockNum(37) = 25
                TermStockNum(38) = 26
                TermStockNum(39) = -13
                TermStockNum(40) = 27
                TermStockNum(41) = 28
                TermStockNum(42) = -14
                TermStockNum(43) = 29
                TermStockNum(44) = 30
                TermStockNum(45) = -15
                TermStockNum(46) = 31
                TermStockNum(47) = 32
                TermStockNum(48) = -16
                TermStockNum(49) = 33
                TermStockNum(50) = 34
                TermStockNum(51) = -17
                TermStockNum(52) = 35
                TermStockNum(53) = 36
                TermStockNum(54) = -18
                TermStockNum(55) = 65
                TermStockNum(56) = 66
                TermStockNum(57) = -19
                TermStockNum(58) = 75
                TermStockNum(59) = 76
                TermStockNum(60) = -20
                TermStockNum(61) = 37
                TermStockNum(62) = 38
                TermStockNum(63) = -21
                TermStockNum(64) = 39
                TermStockNum(65) = 40
                TermStockNum(66) = -22
                TermStockNum(67) = 41
                TermStockNum(68) = 42
                TermStockNum(69) = -23
                TermStockNum(70) = 43
                TermStockNum(71) = 44
                TermStockNum(72) = -24
                TermStockNum(73) = 45
                TermStockNum(74) = 46
                TermStockNum(75) = -25
                TermStockNum(76) = 47
                TermStockNum(77) = 48
                TermStockNum(78) = -26
                TermStockNum(79) = 49
                TermStockNum(80) = 50
                TermStockNum(81) = -27
                TermStockNum(82) = 51
                TermStockNum(83) = 52
                TermStockNum(84) = -28
                TermStockNum(85) = 53
                TermStockNum(86) = 54
                TermStockNum(87) = -29
                TermStockNum(88) = 55
                TermStockNum(89) = 56
                TermStockNum(90) = -30
                TermStockNum(91) = 57
                TermStockNum(92) = 58
                TermStockNum(93) = -31
                TermStockNum(94) = 59
                TermStockNum(95) = 60
                TermStockNum(96) = -32
                TermStockNum(97) = 61
                TermStockNum(98) = 62
                TermStockNum(99) = -33
                TermStockNum(100) = 63
                TermStockNum(101) = 64
                TermStockNum(102) = -34
                TermStockNum(103) = 67
                TermStockNum(104) = 68
                TermStockNum(105) = -35
                TermStockNum(106) = 69
                TermStockNum(107) = 70
                TermStockNum(108) = -36
                TermStockNum(109) = 71
                TermStockNum(110) = 72
                TermStockNum(111) = -37
                TermStockNum(112) = 73
                TermStockNum(113) = 74
            Case Is > 76
                ReDim TermStockNum(NumStk + NumStk / 2 - 1)
                TermStockNum(1) = -1
                TermStockNum(2) = 1
                TermStockNum(3) = 2
                TermStockNum(4) = -2
                TermStockNum(5) = 3
                TermStockNum(6) = 4
                TermStockNum(7) = 5
                TermStockNum(8) = 6
                TermStockNum(9) = -3
                TermStockNum(10) = 7
                TermStockNum(11) = 8
                TermStockNum(12) = -4
                TermStockNum(13) = 9
                TermStockNum(14) = 10
                TermStockNum(15) = -5
                TermStockNum(16) = 11
                TermStockNum(17) = 12
                TermStockNum(18) = -6
                TermStockNum(19) = 13
                TermStockNum(20) = 14
                TermStockNum(21) = -7
                TermStockNum(22) = 15
                TermStockNum(23) = 16
                TermStockNum(24) = -8
                TermStockNum(25) = 17
                TermStockNum(26) = 18
                TermStockNum(27) = -9
                TermStockNum(28) = 19
                TermStockNum(29) = 20
                TermStockNum(30) = -10
                TermStockNum(31) = 21
                TermStockNum(32) = 22
                TermStockNum(33) = -11
                TermStockNum(34) = 23
                TermStockNum(35) = 24
                TermStockNum(36) = -12
                TermStockNum(37) = 25
                TermStockNum(38) = 26
                TermStockNum(39) = -13
                TermStockNum(40) = 27
                TermStockNum(41) = 28
                TermStockNum(42) = -14
                TermStockNum(43) = 29
                TermStockNum(44) = 30
                TermStockNum(45) = -15
                TermStockNum(46) = 31
                TermStockNum(47) = 32
                TermStockNum(48) = -16
                TermStockNum(49) = 33
                TermStockNum(50) = 34
                TermStockNum(51) = -17
                TermStockNum(52) = 35
                TermStockNum(53) = 36
                TermStockNum(54) = -18
                TermStockNum(55) = 65
                TermStockNum(56) = 66
                TermStockNum(57) = -19
                TermStockNum(58) = 75
                TermStockNum(59) = 76
                TermStockNum(60) = -20
                TermStockNum(61) = 37
                TermStockNum(62) = 38
                TermStockNum(63) = -21
                TermStockNum(64) = 39
                TermStockNum(65) = 40
                TermStockNum(66) = -22
                TermStockNum(67) = 41
                TermStockNum(68) = 42
                TermStockNum(69) = -23
                TermStockNum(70) = 43
                TermStockNum(71) = 44
                TermStockNum(72) = -24
                TermStockNum(73) = 45
                TermStockNum(74) = 46
                TermStockNum(75) = -25
                TermStockNum(76) = 47
                TermStockNum(77) = 48
                TermStockNum(78) = -26
                TermStockNum(79) = 49
                TermStockNum(80) = 50
                TermStockNum(81) = -27
                TermStockNum(82) = 51
                TermStockNum(83) = 52
                TermStockNum(84) = -28
                TermStockNum(85) = 53
                TermStockNum(86) = 54
                TermStockNum(87) = -29
                TermStockNum(88) = 55
                TermStockNum(89) = 56
                TermStockNum(90) = -30
                TermStockNum(91) = 57
                TermStockNum(92) = 58
                TermStockNum(93) = -31
                TermStockNum(94) = 59
                TermStockNum(95) = 60
                TermStockNum(96) = -32
                TermStockNum(97) = 61
                TermStockNum(98) = 62
                TermStockNum(99) = -33
                TermStockNum(100) = 63
                TermStockNum(101) = 64
                TermStockNum(102) = -34
                TermStockNum(103) = 67
                TermStockNum(104) = 68
                TermStockNum(105) = -35
                TermStockNum(106) = 69
                TermStockNum(107) = 70
                TermStockNum(108) = -36
                TermStockNum(109) = 71
                TermStockNum(110) = 72
                TermStockNum(111) = -37
                TermStockNum(112) = 73
                TermStockNum(113) = 74
                TermStockNum(114) = -38
                TermStockNum(115) = 77
                TermStockNum(116) = 78
        End Select

        '- TFish is Array of Terminal Fisheries for each Terminal Run
        ReDim TFish(NumStk / 2, 10)
        TFish(1, 1) = 3 : TFish(1, 2) = 39 : TFish(1, 3) = 40 : TFish(1, 4) = 73
        TFish(2, 1) = 3 : TFish(2, 2) = 39 : TFish(2, 3) = 40 : TFish(2, 4) = 73
        TFish(3, 1) = 3 : TFish(3, 2) = 46 : TFish(3, 3) = 47 : TFish(3, 4) = 73
        TFish(4, 1) = 3 : TFish(4, 2) = 46 : TFish(4, 3) = 47 : TFish(4, 4) = 73
        TFish(5, 1) = 3 : TFish(5, 2) = 46 : TFish(5, 3) = 47 : TFish(5, 4) = 73
        TFish(6, 1) = 1 : TFish(6, 2) = 73
        TFish(7, 1) = 1 : TFish(7, 2) = 73
        TFish(8, 1) = 1 : TFish(8, 2) = 73
        TFish(9, 1) = 4 : TFish(9, 2) = 48 : TFish(9, 3) = 51 : TFish(9, 4) = 52 : TFish(9, 5) = 73
        TFish(10, 1) = 5 : TFish(10, 2) = 60 : TFish(10, 3) = 61 : TFish(10, 4) = 62 : TFish(10, 5) = 63 : TFish(10, 6) = 73
        TFish(11, 1) = 1 : TFish(11, 2) = 73
        TFish(12, 1) = 5 : TFish(12, 2) = 68 : TFish(12, 3) = 69 : TFish(12, 4) = 70 : TFish(12, 5) = 71 : TFish(12, 6) = 73
        TFish(13, 1) = 9 : TFish(13, 2) = 60 : TFish(13, 3) = 61 : TFish(13, 4) = 62 : TFish(13, 5) = 63 : TFish(13, 6) = 68 : TFish(13, 7) = 69 : TFish(13, 8) = 70 : TFish(13, 9) = 71 : TFish(13, 10) = 73
        TFish(14, 1) = 1 : TFish(14, 2) = 73
        TFish(15, 1) = 3 : TFish(15, 2) = 65 : TFish(15, 3) = 66 : TFish(15, 4) = 73
        TFish(16, 1) = 3 : TFish(16, 2) = 65 : TFish(16, 3) = 66 : TFish(16, 4) = 73
        TFish(17, 1) = 1 : TFish(17, 2) = 73
        TFish(18, 1) = 1 : TFish(18, 2) = 73
        TFish(19, 1) = 1 : TFish(19, 2) = 73
        TFish(20, 1) = 3 : TFish(20, 2) = 28 : TFish(20, 3) = 72 : TFish(20, 4) = 73
        TFish(21, 1) = 3 : TFish(21, 2) = 28 : TFish(21, 3) = 72 : TFish(21, 4) = 73
        TFish(22, 1) = 3 : TFish(22, 2) = 28 : TFish(22, 3) = 72 : TFish(22, 4) = 73
        TFish(23, 1) = 3 : TFish(23, 2) = 28 : TFish(23, 3) = 72 : TFish(23, 4) = 73
        TFish(24, 1) = 3 : TFish(24, 2) = 28 : TFish(24, 3) = 72 : TFish(24, 4) = 73
        TFish(25, 1) = 3 : TFish(25, 2) = 28 : TFish(25, 3) = 72 : TFish(25, 4) = 73
        TFish(26, 1) = 3 : TFish(26, 2) = 28 : TFish(26, 3) = 72 : TFish(26, 4) = 73
        TFish(27, 1) = 3 : TFish(27, 2) = 28 : TFish(27, 3) = 72 : TFish(27, 4) = 73
        TFish(28, 1) = 3 : TFish(28, 2) = 28 : TFish(28, 3) = 72 : TFish(28, 4) = 73
        TFish(29, 1) = 2 : TFish(29, 2) = 72 : TFish(29, 3) = 73
        TFish(30, 1) = 2 : TFish(30, 2) = 72 : TFish(30, 3) = 73
        TFish(31, 1) = 2 : TFish(31, 2) = 72 : TFish(31, 3) = 73
        TFish(32, 1) = 2 : TFish(32, 2) = 72 : TFish(32, 3) = 73
        TFish(33, 1) = 2 : TFish(33, 2) = 72 : TFish(33, 3) = 73
        TFish(34, 1) = 3 : TFish(34, 2) = 28 : TFish(34, 3) = 72 : TFish(34, 4) = 73
        TFish(35, 1) = 2 : TFish(35, 2) = 72 : TFish(35, 3) = 73
        TFish(36, 1) = 2 : TFish(36, 2) = 72 : TFish(36, 3) = 73
        TFish(37, 1) = 3 : TFish(37, 2) = 25 : TFish(37, 3) = 72 : TFish(37, 4) = 73
        TFish(38, 1) = 2 : TFish(38, 2) = 72 : TFish(38, 3) = 73

        ''changed 1/12/2015 AHB for new calibration run bkFRAM definitions
        'TFish(1, 1) = 3 : TFish(1, 2) = 39 : TFish(1, 3) = 40 : TFish(1, 4) = 73
        'TFish(2, 1) = 1 : TFish(2, 2) = 73
        'TFish(3, 1) = 1 : TFish(3, 2) = 73
        'TFish(4, 1) = 1 : TFish(4, 2) = 73
        'TFish(5, 1) = 1 : TFish(5, 2) = 73
        'TFish(6, 1) = 1 : TFish(6, 2) = 73
        'TFish(7, 1) = 1 : TFish(7, 2) = 73
        'TFish(8, 1) = 1 : TFish(8, 2) = 73
        'TFish(9, 1) = 4 : TFish(9, 2) = 48 : TFish(9, 3) = 51 : TFish(9, 4) = 52 : TFish(9, 5) = 73
        'TFish(10, 1) = 3 : TFish(10, 2) = 62 : TFish(10, 3) = 63 : TFish(10, 4) = 73
        'TFish(11, 1) = 1 : TFish(11, 2) = 73
        'TFish(12, 1) = 1 : TFish(12, 2) = 73
        'TFish(13, 1) = 3 : TFish(13, 2) = 62 : TFish(13, 3) = 63 : TFish(13, 4) = 73
        'TFish(14, 1) = 1 : TFish(14, 2) = 73
        'TFish(15, 1) = 1 : TFish(15, 2) = 73
        'TFish(16, 1) = 1 : TFish(16, 2) = 73
        'TFish(17, 1) = 1 : TFish(17, 2) = 73
        'TFish(18, 1) = 1 : TFish(18, 2) = 73
        'TFish(19, 1) = 1 : TFish(19, 2) = 73
        'TFish(20, 1) = 3 : TFish(20, 2) = 28 : TFish(20, 3) = 72 : TFish(20, 4) = 73
        'TFish(21, 1) = 3 : TFish(21, 2) = 28 : TFish(21, 3) = 72 : TFish(21, 4) = 73
        'TFish(22, 1) = 3 : TFish(22, 2) = 28 : TFish(22, 3) = 72 : TFish(22, 4) = 73
        'TFish(23, 1) = 3 : TFish(23, 2) = 28 : TFish(23, 3) = 72 : TFish(23, 4) = 73
        'TFish(24, 1) = 3 : TFish(24, 2) = 28 : TFish(24, 3) = 72 : TFish(24, 4) = 73
        'TFish(25, 1) = 3 : TFish(25, 2) = 28 : TFish(25, 3) = 72 : TFish(25, 4) = 73
        'TFish(26, 1) = 3 : TFish(26, 2) = 28 : TFish(26, 3) = 72 : TFish(26, 4) = 73
        'TFish(27, 1) = 3 : TFish(27, 2) = 28 : TFish(27, 3) = 72 : TFish(27, 4) = 73
        'TFish(28, 1) = 3 : TFish(28, 2) = 28 : TFish(28, 3) = 72 : TFish(28, 4) = 73
        'TFish(29, 1) = 2 : TFish(29, 2) = 72 : TFish(29, 3) = 73
        'TFish(30, 1) = 2 : TFish(30, 2) = 72 : TFish(30, 3) = 73
        'TFish(31, 1) = 2 : TFish(31, 2) = 72 : TFish(31, 3) = 73
        'TFish(32, 1) = 2 : TFish(32, 2) = 72 : TFish(32, 3) = 73
        'TFish(33, 1) = 2 : TFish(33, 2) = 72 : TFish(33, 3) = 73
        'TFish(34, 1) = 3 : TFish(34, 2) = 28 : TFish(34, 3) = 72 : TFish(34, 4) = 73
        'TFish(35, 1) = 2 : TFish(35, 2) = 72 : TFish(35, 3) = 73
        'TFish(36, 1) = 2 : TFish(36, 2) = 72 : TFish(36, 3) = 73
        'TFish(37, 1) = 3 : TFish(37, 2) = 25 : TFish(37, 3) = 72 : TFish(37, 4) = 73
        'TFish(38, 1) = 1 : TFish(38, 2) = 73

        'TFish(1, 1) = 1 : TFish(1, 2) = 73
        'TFish(2, 1) = 1 : TFish(2, 2) = 73
        'TFish(3, 1) = 1 : TFish(3, 2) = 73
        'TFish(4, 1) = 1 : TFish(4, 2) = 73
        'TFish(5, 1) = 1 : TFish(5, 2) = 73
        'TFish(6, 1) = 1 : TFish(6, 2) = 73
        'TFish(7, 1) = 1 : TFish(7, 2) = 73
        'TFish(8, 1) = 1 : TFish(8, 2) = 73
        'TFish(9, 1) = 1 : TFish(9, 2) = 73
        'TFish(10, 1) = 1 : TFish(10, 2) = 73
        'TFish(11, 1) = 1 : TFish(11, 2) = 73
        'TFish(12, 1) = 1 : TFish(12, 2) = 73
        'TFish(13, 1) = 1 : TFish(13, 2) = 73
        'TFish(14, 1) = 1 : TFish(14, 2) = 73
        'TFish(15, 1) = 1 : TFish(15, 2) = 73
        'TFish(16, 1) = 1 : TFish(16, 2) = 73
        'TFish(17, 1) = 1 : TFish(17, 2) = 73
        'TFish(18, 1) = 1 : TFish(18, 2) = 73
        'TFish(19, 1) = 1 : TFish(19, 2) = 73
        'TFish(20, 1) = 1 : TFish(20, 2) = 73
        'TFish(21, 1) = 1 : TFish(21, 2) = 73
        'TFish(22, 1) = 1 : TFish(22, 2) = 73
        'TFish(23, 1) = 1 : TFish(23, 2) = 73
        'TFish(24, 1) = 1 : TFish(24, 2) = 73
        'TFish(25, 1) = 1 : TFish(25, 2) = 73
        'TFish(26, 1) = 1 : TFish(26, 2) = 73
        'TFish(27, 1) = 1 : TFish(27, 2) = 73
        'TFish(28, 1) = 1 : TFish(28, 2) = 73
        'TFish(29, 1) = 1 : TFish(29, 2) = 73
        'TFish(30, 1) = 1 : TFish(30, 2) = 73
        'TFish(31, 1) = 1 : TFish(31, 2) = 73
        'TFish(32, 1) = 1 : TFish(32, 2) = 73
        'TFish(33, 1) = 1 : TFish(33, 2) = 73
        'TFish(34, 1) = 1 : TFish(34, 2) = 73
        'TFish(35, 1) = 1 : TFish(35, 2) = 73
        'TFish(36, 1) = 1 : TFish(36, 2) = 73
        'TFish(37, 1) = 1 : TFish(37, 2) = 73


        '- TTime is Array of Terminal Time Steps for Terminal Fisheries above
        ReDim TTime(NumStk / 2, 2)
        TTime(1, 1) = 3 : TTime(1, 2) = 3
        TTime(2, 1) = 2 : TTime(2, 2) = 3
        TTime(3, 1) = 2 : TTime(3, 2) = 3
        TTime(4, 1) = 2 : TTime(4, 2) = 3
        TTime(5, 1) = 2 : TTime(5, 2) = 3
        TTime(6, 1) = 3 : TTime(6, 2) = 3
        TTime(7, 1) = 3 : TTime(7, 2) = 3
        TTime(8, 1) = 3 : TTime(8, 2) = 3
        TTime(9, 1) = 3 : TTime(9, 2) = 3
        TTime(10, 1) = 3 : TTime(10, 2) = 3
        TTime(11, 1) = 3 : TTime(11, 2) = 3
        TTime(12, 1) = 3 : TTime(12, 2) = 3
        TTime(13, 1) = 3 : TTime(13, 2) = 3
        TTime(14, 1) = 2 : TTime(14, 2) = 3
        TTime(15, 1) = 3 : TTime(15, 2) = 3
        TTime(16, 1) = 3 : TTime(16, 2) = 3
        TTime(17, 1) = 3 : TTime(17, 2) = 3
        TTime(18, 1) = 2 : TTime(18, 2) = 3
        TTime(19, 1) = 3 : TTime(19, 2) = 3
        TTime(20, 1) = 3 : TTime(20, 2) = 3
        TTime(21, 1) = 3 : TTime(21, 2) = 3
        TTime(22, 1) = 3 : TTime(22, 2) = 3
        TTime(23, 1) = 3 : TTime(23, 2) = 3
        TTime(24, 1) = 2 : TTime(24, 2) = 3
        TTime(25, 1) = 3 : TTime(25, 2) = 3
        TTime(26, 1) = 1 : TTime(26, 2) = 1
        TTime(27, 1) = 1 : TTime(27, 2) = 1
        TTime(28, 1) = 3 : TTime(28, 2) = 3
        TTime(29, 1) = 3 : TTime(29, 2) = 3
        TTime(30, 1) = 3 : TTime(30, 2) = 3
        TTime(31, 1) = 3 : TTime(31, 2) = 3
        TTime(32, 1) = 3 : TTime(32, 2) = 3
        TTime(33, 1) = 3 : TTime(33, 2) = 3
        TTime(34, 1) = 3 : TTime(34, 2) = 3
        TTime(35, 1) = 3 : TTime(35, 2) = 3
        TTime(36, 1) = 3 : TTime(36, 2) = 3
        TTime(37, 1) = 3 : TTime(37, 2) = 3
        TTime(38, 1) = 3 : TTime(38, 2) = 3

        '- ChinSurvMult is Array of Natural Mortality Multipliers for Backwards Cohort Calculation
        ReDim ChinSurvMult(5)
        ReDim SurvMultSp(5)
        ChinSurvMult(3) = 1.43 '- 0.7 inverse
        ChinSurvMult(4) = 1.25 '- 0.8 inverse
        ChinSurvMult(5) = 1.11 '- 0.9 inverse
        'Pete Jul 2014 fix to spring stocks
        SurvMultSp(3) = 1.212 '0.7 inverse, annual survival, rescaled to 7 month time step
        SurvMultSp(4) = 1.132 '0.5 inverse, annual survival, rescaled to 7 month time step
        SurvMultSp(5) = 1.062 '0.8 inverse, annual survival, rescaled to 7 month time step


    End Sub

   Private Sub ExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitButton.Click
      Me.Close()
      FVS_MainMenu.Visible = True
      Exit Sub
   End Sub

   Private Sub SaveScalersButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveScalersButton.Click
      BackFramSave = True
        Me.Visible = False
        MsgBox("After saving recordset, new recruit scalars will only be saved in temporary memory. For permanent saving, RUN forward and SAVE the run.")
      FVS_SaveModelRunInputs.ShowDialog()
      BackFramSave = False
      Exit Sub
   End Sub

   Private Sub MSMRecsButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSMRecsButton.Click

   End Sub

End Class