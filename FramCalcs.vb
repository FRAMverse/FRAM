Imports System
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.IO.File
Imports Microsoft.Office.Interop
Imports System.Data.OleDb
Imports System.Data
Imports System.Globalization

Module FramCalcs
   Public sw As StreamWriter
   Public sr As StreamReader
   Public PrnLine As String
   Public xlApp As Excel.Application
   Public ExcelWasNotRunning As Boolean
   Public WorkBookWasNotOpen As Boolean
   Public xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
   Public xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet

   Sub RunCalcs()

      '- ReDim Calculation Arrays
      ReDim LandedCatch(NumStk, MaxAge, NumFish, NumSteps)
        ReDim NonRetention(NumStk, MaxAge, NumFish, NumSteps)
        ReDim NRLegal(2, NumStk, MaxAge, NumFish, NumSteps)
      ReDim Shakers(NumStk, MaxAge, NumFish, NumSteps)
      ReDim DropOff(NumStk, MaxAge, NumFish, NumSteps)
      ReDim MSFLandedCatch(NumStk, MaxAge, NumFish, NumSteps)
      ReDim MSFNonRetention(NumStk, MaxAge, NumFish, NumSteps)
      ReDim MSFShakers(NumStk, MaxAge, NumFish, NumSteps)
      ReDim MSFDropOff(NumStk, MaxAge, NumFish, NumSteps)
      ReDim Encounters(NumStk, MaxAge, NumFish, NumSteps)
      ReDim Cohort(NumStk, MaxAge, 4, NumSteps)
      ReDim Escape(NumStk, MaxAge, NumSteps)
      ReDim TotalLandedCatch(NumFish * 2, NumSteps)
      ReDim TotalNonRetention(NumFish, NumSteps)
      ReDim TotalEncounters(NumFish, NumSteps)
      ReDim TotalShakers(NumFish, NumSteps)
      ReDim TotalDropOff(NumFish, NumSteps)
      ReDim CohoTime4Cohort(NumStk)

      '###################################################Pete-12/17/12.

      '###################################################Pete-12/17/12.



      ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
      'ReDim NSShakerExtTotal(NumFish, NumSteps)
      'ReDim MSFShakerExtTotal(NumFish, NumSteps)
      'ReDim MSFEncountersTotal(NumFish, NumSteps)
      'ReDim NSEncountersTotal(NumFish, NumSteps)

      ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

      File_Name = FVSdatabasepath & "\FramCheck.Txt"
      If Exists(File_Name) Then Delete(File_Name)
      sw = CreateText(File_Name)
      PrnLine = "Command File =" + FVSdatabasepath + "\" & RunIDNameSelect.ToString & "     " & Date.Today.ToString
      sw.WriteLine(PrnLine)
      sw.WriteLine(" ")

      '- Check for TAMM and BackWards FRAM Selections
      If RunTAMMIter = 1 And SpeciesName = "COHO" Then Call ReadCohoTAMM()
      If RunTAMMIter = 1 And SpeciesName = "CHINOOK" Then Call ReadChinookTAMM()
      If RunBackFramFlag = 1 Then
         Call RunBackFRAM()
         sw.Flush()
         sw.Close()
         Exit Sub
      End If

      '------------------------------------------------------------
      '----- MAIN FRAM Processing Loop

      FVS_RunModel.RunProgressLabel.Visible = True
      FVS_RunModel.Cursor = Cursors.WaitCursor
        ReDim MSFLegalEncounters(NumStk, MaxAge, NumFish, NumSteps)
        ReDim MSFTotalEncounters(NumFish, NumSteps)
      '- Calculate Starting Cohort Size using Scalers and Base Period Cohorts
      Call ScaleCohort()

        Dim myPoint As Point = FVS_RunModel.RunProgressLabel.Location



        'Dim testfile As String
        'testfile = "C:\data\FRAM\SizeLimits\Testfile.txt"
        'FileOpen(77, testfile, OpenMode.Output)


        For TStep = 1 To NumSteps

            
            '- Label Update
            FVS_RunModel.Refresh()
            FVS_RunModel.RunProgressLabel.Text = " Time Step - " & TStep.ToString & " "
            myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
            FVS_RunModel.RunProgressLabel.Location = myPoint
            FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
            FVS_RunModel.RunProgressLabel.Refresh()

            Call NatMort()

            

            Call CompCatch(PTerm)

            Call IncMort(PTerm)

            Call Mature()
            If TStep = 4 Then
                TStep = 4
            End If

            Call CompCatch(Term)

            Call IncMort(Term)

            Call CompEscape()



            '- Put Cohort Numbers into Next Time Step
            Dim Jim1, Jim2 As Double
            For Stk As Integer = 1 To NumStk
                For Age As Integer = MinAge To MaxAge
                    If TStep = 3 And SpeciesName = "CHINOOK" Then
                        If Age < 5 Then
                            Cohort(Stk, Age + 1, 0, 4) = Cohort(Stk, Age, 0, 3)
                        End If
                    Else
                        If TStep < NumSteps Then
                            Cohort(Stk, Age, 0, TStep + 1) = Cohort(Stk, Age, 0, TStep)
                            Jim1 = Cohort(Stk, Age, 0, TStep + 1)
                            Jim2 = Cohort(Stk, Age, 0, TStep)
                        End If
                    End If
                Next
                '---- Create Age 2 Cohort for Chinook Time 4 Using Original Scaler
                If TStep = 3 And SpeciesName = "CHINOOK" Then
                    Cohort(Stk, 2, 0, 4) = BaseCohortSize(Stk, 2) * StockRecruit(Stk, 2, 1)

                    '-Pete Feb 2014----Code for recycling Age 3 and 4 fish in TS 4 for stocks lacking age 2s (mostly Col R springs) 
                    '                  or age 3s (Col R summers) in the model (some years)
                    '                  Note: The seemingly sloppy extra conditions are merely there to prevent recycling ages that are
                    '                  actually the end of days for a particular stock or single brood production cases (i.e., there aren't 3s behind 4s, etc.)

                    If T4CohortFlag = False Then ' allows old style processing to recreate old runs
                        If (StockRecruit(Stk, 2, 1) = 0 And StockRecruit(Stk, 3, 1) > 0 And StockRecruit(Stk, 4, 1) > 0 And StockRecruit(Stk, 5, 1) > 0) Then
                            Cohort(Stk, 3, 0, 4) = BaseCohortSize(Stk, 3) * StockRecruit(Stk, 3, 1)
                        End If

                        If (StockRecruit(Stk, 3, 1) = 0 And StockRecruit(Stk, 4, 1) > 0 And StockRecruit(Stk, 5, 1) > 0) Then
                            Cohort(Stk, 4, 0, 4) = BaseCohortSize(Stk, 4) * StockRecruit(Stk, 4, 1)
                        End If
                    End If
                    '-Pete Feb 2014----Code for recycling Age 3 fish in TS 4 for stocks lacking age 2s in the ocean

                End If
            Next

        Next TStep

        'FileClose(77)
      Dim TotalSum As Double
      '-Save Calculation Estimates to Database Table except if TAMM Run
        If RunTAMMIter = 0 Then
            For Fish As Integer = 1 To NumFish
                For TStep As Integer = 1 To NumSteps



                    '- Retention Fishery Scaler
                    If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Then
                        TotalSum = 0
                        For Stk As Integer = 1 To NumStk
                            For Age As Integer = MinAge To MaxAge
                                TotalSum += LandedCatch(Stk, Age, Fish, TStep)
                            Next
                        Next
                        If TotalSum > 0 Then
                            FisheryQuota(Fish, TStep) = CDbl(TotalSum / ModelStockProportion(Fish))
                        Else
                            FisheryQuota(Fish, TStep) = 0
                        End If
                    End If
                    '- Retention Quota
                    'If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                    '   If OptionReplaceQuota = True Then
                    '      If FisheryFlag(Fish, TStep) = 2 Then
                    '         FisheryFlag(Fish, TStep) = 1
                    '      ElseIf FisheryFlag(Fish, TStep) = 27 Then
                    '         FisheryFlag(Fish, TStep) = 17
                    '      ElseIf FisheryFlag(Fish, TStep) = 28 Then
                    '         FisheryFlag(Fish, TStep) = 18
                    '      End If
                    '   End If
                    'End If
                    '- MSF Scaler

                    If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Then
                        TotalSum = 0
                        For Stk As Integer = 1 To NumStk
                            For Age As Integer = MinAge To MaxAge
                                TotalSum += MSFLandedCatch(Stk, Age, Fish, TStep)
                            Next
                        Next
                        If TotalSum > 0 Then
                            MSFFisheryQuota(Fish, TStep) = CDbl(TotalSum / ModelStockProportion(Fish))
                        Else
                            MSFFisheryQuota(Fish, TStep) = 0
                        End If
                    End If
                    '- Retention Quota
                    'If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                    '   If OptionReplaceQuota = True Then
                    '      If FisheryFlag(Fish, TStep) = 8 Then
                    '         FisheryFlag(Fish, TStep) = 1
                    '      ElseIf FisheryFlag(Fish, TStep) = 18 Then
                    '         FisheryFlag(Fish, TStep) = 17
                    '      ElseIf FisheryFlag(Fish, TStep) = 28 Then
                    '         FisheryFlag(Fish, TStep) = 17
                    '      End If
                    '   End If
                    'End If
                Next
            Next
            '---------------------------------------------------------------------------------
            'tag111
            If FinalUpdatePass = True Or UpdateRunEncounterRateAdjustment = False Then
                Call SaveDat()
            End If
        End If

        '--- Call TAMM Procedures
        If RunTAMMIter = 1 Then
            FVS_RunModel.RunProgressLabel.Text = " TAMM Iterations "
            FVS_RunModel.RunProgressLabel.Refresh()

            PrnLine = "Tamm Input File =" & TAMMSpreadSheet
            sw.WriteLine(PrnLine)
            sw.WriteLine(" ")

            If SpeciesName = "COHO" Then
                Call TammCohoProc()
            ElseIf SpeciesName = "CHINOOK" Then
                Call TammChinookProc()
                If TammChinookConverge = 1 Then
                    MsgBox("Chinook TAMM Did Not Converge!" & vbCrLf & "TAMM Transfer Files not Created", MsgBoxStyle.OkOnly)
                Else
                    For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 7 Then
                                FisheryQuota(Fish, TStep) = CDbl(TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish))
                            ElseIf FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 8 Then
                                'If OptionReplaceQuota = True Then
                                '   If FisheryFlag(Fish, TStep) = 2 Then
                                '      FisheryFlag(Fish, TStep) = 1
                                '   Else
                                '      FisheryFlag(Fish, TStep) = 7
                                '   End If
                                'End If
                            End If
                        Next
                    Next
                    'tag111
                    If FinalUpdatePass = True Or UpdateRunEncounterRateAdjustment = False Then
                        Call SaveDat()
                    End If
                    End If
            End If
        End If

        '- Check for Negative TAMM Escapements
        If AnyNegativeEscapement = 1 Then
            PrnLine = "Negative Escapements"
            sw.WriteLine(PrnLine)
            For Stk As Integer = 1 To NumStk
                For Age As Integer = MinAge To MaxAge
                    For TStep As Integer = 1 To NumSteps
                        If Escape(Stk, Age, TStep) < 0 Then
                            PrnLine = "   Stock=" & StockName(Stk).ToString & " Age=" & Age.ToString & " = " & Escape(Stk, Age, TStep).ToString("#####0.0")
                            sw.WriteLine(PrnLine)
                        End If
                    Next TStep
                Next Age
            Next Stk
        End If

        '###################################################Pete-12/17/12.
        If SpeciesName = "COHO" And MSFBiasFlag = True Then
            Dim ProbStkList As String
            Dim msgFlag As Boolean

            For TStep As Integer = 1 To NumSteps
                For Stk As Integer = 1 To NumStk
                    If ERgtrOne(TStep, Stk) = True Then
                        msgFlag = True
                        ProbStkList += "Time step " & TStep.ToString & " " & StockTitle(Stk).ToString & vbCrLf
                    End If
                Next Stk
            Next TStep

            If msgFlag = True Then

                'PrnLine = "The ER Exceeds 100% for the following stocks & time steps:" & vbCrLf & vbCrLf & _
                '       ProbStkList & vbCrLf & _
                '       "Bias-corrected MSF calculations are invalid." & vbCrLf & _
                '       "Modify fishery inputs & re-run as necessary."
                'sw.WriteLine(PrnLine)
                MessageBox.Show("The ER Exceeded 100% for the following stocks & time steps:" & vbCrLf & vbCrLf & _
                       ProbStkList & vbCrLf & _
                       "Bias-corrected MSF calculations may be invalid." & vbCrLf & _
                       "Modify fishery inputs & re-run as necessary.")
            End If
        End If
        '###################################################Pete-12/17/12.

        '- Close FramCheck.Txt file used for Comments & Errors
        sw.Flush()
        sw.Close()

        'If AnyNegativeEscapement = 1 Then
        '   MsgBox("Negative Escapements were Detected for this Run" & vbCrLf & "Please check 'FramChk' file for Details")
        'End If

        If OptionChinookBYAEQ = 1 Then
            FVS_RunModel.RunProgressLabel.Text = " Brood Year Report "
            FVS_RunModel.RunProgressLabel.Refresh()
            Call BYERReport()
        End If

        FVS_RunModel.Cursor = Cursors.Default

        If KeepIter = True Then
            Call RunCalcs()
        End If

        'If AnyNegativeEscapement = 1 Then
        '    MsgBox("You have negative escapements. Please check the PopStat report!")
        'End If
        'AnyNegativeEscapement = 0


    End Sub

   Sub ReadCohoTAMM()

      Dim I, J, K As Integer
      ReDim SaveTermFlag(86, 1)
        ReDim SaveTermQuota(86, 1)
        ReDim SaveCoastalQuota(NumFish, NumSteps)

      TammIteration = 0

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      '- Test if TAMM Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(TAMMSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(TAMMSpreadSheet)
        xlApp.WindowState = Excel.XlWindowState.xlMinimized
        xlApp.Application.Interactive = False
SkipWBOpen:
      xlWorkSheet = xlWorkBook.Sheets("Tami")
        xlApp.Application.DisplayAlerts = False
        xlApp.Application.Interactive = False

      '- Save Original FisheryQuota and FisheryFlag Values before iteration to restore when done
      For Fish As Integer = 80 To 166
         For TStep As Integer = 4 To 5
            SaveTermFlag(Fish - 80, TStep - 4) = FisheryFlag(Fish, TStep)
            SaveTermQuota(Fish - 80, TStep - 4) = FisheryQuota(Fish, TStep)
         Next
      Next

      '**************************************************************************
      '- Read Tamm Fishery Controls from TAMI worksheet
      '- Dims 80-166 match current Coho Fram Puget Sound Fishery Numbers

      ReDim CohoTammRate(5, 166)
      ReDim CohoTammFlag(5, 166)
      ReDim CohoTammFish(5, 166)
      Dim crate As Double
      Dim cflag As Integer
      Dim AnyTerminalControl As Integer
      Dim jims As String
      AnyTerminalControl = 0
      For I = 7 To 93
         For K = 4 To 5
            J = xlWorkSheet.Cells(I, 1).Value
            If K = 4 Then
               If IsNumeric(xlWorkSheet.Cells(I, 3).Value) And xlWorkSheet.Cells(I, 3).Value > -0.001 And xlWorkSheet.Cells(I, 3).Value < 99999 Then
                  CohoTammRate(4, J) = xlWorkSheet.Cells(I, 3).Value
                  If IsNumeric(xlWorkSheet.Cells(I, 4).Value) And xlWorkSheet.Cells(I, 4).Value > 0 And xlWorkSheet.Cells(I, 4).Value < 5 Then
                     CohoTammFlag(4, J) = xlWorkSheet.Cells(I, 4).Value
                  Else
                     CohoTammRate(4, J) = 0
                     CohoTammFlag(4, J) = -99
                     If xlWorkSheet.Cells(I, 4).Value <> -99 Then
                        AnyTerminalControl = 1
                     End If
                  End If
               ElseIf xlWorkSheet.Cells(I, 3).Value() = -99 Then
                  CohoTammRate(4, J) = 0
                  CohoTammFlag(4, J) = -99
               Else
                  CohoTammRate(4, J) = 0
                  CohoTammFlag(4, J) = -99
                  jims = xlWorkSheet.Cells(I, 3).value
                  If xlWorkSheet.Cells(I, 3).value = Nothing Or xlWorkSheet.Cells(I, 3).ToString = "" Then GoTo NextTammStep
                  AnyTerminalControl = 1
               End If
               'CohoTammRate(4, J) = xlWorkSheet.Cells(I, 3).Value
               'CohoTammFlag(4, J) = xlWorkSheet.Cells(I, 4).Value
               'crate = xlWorkSheet.Cells(I, 3).Value
               'cflag = xlWorkSheet.Cells(I, 4).Value
            Else
               If J = 101 Then
                  crate = xlWorkSheet.Cells(I, 5).Value
                  cflag = xlWorkSheet.Cells(I, 6).Value
                  Jim = 1
               End If
               If IsNumeric(xlWorkSheet.Cells(I, 5).Value) And xlWorkSheet.Cells(I, 5).Value > -0.001 And xlWorkSheet.Cells(I, 5).Value < 99999 Then
                  CohoTammRate(5, J) = xlWorkSheet.Cells(I, 5).Value
                  If IsNumeric(xlWorkSheet.Cells(I, 6).Value) And xlWorkSheet.Cells(I, 6).Value > 0 And xlWorkSheet.Cells(I, 6).Value < 5 Then
                     CohoTammFlag(5, J) = xlWorkSheet.Cells(I, 6).Value
                  Else
                     CohoTammRate(5, J) = 0
                     CohoTammFlag(5, J) = -99
                     If xlWorkSheet.Cells(I, 6).Value <> -99 Then
                        AnyTerminalControl = 1
                     End If
                  End If
               ElseIf xlWorkSheet.Cells(I, 5).Value() = -99 Then
                  CohoTammRate(5, J) = 0
                  CohoTammFlag(5, J) = -99
               Else
                  CohoTammRate(5, J) = 0
                  CohoTammFlag(5, J) = -99
                  If xlWorkSheet.Cells(I, 5).value = Nothing Or xlWorkSheet.Cells(I, 5).ToString = "" Then GoTo NextTammStep
                  AnyTerminalControl = 1
               End If
               'CohoTammRate(5, J) = xlWorkSheet.Cells(I, 5).Value
               'CohoTammFlag(5, J) = xlWorkSheet.Cells(I, 6).Value
               crate = xlWorkSheet.Cells(I, 5).Value
               cflag = xlWorkSheet.Cells(I, 6).Value
            End If
            '- Set Fram Fishery Controls based on Tamm Input Values
            '- If Flag=-99 use Fram Control
            If CohoTammFlag(K, J) > 2 Then '- TAA and ETRS type controls
               'AnyTerminalControl = 1
               FisheryScaler(J, K) = 0.33 '- 1st pass 1/3 of base
               FisheryFlag(J, K) = 1
               If CohoTammFlag(K, J) = 3 Then
                  CohoTammFish(K, J) = xlWorkSheet.Cells(I, 7).Value
               Else
                  CohoTammFish(K, J) = xlWorkSheet.Cells(I, 8).Value
               End If
               '- TAMM Input Error ... User chose Wrong Control Type
               If CohoTammFish(K, J) = 0 Then
                  MsgBox("TAMM Error" & vbCrLf & "Tamm Fishery Number = Zero" & vbCrLf & "Cannot Use this Control on this Terminal Fishery!", vbOKOnly, "Tamm Input Error")
                  Exit Sub
               End If
            ElseIf CohoTammFlag(K, J) = 1 Then
               '- User Set Fishery Control in TAMM = Quota
               If FisheryQuota(J, K) <> CohoTammRate(K, J) Then
                  FisheryQuota(J, K) = CohoTammRate(K, J)
                  ChangeFishScalers = True
               End If
               FisheryFlag(J, K) = 2
               CohoTammFish(K, J) = 0
            ElseIf CohoTammFlag(K, J) = 2 Then
               '- User Set Fishery Control in TAMM = Effort Scaler
               If FisheryScaler(J, K) <> CohoTammRate(K, J) Then
                  FisheryScaler(J, K) = CohoTammRate(K, J)
                  ChangeFishScalers = True
               End If
               FisheryFlag(J, K) = 1
               CohoTammFish(K, J) = 0
            End If
NextTammStep:
         Next K
      Next I

        If AnyTerminalControl = 1 Then
            MsgBox("Error was detected in the TAMI Controls" & vbCrLf & "Please Re-Check after run", MsgBoxStyle.OkOnly)
        End If
        '******************************************************************************************************************
        ' Read in Coastal Catches for Coastal Iterations 
        xlWorkSheet = xlWorkBook.Sheets("WACoastTerminal")
        xlApp.Application.DisplayAlerts = False

        '- Save Original FisheryQuota and FisheryFlag Values before iteration to restore when done


        For TStep As Integer = 4 To 5
            I = 151
            For Fish As Integer = 45 To 74
                If Fish < 53 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish > 53 And Fish < 57 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish > 61 And Fish < 64 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish = 65 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish = 68 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish > 69 And Fish < 72 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish = 73 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                ElseIf Fish = 74 Then
                    Try
                        SaveCoastalQuota(Fish, TStep) = xlWorkSheet.Cells(I, 13 + TStep).Value
                    Catch ex As Exception
                        SaveCoastalQuota(Fish, TStep) = 0
                    End Try
                    I = I + 2
                Else
                    SaveCoastalQuota(Fish, TStep) = 0
                End If
            Next
        Next



        '- Leave Excel Open until Processing Complete .. xlWorkSheet. Updated
        'xlApp.Application.DisplayAlerts = False

   End Sub

   Sub ReadChinookTAMM()

      '- CHINOOK TAMM Variables
      Dim FoundFish, Area As Integer
      ReDim TammCatch(16, 4)
      ReDim TammEscape(7, 4)
      ReDim TammEstimate(16, 4)
      ReDim TammTermRun(7)
      ReDim TammPSER(24, 4)
      ReDim TammScaler(12, 4)

      '- Array to Link Tamm fishery Numbers to Fram Fishery Numbers
      Dim FramTammFish(71) As Integer
      FramTammFish(39) = 1 '--- Nooksack NT Net
      FramTammFish(40) = 2 '--- Nooksack TR Net
      FramTammFish(46) = 3 '--- Skagit NT Net
      FramTammFish(47) = 4 '--- Skagit TR Net
      FramTammFish(49) = 5 '--- StSno NT Net
      FramTammFish(50) = 6 '--- StSno TR Net
      FramTammFish(51) = 7 '--- Tulalip NT Net
      FramTammFish(52) = 8 '--- Tulalip TR Net
      FramTammFish(58) = 11 '--- Area 10/11 NT Net
      FramTammFish(59) = 12 '--- Area 10/11 TR Net
      FramTammFish(60) = 13 '--- Area 10A NT Net
      FramTammFish(61) = 14 '--- Area 10A TR Net
      FramTammFish(62) = 15 '--- Area 10E Sport
      FramTammFish(63) = 16 '--- Area 10E TR Net
      FramTammFish(65) = 9  '--- Hood Canal NT Net
      FramTammFish(66) = 10 '--- Hood Canal TR Net
      FramTammFish(68) = 17 '--- SPS NT Net
      FramTammFish(69) = 18 '--- SPS TR Net
      FramTammFish(70) = 19 '--- Area 13A NT Net
      FramTammFish(71) = 20 '--- Area 13A TR Net

      TammIteration = 0

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      '- Test if TAMM Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(TAMMSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(TAMMSpreadSheet)
SkipWBOpen:
      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = True
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      xlWorkSheet = xlWorkBook.Sheets("Tami")

      '**************************************************************************
      '- Read Tamm Fishery Controls from TAMI worksheet

      xlApp.Application.DisplayAlerts = False
      For Area = 1 To 24
         For TStep = 2 To 4
            If IsNumeric(xlWorkSheet.Cells(Area + 7, TStep + 1).Value) Then
               TammPSER(Area, TStep) = xlWorkSheet.Cells(Area + 7, TStep + 1).Value
            Else
               TammPSER(Area, TStep) = 0
            End If
         Next TStep
         Select Case Area
            Case 1
               TNkFWSpt! = xlWorkSheet.Cells(Area + 7, 6).Value
               TNkMSA! = xlWorkSheet.Cells(Area + 7, 7).Value
            Case 3
               TSkFWSpt! = xlWorkSheet.Cells(Area + 7, 6).Value
               TSkMSA! = xlWorkSheet.Cells(Area + 7, 7).Value
            Case 5
               TSnFWSpt! = xlWorkSheet.Cells(Area + 7, 6).Value
               TSnMSA! = xlWorkSheet.Cells(Area + 7, 7).Value
            Case 9
               THCFWSpt! = xlWorkSheet.Cells(Area + 7, 6).Value
         End Select
      Next Area
      SpsYrSpl = xlWorkSheet.Cells(Area + 7, 3).Value

      '- Leave Excel Open until Processing Complete .. xlWorkSheet. Updated
      xlApp.Application.DisplayAlerts = False

      ''- Save WorkBook and Close Application if Necessary
      'xlApp.Application.DisplayAlerts = False
      'xlWorkBook.Save()
      'If WorkBookWasNotOpen = True Then
      '   xlWorkBook.Close()
      'End If
      'If ExcelWasNotRunning = True Then
      '   xlApp.Application.Quit()
      '   xlApp.Quit()
      'Else
      '   xlApp.Visible = True
      '   xlApp.WindowState = Excel.XlWindowState.xlMinimized
      'End If
      xlApp.Visible = True
        xlApp.Application.DisplayAlerts = True

      'xlApp = Nothing

      '--- Chinook Quota and Effort Data is replaced by TAMM Input Variables when
      '--- TAMM Processing is Requested.  The Time Period 2 and 4 Values are input
      '--- as Target (i.e. Quota) values and should be computed directly

      If TammChinookRunFlag > 1 Then GoTo SkipTami2

      For TStep As Integer = 2 To 4
            For Fish As Integer = 39 To 71  '--- Chinook Specific Terminal Fishery #'s
                If Fish = 61 Then
                    Fish = 61
                End If
                If (TStep = 2 Or TStep = 4) Then
                    If FramTammFish(Fish) <> 0 Then
                        If TammPSER(FramTammFish(Fish), TStep) = -88 Then
                            GoTo SkipFlag88
                        End If
                    End If
                    If (Fish = 46 Or Fish = 47) Then '--- Skagit Net TAMM Targets
                        If (TStep = 4) Then
                            FisheryFlag(Fish, 4) = 2       '--- Target Flag
                            If Fish = 46 Then
                                FisheryQuota(Fish, 4) = TammPSER(3, 4) '--- Skagit NT Net
                            Else
                                FisheryQuota(Fish, 4) = TammPSER(4, 4) '--- Skagit TR Net
                            End If
                        End If
                        If (TStep = 2) Then   '--- Time 2 either Rate or Target
                            If TammPSER(3, 2) > 1.0 Then
                                FisheryFlag(46, 2) = 2  '--- Target Flag Time 2
                                FisheryQuota(46, 2) = TammPSER(3, 2) '--- Skagit NT Net
                            End If
                            If TammPSER(4, 2) > 1.0 Then
                                FisheryFlag(47, 2) = 2    '--- Target Flag Time 2
                                FisheryQuota(47, 2) = TammPSER(4, 2)    '--- Skagit TR Net
                            End If
                        End If
                    Else       '--- Non-Skagit Target Fisheries Times 2 & 4
                        FoundFish = 1
                        Select Case Fish
                            Case 39
                                FisheryScaler(Fish, TStep) = TammPSER(1, TStep) '--- Nook. NT Net
                            Case 40
                                FisheryScaler(Fish, TStep) = TammPSER(2, TStep) '--- Nook. TR Net
                            Case 49
                                FisheryScaler(Fish, TStep) = TammPSER(5, TStep) '--- StSno NT Net
                            Case 50
                                FisheryScaler(Fish, TStep) = TammPSER(6, TStep) '--- StSno TR Net
                            Case 51
                                FisheryScaler(Fish, TStep) = TammPSER(7, TStep) '--- Tula. NT Net
                            Case 52
                                FisheryScaler(Fish, TStep) = TammPSER(8, TStep) '--- Tula. TR Net
                            Case 58
                                FisheryScaler(Fish, TStep) = TammPSER(11, TStep) '--- 10-11 NT Net
                            Case 59
                                FisheryScaler(Fish, TStep) = TammPSER(12, TStep) '--- 10-11 TR Net
                            Case 60
                                FisheryScaler(Fish, TStep) = TammPSER(13, TStep) '--- 10A NT NEt
                            Case 61
                                FisheryScaler(Fish, TStep) = TammPSER(14, TStep) '--- 10A TR Net
                            Case 62
                                FisheryScaler(Fish, TStep) = TammPSER(15, TStep) '--- 10E Sport
                            Case 63
                                FisheryScaler(Fish, TStep) = TammPSER(16, TStep) '--- 10E TR Net
                            Case 65
                                FisheryScaler(Fish, TStep) = TammPSER(9, TStep)  '--- HC NT Net
                            Case 66
                                FisheryScaler(Fish, TStep) = TammPSER(10, TStep) '--- HC TR Net
                            Case 68
                                FisheryScaler(Fish, TStep) = TammPSER(17, TStep) '--- SPS NT Net
                            Case 69
                                FisheryScaler(Fish, TStep) = TammPSER(18, TStep) '--- SPS TR Net
                            Case 70
                                FisheryScaler(Fish, TStep) = TammPSER(19, TStep) '--- 13A NT Net
                            Case 71
                                FisheryScaler(Fish, TStep) = TammPSER(20, TStep) '--- 13A TR Net
                            Case Else
                                FoundFish = 0
                        End Select
                        If FoundFish = 1 Then
                            FisheryFlag(Fish, TStep) = 2  '--- Target Quota Flag
                            FisheryQuota(Fish, TStep) = FisheryScaler(Fish, TStep)
                        End If
                    End If
                Else     '---- SPS Net Fisheries Time 3 are Quotas for now
                    If FramTammFish(Fish) <> 0 Then
                        If TammPSER(FramTammFish(Fish), TStep) = -88 Then
                            GoTo SkipFlag88
                        End If
                    End If
                    If TStep = 3 Then
                        FoundFish = 1
                        Select Case Fish
                            '--- Added Rate/Quota Values for Time-3 Non-SPS Fisheries
                            '--- because of TAMM iteration problems (Negative Esc's)
                            Case 39
                                FisheryScaler(Fish, TStep) = TammPSER(1, TStep) '--- Nook. NT Net
                                'If FisheryScaler(Fish, TStep) < 10.0 Then FoundFish = 2 'AHB 08/17/2016
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2 'AHB 08/17/2016
                            Case 40
                                'If FisheryScaler(Fish, TStep) < 10.0 Then FoundFish = 2 'AHB 08/17/2016
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2 'AHB 08/17/2016
                            Case 49
                                'If FisheryScaler(Fish, TStep) < 10.0 Then FoundFish = 2 'AHB 08/17/2016
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2 'AHB 08/17/2016
                            Case 50
                                'If FisheryScaler(Fish, TStep) < 10.0 Then FoundFish = 2 'AHB 08/17/2016
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2 'AHB 08/17/2016
                            Case 51
                                FisheryScaler(Fish, TStep) = TammPSER(7, TStep) '--- Tula. NT Net
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2
                            Case 52
                                FisheryScaler(Fish, TStep) = TammPSER(8, TStep) '--- Tula. TR Net
                                If FisheryScaler(Fish, TStep) < 1.0 Then FoundFish = 2
                                '--- SPS Fisheries Below
                            Case 58
                                FisheryScaler(Fish, TStep) = TammPSER(11, TStep) '--- 10-11 NT Net
                            Case 59
                                FisheryScaler(Fish, TStep) = TammPSER(12, TStep) '--- 10-11 TR Net
                            Case 60
                                FisheryScaler(Fish, TStep) = TammPSER(13, TStep) '--- 10A NT NEt
                            Case 61
                                FisheryScaler(Fish, TStep) = TammPSER(14, TStep) '--- 10A TR Net
                            Case 62
                                FisheryScaler(Fish, TStep) = TammPSER(15, TStep) '--- 10E Sport
                            Case 63
                                FisheryScaler(Fish, TStep) = TammPSER(16, TStep) '--- 10E TR Net
                            Case 68
                                FisheryScaler(Fish, TStep) = TammPSER(17, TStep) '--- SPS NT Net
                            Case 69
                                FisheryScaler(Fish, TStep) = TammPSER(18, TStep) '--- SPS TR Net
                            Case 70
                                FisheryScaler(Fish, TStep) = TammPSER(19, TStep) '--- 13A NT Net
                            Case 71
                                FisheryScaler(Fish, TStep) = TammPSER(20, TStep) '--- 13A TR Net
                            Case Else
                                FoundFish = 0
                        End Select
                        If FoundFish = 1 Then
                            FisheryFlag(Fish, TStep) = 2 '--- Target Quota Flag
                            FisheryQuota(Fish, TStep) = FisheryScaler(Fish, TStep)
                        End If
                        If FoundFish = 2 Then
                            FisheryFlag(Fish, TStep) = 1 '--- Effort Scalar Flag
                        End If
                    End If
                End If
SkipFlag88:
            Next Fish
      Next TStep
SkipTami2:

   End Sub

   Sub RunBackFRAM()

      '----- MAIN Backwards FRAM Processing Loop
      Call ScaleCohort()     
        For TStep = 1 To NumSteps

            
            Call NatMort()
            Call CompCatch(PTerm)
            Call IncMort(PTerm)
            
            Call Mature()
            Call CompCatch(Term)
            Call IncMort(Term)
            Call CompEscape()
            '- Put Cohort Numbers into Next Time Step
            For Stk As Integer = 1 To NumStk
                For Age As Integer = MinAge To MaxAge
                    If TStep < NumSteps Then
                        Cohort(Stk, Age, PTerm, TStep + 1) = Cohort(Stk, Age, PTerm, TStep)
                    End If
                Next
            Next
            '- Check for Negative Escapements
            If AnyNegativeEscapement = 1 Then
                PrnLine = "Negative Escapements"
                sw.WriteLine(PrnLine)
                For Stk As Integer = 1 To NumStk
                    For Age As Integer = MinAge To MaxAge
                        If Escape(Stk, Age, TStep) < 0 Then
                            PrnLine = "   Stock=" & StockName(Stk).ToString & " Age=" & Age.ToString & " = " & Escape(Stk, Age, TStep).ToString("#####0.0")
                            sw.WriteLine(PrnLine)
                        End If
                    Next Age
                Next Stk
            End If
        Next

      '- Move this Call to FVS_BackwardsFram (after all iterations)
      'Call SaveDat()

   End Sub


   Sub CompCatch(ByVal TerminalType As Integer)
        ReDim NSFQuotaTotal(NumFish, TStep), MSFQuotaTotal(NumFish, TStep)
        Dim SelBiasVersion As Integer
        Dim SecondPass As Boolean
        Dim JimD As Double

        ''<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
        ''Additional Pete Size Limit/External Shaker work to ensure stability of total encounters
        ''And a total encounters ratio that's consistent with test fishery data
        'Dim EqualizeSizeLimit As Boolean
        'Dim TempScalar, MSFTempScalar As Double
        'Dim Adjustment As Double

        'EqualizeSizeLimit = False

        ''<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>

        '**************************************************************************
        'Compute catch depending upon value of FisheryFlag parameter

        ' FisheryFlag = 0 - No Fishery
        ' FisheryFlag = 1 - Fishery Scale Factor
        ' FisheryFlag = 2 - Quota

        ' Selective Flag = 7 - Fishery Scale Factor
        ' Selective Flag = 8 - Quota

        '**************************************************************************

        'PASS 1 CATCH: COMPUTE Landed Catch with Fishery Scaler

        '- Size Limit Degugging Flag - Set to ONE for Printing to FramCheck.Txt
        SkipJim = 1
        SelBiasVersion = 0 '- MSF Bias Debugging Flag = 5 to run with this

        PrnLine = "=================== CompCatch Results ========================"
        If SkipJim = 1 Then sw.WriteLine(PrnLine)
        PrnLine = "--------Stk Age Fsh TSp LandCatch Cohort     BaseExpRt FishScl StkFshI LegalProp"
        If SkipJim = 1 Then sw.WriteLine(PrnLine)
        
        For Fish As Integer = 1 To NumFish
            
            If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextScalerFishery ' if there is no catch in the base period
            '- Fishery/Time-Step can only be Terminal or Pre-Terminal
            If TerminalFisheryFlag(Fish, TStep) = TerminalType Then

                '- Zero Totals Arrays for TAMM Iteration Calculations
                TotalLandedCatch(Fish, TStep) = 0
                TotalLandedCatch(NumFish + Fish, TStep) = 0
                TotalEncounters(Fish, TStep) = 0
                TotalDropOff(Fish, TStep) = 0
                NSFQuotaTotal(Fish, TStep) = 0
                MSFQuotaTotal(Fish, TStep) = 0

                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                'MSFEncountersTotal(Fish, TStep) = 0 ' Wipe these out...
                'NSEncountersTotal(Fish, TStep) = 0  ' Wipe these out...
                'MSFTempScalar = MSFFisheryScaler(Fish, TStep) 'A holder for replacement; required for rescaling to TF prediction
                'TempScalar = FisheryScaler(Fish, TStep) 'A holder for replacement; required for rescaling to TF prediction
                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                If SizeLimitFix = True And MinSizeLimit(Fish, TStep) < ChinookBaseSizeLimit(Fish, TStep) Then
                    SizeLimitFixLanded(Fish, TerminalType)
                Else

                    For Stk As Integer = 1 To NumStk
                        For Age As Integer = MinAge To MaxAge

                            '- Zero Calculation Arrays for TAMM Iteration Calculations
                            LandedCatch(Stk, Age, Fish, TStep) = 0
                            DropOff(Stk, Age, Fish, TStep) = 0
                            Encounters(Stk, Age, Fish, TStep) = 0
                            NonRetention(Stk, Age, Fish, TStep) = 0
                            MSFLandedCatch(Stk, Age, Fish, TStep) = 0
                            MSFDropOff(Stk, Age, Fish, TStep) = 0
                            MSFEncounters(Stk, Age, Fish, TStep) = 0
                            MSFNonRetention(Stk, Age, Fish, TStep) = 0

                            '- Compute Legal Proportion by Stock, Age, and Time-Step
                            ChinookBaseLegProp = False

                            Call CompLegProp(Stk, Age, Fish, TerminalType)

                            ''- Check if New Size Limit is different from Base Period Size Limit
                            'If SpeciesName = "CHINOOK" And MinSizeLimit(Fish, TStep) <> ChinookBaseSizeLimit(Fish, TStep) Then
                            '   ChinookBaseLegProp = True
                            '   Call CompLegProp(Stk, Age, Fish, TerminalType)
                            'End If


                            ''****************************************************************************************
                            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                            ''Only use different legal and sublegal proportions for 1) scenarios involving <22" limits,
                            ''2) Puget Sound sport fisheries, and 3) combo fisheries with different limits during NS & MSF periods



                            'If SizeLimitScenario = True Then
                            '   Select Case Fish
                            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                            '         If AltFlag(Fish, TStep) > 8 And AltLimitMSF(Fish, TStep) <> AltLimitNS(Fish, TStep) _
                            '            And AltLimitMSF(Fish, TStep) > 0 And AltLimitNS(Fish, TStep) > 0 Then
                            '            LegalProportion = NSLegalProp
                            '         End If
                            '         '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>'New code to stabilize catch. Jan 25 2013
                            '         If EqualizeSizeLimit = True Then
                            '            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Then
                            '               If LegalProportion > 0 Then
                            '                  Adjustment = (CNRLegalProp * (1 + ExternalBaseRatio(Fish, TStep))) / (LegalProportion * (1 + LSRatioNS(Fish, TStep)))
                            '                  FisheryScaler(Fish, TStep) = Adjustment * TempScalar
                            '               End If
                            '            End If
                            '         End If
                            '         '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                            '   End Select
                            'End If
                            ''############################# END NEW CODE ############################ Pete-Jan. 2013
                            ''****************************************************************************************
                            

                            '- Retention Fishery Scalers 
                            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Then

                                '--- Special Case Chinook Tulalip Bay Net HR=.99
                                ' ... Catch All Mature Fish
                                If SpeciesName = "CHINOOK" And NumStk < 50 And TStep = 3 And (Fish = 51 Or Fish = 52) And FisheryScaler(52, 3) = 0.99 And Stk = 10 Then
                                    '- Regular Chinook FRAM
                                    LandedCatch(Stk, Age, Fish, TStep) = 0
                                ElseIf SpeciesName = "CHINOOK" And NumStk > 50 And TStep = 3 And (Fish = 52) And FisheryScaler(52, 3) = 0.99 And (Stk = 19 Or Stk = 20) Then
                                    '- Selective Fishery Version
                                    LandedCatch(Stk, Age, Fish, TStep) = 0
                                Else


                                    '--- Main FRAM Harvest Algorithm --------------------
                                    If Stk = 53 And TStep = 2 And Fish = 30 And Age = 4 Then
                                        Jim = 1
                                    End If

                                    LandedCatch(Stk, Age, Fish, TStep) = _
                                       Cohort(Stk, Age, TerminalType, TStep) * _
                                       BaseExploitationRate(Stk, Age, Fish, TStep) * _
                                       FisheryScaler(Fish, TStep) * _
                                       StockFishRateScalers(Stk, Fish, TStep) * _
                                       LegalProportion

                                    JimD = LandedCatch(Stk, Age, Fish, TStep)

                                    '- DEBUG Code to Check CompCatch Calculations
                                    'If SkipJim = 1 And (Fish = 1) And TStep = 2 And LandedCatch(Stk, Age, Fish, TStep) <> 0 Then
                                    'If SkipJim = 1 And LandedCatch(Stk, Age, Fish, TStep) <> 0 Then
                                    '   PrnLine = String.Format("test-{0,3}{1,4}{2,4}{3,4}", Stk.ToString, Age.ToString, Fish.ToString, TStep.ToString)
                                    '   PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                    '   PrnLine &= String.Format("{0,13}", Cohort(Stk, Age, TerminalType, TStep).ToString("######0.0000"))
                                    '   PrnLine &= String.Format("{0,11}", BaseExploitationRate(Stk, Age, Fish, TStep).ToString("0.00000000"))
                                    '   PrnLine &= String.Format("{0,8}", FisheryScaler(Fish, TStep).ToString("0.00000"))
                                    '   PrnLine &= String.Format("{0,8}", StockFishRateScalers(Stk, Fish, TStep).ToString("##0.000"))
                                    '   PrnLine &= String.Format("{0,8}", ModelStockProportion(Fish).ToString("0.00000"))
                                    '   PrnLine &= String.Format("{0,8}", LegalProportion.ToString("##0.0000"))
                                    '   PrnLine &= String.Format("{0,11}", FisheryName(Fish).ToString)
                                    '   PrnLine &= String.Format("{0,11}", StockName(Stk).ToString)
                                    '   sw.WriteLine(PrnLine)
                                    'End If
                                    Encounters(Stk, Age, Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                                    TotalEncounters(Fish, TStep) = TotalEncounters(Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                                    ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                                    'NSEncountersTotal(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                                    ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                                End If
                            End If




                            ''****************************************************************************************
                            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                            ''Only use different legal and sublegal proportions for 1) scenarios involving <22" limits,
                            ''2) Puget Sound sport fisheries, and 3) combo fisheries with different limits during NS & MSF periods
                            'If SizeLimitScenario = True Then
                            '   Select Case Fish
                            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                            '         If AltFlag(Fish, TStep) > 8 And AltLimitMSF(Fish, TStep) <> AltLimitNS(Fish, TStep) _
                            '            And AltLimitMSF(Fish, TStep) > 0 And AltLimitNS(Fish, TStep) > 0 Then
                            '            LegalProportion = MSFLegalProp
                            '         End If
                            '         '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>'New code to stabilize catch. Jan 25 2013
                            '         If EqualizeSizeLimit = True Then
                            '            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Then
                            '               If LegalProportion > 0 Then
                            '                  Adjustment = (CNRLegalProp * (1 + ExternalBaseRatio(Fish, TStep))) / (LegalProportion * (1 + LSRatioMSF(Fish, TStep)))
                            '                  MSFFisheryScaler(Fish, TStep) = Adjustment * MSFTempScalar
                            '               End If
                            '            End If
                            '         End If
                            '         '<><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><><>
                            '   End Select
                            'End If
                            ''############################# END NEW CODE ############################ Pete-Jan. 2013
                            ''****************************************************************************************

                            '- MSF Fishery Scaler
                            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Then
                                MSFLandedCatch(Stk, Age, Fish, TStep) = _
                                   Cohort(Stk, Age, TerminalType, TStep) * _
                                   BaseExploitationRate(Stk, Age, Fish, TStep) * _
                                   MSFFisheryScaler(Fish, TStep) * _
                                   StockFishRateScalers(Stk, Fish, TStep) * _
                                   LegalProportion
                                MSFEncounters(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep)

                                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                                'MSFEncountersTotal(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

                                TotalEncounters(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                '--- Use Selective Incidental Rate on ALL fish encountered
                                MSFDropOff(Stk, Age, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * MSFLandedCatch(Stk, Age, Fish, TStep)
                                TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                                '- All Stocks in Marked/UnMarked pairs
                                If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                                    MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                                    MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                                    TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                                Else           '--- UnMarked (Wild) in Selective
                                    MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                                    MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                                    TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                                End If
                            End If




                            ''****************************************************************************************
                            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                            ''Only use different legal and sublegal proportions for 1) scenearios involving <22" limits,
                            ''2) Puget Sound sport fisheries, and 3) combo fisheries with different limits during NS & MSF periods
                            'If SizeLimitScenario = True Then
                            '   Select Case Fish
                            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                            '         If AltFlag(Fish, TStep) > 8 And AltLimitMSF(Fish, TStep) <> AltLimitNS(Fish, TStep) _
                            '            And AltLimitMSF(Fish, TStep) > 0 And AltLimitNS(Fish, TStep) > 0 Then
                            '            LegalProportion = NSLegalProp
                            '         End If
                            '   End Select
                            'End If
                            ''############################# END NEW CODE ############################ Pete-Jan. 2013
                            ''****************************************************************************************

                            '- Retention Quota Fishery
                            If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                '- First Pass for Quota Fisheries - Landed Catch as if FisheryScaler = 1

                                If TStep = 4 And Stk = 29 And Fish = 15 And Age = 4 Then
                                    Jim = 1
                                End If

                                




                                LandedCatch(Stk, Age, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                                'Encounters(Stk, Age, Fish, TStep) += Encounters(Stk, Age, Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                                'TotalEncounters(Fish, TStep) = TotalEncounters(Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                                NSFQuotaTotal(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)

                                If Double.IsNaN(LandedCatch(Stk, Age, Fish, TStep)) Then
                                    MsgBox("Invalid Landed Catch for Stk " & Stk & ", Fishery " & Fish & ", Time Step " & TStep & ".")
                                End If
                            End If






                            ''****************************************************************************************
                            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                            ''Only use different legal and sublegal proportions for 1) scenearios involving <22" limits,
                            ''2) Puget Sound sport fisheries, and 3) combo fisheries with different limits during NS & MSF periods
                            'If SizeLimitScenario = True Then
                            '   Select Case Fish
                            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                            '         If AltFlag(Fish, TStep) > 8 And AltLimitMSF(Fish, TStep) <> AltLimitNS(Fish, TStep) _
                            '            And AltLimitMSF(Fish, TStep) > 0 And AltLimitNS(Fish, TStep) > 0 Then
                            '            LegalProportion = MSFLegalProp
                            '         End If
                            '   End Select
                            'End If
                            ''############################# END NEW CODE ############################ Pete-Jan. 2013
                            ''****************************************************************************************

                            '- MSF Quota Fishery
                            If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                                '- First Pass for Quota Fisheries - Landed Catch as if FisheryScaler = 1
                                MSFLandedCatch(Stk, Age, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                                MSFEncounters(Stk, Age, Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                'TotalEncounters(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                '--- Use Selective Incidental Rate on ALL fish encountered
                                MSFDropOff(Stk, Age, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * MSFLandedCatch(Stk, Age, Fish, TStep)
                                'TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                                If (Stk Mod 2) = 0 Then
                                    '- Marked Fish in Selective Quota
                                    MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                                    MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) - (MSFNonRetention(Stk, Age, Fish, TStep) / MarkSelectiveMortRate(Fish, TStep))
                                    'TotalNonRetention(Fish, TStep) +=  MSFNonRetention(Stk, Age, Fish, TStep)
                                Else
                                    '--- UnMarked (Wild) in Selective Quota
                                    MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                                    MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                                    'TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                                End If
                                If Double.IsNaN(MSFLandedCatch(Stk, Age, Fish, TStep)) Then
                                    MsgBox("Invalid MSF Landed Catch Size for Stk " & Stk & ", Fishery " & Fish & ", Time Step " & TStep & ".")
                                End If
                                MSFQuotaTotal(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                            End If

                            '- All Stocks Landed Catch in Fisheries 1 to NumFish
                            TotalLandedCatch(Fish, TStep) = TotalLandedCatch(Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)

                            '- Only UnMarked (Wild) in Fisheries NumFish+1 to NumFish*2
                            If (Stk Mod 2) <> 0 Then
                                TotalLandedCatch(NumFish + Fish, TStep) = TotalLandedCatch(NumFish + Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                            End If

                        Next Age
                    Next Stk
                End If
            End If
NextScalerFishery:
            'Debug.Print("Fishery " & FisheryName(Fish) & "TS " & TStep & " Landed = " & TotalLandedCatch(Fish, TStep))
        Next Fish

      'Pass #2 - COMPUTE CATCH IN FISHERIES WITH QUOTAS 

        For Fish As Integer = 1 To NumFish
            
            


            If TerminalFisheryFlag(Fish, TStep) = TerminalType Then

                '- Retention Quota Fishery Pass #2
                If (FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28) Then
                    If NSFQuotaTotal(Fish, TStep) > 0 Then
                        FisheryScaler(Fish, TStep) = ModelStockProportion(Fish) * FisheryQuota(Fish, TStep) / NSFQuotaTotal(Fish, TStep)
                        For Stk As Integer = 1 To NumStk
                            For Age As Integer = MinAge To MaxAge
                                If Stk = 29 And Age = 4 And TStep = 4 And Fish = 15 Then
                                    Jim = 1
                                End If
                                '- Subtract Landed Catch from 1st Pass from Total
                                TotalLandedCatch(Fish, TStep) -= LandedCatch(Stk, Age, Fish, TStep)

                                If (Stk Mod 2) <> 0 Then      '--- UnMarked Fish
                                    TotalLandedCatch(NumFish + Fish, TStep) -= LandedCatch(Stk, Age, Fish, TStep)
                                End If
                                '- Compute new Landed Catch and add back to Total
                                LandedCatch(Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * LandedCatch(Stk, Age, Fish, TStep)
                                Encounters(Stk, Age, Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                                TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                                If (Stk Mod 2) <> 0 Then      '--- UnMarked Fish
                                    TotalLandedCatch(NumFish + Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                                End If
                                TotalEncounters(Fish, TStep) = TotalEncounters(Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                            Next Age
                        Next Stk

                        ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                        'NSEncountersTotal(Fish, TStep) = FisheryQuota(Fish, TStep) * ModelStockProportion(Fish)
                        ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

                    Else
                        FisheryScaler(Fish, TStep) = 0
                    End If

                End If

                '- MSF Quota Fishery Pass #2
                If (FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28) Then
                    If MSFQuotaTotal(Fish, TStep) > 0 Then
                        MSFFisheryScaler(Fish, TStep) = ModelStockProportion(Fish) * MSFFisheryQuota(Fish, TStep) / MSFQuotaTotal(Fish, TStep)
                        For Stk As Integer = 1 To NumStk
                            For Age As Integer = MinAge To MaxAge
                                '- Subtract Landed Catch from 1st Pass from Total
                                TotalLandedCatch(Fish, TStep) -= MSFLandedCatch(Stk, Age, Fish, TStep)
                                If (Stk Mod 2) <> 0 Then      '--- UnMarked Fish
                                    TotalLandedCatch(NumFish + Fish, TStep) -= MSFLandedCatch(Stk, Age, Fish, TStep)
                                End If
                                '- Compute new Landed Catch and add back to Total
                                MSFLandedCatch(Stk, Age, Fish, TStep) = MSFFisheryScaler(Fish, TStep) * MSFLandedCatch(Stk, Age, Fish, TStep)
                                TotalLandedCatch(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                If (Stk Mod 2) <> 0 Then      '--- UNMarked Fish
                                    TotalLandedCatch(NumFish + Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                                End If
                                '--- Adjust CNR and Shakers to Quota Scalar
                                MSFDropOff(Stk, Age, Fish, TStep) = MSFDropOff(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep)
                                TotalDropOff(Fish, TStep) += MSFDropOff(Stk, Age, Fish, TStep)
                                MSFNonRetention(Stk, Age, Fish, TStep) = MSFNonRetention(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep)
                                TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                                MSFEncounters(Stk, Age, Fish, TStep) = MSFEncounters(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep)

                                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
                                'MSFEncountersTotal(Fish, TStep) += MSFEncounters(Stk, Age, Fish, TStep)
                                ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

                                TotalEncounters(Fish, TStep) += MSFEncounters(Stk, Age, Fish, TStep)
                            Next Age
                        Next Stk
                    Else
                        MSFFisheryScaler(Fish, TStep) = 0
                    End If
                End If

            End If

        Next Fish



      Age = 3

      'For Fish = 1 To NumFish
      '   '   Fish = 140
      '   For Stk = 1 To NumStk
      '      '      'If SkipJim = 1 And LandedCatch(Stk, Age, Fish, TStep) <> 0 And Cohort(Stk, Age, TerminalType, TStep) <> 0 Then
      '      If LandedCatch(Stk, Age, Fish, TStep) <> 0 And Cohort(Stk, Age, TerminalType, TStep) <> 0 Then
      '         PrnLine = String.Format("test-{0,3}{1,4}{2,4}{3,4}", Stk.ToString, Age.ToString, Fish.ToString, TStep.ToString)
      '         PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
      '         PrnLine &= String.Format("{0,13}", Cohort(Stk, Age, TerminalType, TStep).ToString("######0.0000"))
      '         PrnLine &= String.Format("{0,11}", BaseExploitationRate(Stk, Age, Fish, TStep).ToString("0.00000000"))
      '         PrnLine &= String.Format("{0,8}", FisheryScaler(Fish, TStep).ToString("0.00000"))
      '         PrnLine &= String.Format("{0,8}", StockFishRateScalers(Stk, Fish, TStep).ToString("##0.000"))
      '         PrnLine &= String.Format("{0,8}", ModelStockProportion(Fish).ToString("0.00000"))
      '         PrnLine &= String.Format("{0,8}", LegalProportion.ToString("##0.0000"))
      '         PrnLine &= String.Format("{0,11}", FisheryName(Fish).ToString)
      '         PrnLine &= String.Format("{0,11}", StockName(Stk).ToString)
      '         sw.WriteLine(PrnLine)
      '      End If
      '   Next
      'Next
        'End If


        If SpeciesName = "COHO" And MSFBiasFlag = True Then Call MSFBiasCorrectionCalcs(TerminalType)



        '--- COMPUTE TERMINAL AREA NET CATCHES FOR CHINOOK TAMM ---

        If (RunTAMMIter = 1 And SpeciesName = "CHINOOK") Then
            '--- GET TOTAL CATCH OF ALL STOCKS IN AREAS 7BCD, 8, 8A, 8D, 12-12D
            TammCatch(1, TStep) = TotalLandedCatch(39, TStep)
            TammCatch(2, TStep) = TotalLandedCatch(40, TStep)

            PrnLine = String.Format("7BCNT-TSTotal {0,1} {1,8}", TStep, TotalLandedCatch(39, TStep.ToString("#####.00")))
            sw.WriteLine(PrnLine)
            PrnLine = String.Format("7BCTR-TSTotal {0,1} {1,8}", TStep, TotalLandedCatch(40, TStep.ToString("#####.00")))
            sw.WriteLine(PrnLine)

            TammCatch(3, TStep) = TotalLandedCatch(46, TStep)
            TammCatch(4, TStep) = TotalLandedCatch(47, TStep)
            TammCatch(5, TStep) = TotalLandedCatch(49, TStep)
            TammCatch(6, TStep) = TotalLandedCatch(50, TStep)
            TammCatch(7, TStep) = TotalLandedCatch(51, TStep)
            TammCatch(8, TStep) = TotalLandedCatch(52, TStep)
            TammCatch(9, TStep) = TotalLandedCatch(65, TStep)
            TammCatch(10, TStep) = TotalLandedCatch(66, TStep)
            TammCatch(11, TStep) = TotalLandedCatch(70, TStep)
            TammCatch(12, TStep) = TotalLandedCatch(71, TStep)

            '--- AGE 2-5 CATCH OF NOOKSACK SPRING IN BELLINGHAM BAY - 7BCD
            TammCatch(13, TStep) = 0.0
            TammCatch(14, TStep) = 0.0
         For Stk As Integer = 2 To 3
            For Age As Integer = MinAge To 5
               If NumStk < 50 Then
                  TammCatch(13, TStep) = TammCatch(13, TStep) + LandedCatch(Stk, Age, 39, TStep)
                  TammCatch(14, TStep) = TammCatch(14, TStep) + LandedCatch(Stk, Age, 40, TStep)
               Else
                  TammCatch(13, TStep) = TammCatch(13, TStep) + LandedCatch(Stk * 2 - 1, Age, 39, TStep) + LandedCatch(Stk * 2, Age, 39, TStep)
                  TammCatch(14, TStep) = TammCatch(14, TStep) + LandedCatch(Stk * 2 - 1, Age, 40, TStep) + LandedCatch(Stk * 2, Age, 40, TStep)
               End If
            Next Age
         Next Stk

        End If

    End Sub


   Sub MSFBiasCorrectionCalcs(ByVal TerminalType As Integer)


      '======================================================================================
      ' Calculate the MSF Bias (Change in Ratio) by Stock for All fisheries in the Time Step
      ' Calculate the Weighted RMR Release Mortality Rate Proportionally by the Marked ER in MSFs
      ' then Proportionally Divided Bias by Stock/Fishery
      ' using Ratio of Selective Fishery Marked ER's

      Dim NSStkERRate(NumStk, NumFish + 1) As Double, NSStkERRateTilde(NumStk, NumFish + 1) As Double
      Dim MSFStkERRate(NumStk, NumFish + 1) As Double, MSFStkERRateTilde(NumStk, NumFish + 1) As Double
      Dim StkERRate(NumStk) As Double, StkERRateTilde(NumStk) As Double

      Dim FishERRate As Double, SelWeightRMR As Double
      Dim BiasCorrectedER(NumStk) As Double, UnMarkedUnBiasedMortality(Stk) As Double
      Dim CorrectedBiasRatio As Double
      Dim SecondPass, MSFBiasIter As Boolean
      Dim MSFtolerance, MSFTestTolerance As Double
      Dim MSFBiasCount As Integer

      Dim Alpha(NumStk) As Double
      Dim Meeew(NumStk) As Double
      Dim StkMort(NumStk) As Double

      Dim NSFishWeight(NumStk, NumFish) As Double, MSFishWeight(NumStk, NumFish) As Double
      Dim NSFishMort(NumStk, NumFish) As Double, MSFishMort(NumStk, NumFish) As Double
      Dim TotalNSLanded(NumFish) As Double, TotalMSFLanded(NumFish) As Double

      Dim EncounterMort(NumStk, NumFish) As Double
      Dim PPNLandedCat(NumStk, NumFish) As Double
      Dim PPNNonRetion(NumStk, NumFish) As Double
      Dim PPNIncidental(NumStk, NumFish) As Double

      ''- DEBUG Code .. Header
      'PrnLine = String.Format("{0,8}", "Num Fishery TStep Num Stock-- EncMort PPNLand PPNNonR PPNIncd")
      'sw.WriteLine(PrnLine)

        ReDim ERgtrOne(NumSteps, NumStk)

        If TammIteration = 1 And TStep = 5 And TerminalType = 1 Then
            Jim = 1
        End If


      '- Compute biased proportion of encounters that die, proportion landed, release morts, and incidental morts in MSF

      For Fish As Integer = 1 To NumFish
         If TerminalFisheryFlag(Fish, TStep) <> TerminalType Then GoTo NextBiasRateFish
         'If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
         For Stk As Integer = 1 To NumStk
            If Stk Mod 2 = 0 Then
               EncounterMort(Stk, Fish) = ((1 - MarkSelectiveMarkMisID(Fish, TStep)) + _
                   MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep)) / (1 + MarkSelectiveIncRate(Fish, TStep))
               PPNLandedCat(Stk, Fish) = (1 - MarkSelectiveMarkMisID(Fish, TStep)) / _
                  (1 - MarkSelectiveMarkMisID(Fish, TStep) + MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
               PPNNonRetion(Stk, Fish) = (MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)) / _
                  (1 - MarkSelectiveMarkMisID(Fish, TStep) + MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
               PPNIncidental(Stk, Fish) = MarkSelectiveIncRate(Fish, TStep) / _
                  (1 - MarkSelectiveMarkMisID(Fish, TStep) + MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
            Else
               EncounterMort(Stk, Fish) = (MarkSelectiveUnMarkMisID(Fish, TStep) + _
                   (1 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep)) / (1 + MarkSelectiveIncRate(Fish, TStep))
               PPNLandedCat(Stk, Fish) = MarkSelectiveUnMarkMisID(Fish, TStep) / _
                  (MarkSelectiveUnMarkMisID(Fish, TStep) + (1 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
               PPNNonRetion(Stk, Fish) = ((1 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)) / _
                  (MarkSelectiveUnMarkMisID(Fish, TStep) + (1 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
               PPNIncidental(Stk, Fish) = MarkSelectiveIncRate(Fish, TStep) / _
                  (MarkSelectiveUnMarkMisID(Fish, TStep) + (1 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep) + MarkSelectiveIncRate(Fish, TStep))
            End If


            '- DEBUG Code print
            'If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
            '   If MarkSelectiveMarkMisID(Fish, TStep) <> 0 Then
            '      PrnLine = String.Format("{0,3}", Fish.ToString("##0"))
            '      PrnLine &= String.Format("{0,12}", FisheryName(Fish))
            '      PrnLine &= String.Format("{0,2}", TStep.ToString(" 0"))
            '      PrnLine &= String.Format("{0,4}", Stk.ToString("###0"))
            '      PrnLine &= String.Format("{0,8}", StockName(Stk))
            '      PrnLine &= String.Format("{0,8}", EncounterMort(Stk, Fish).ToString(" #0.0000"))
            '      PrnLine &= String.Format("{0,8}", PPNLandedCat(Stk, Fish).ToString(" #0.0000"))
            '      PrnLine &= String.Format("{0,8}", PPNNonRetion(Stk, Fish).ToString(" #0.0000"))
            '      PrnLine &= String.Format("{0,8}", PPNIncidental(Stk, Fish).ToString(" #0.0000"))
            '      sw.WriteLine(PrnLine)
            '   End If
            'End If



         Next
NextBiasRateFish:
      Next

      MSFBiasCount = 1
      SecondPass = False
SecondPassEntry:


      Age = 3
      LegalProportion = 1.0
      MSFtolerance = 0.0000001
      MSFBiasIter = False

      '- Zero ER Arrays explicitly because I can
      For Stk As Integer = 1 To NumStk
         StkERRate(Stk) = 0
         StkERRateTilde(Stk) = 0
         For Fish As Integer = 1 To NumFish
            NSStkERRate(Stk, Fish) = 0
            NSStkERRateTilde(Stk, Fish) = 0
            MSFStkERRate(Stk, Fish) = 0
            MSFStkERRateTilde(Stk, Fish) = 0
         Next
      Next

      '- Sum ER by Mark Type for All Stocks and All Fisheries
      For Stk As Integer = 1 To NumStk
            For Fish As Integer = 1 To NumFish

                
                If TerminalFisheryFlag(Fish, TStep) <> TerminalType Then GoTo NextERateFish
                '- Set Scaler to 0.1 for quota First Pass to 0.1 if undefined to avoid high total ER for time step
                'If SecondPass = False Then
                '   If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                '      If FisheryQuota(Fish, TStep) = 0 Then
                '         FisheryScaler(Fish, TStep) = 0
                '      Else
                '         If FisheryScaler(Fish, TStep) = 0 Then FisheryScaler(Fish, TStep) = 0.1
                '      End If
                '   End If
                '   If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                '      If MSFFisheryQuota(Fish, TStep) = 0 Then
                '         MSFFisheryScaler(Fish, TStep) = 0
                '      Else
                '         If MSFFisheryScaler(Fish, TStep) = 0 Then MSFFisheryScaler(Fish, TStep) = 0.1
                '      End If
                '   End If
                'End If
                '- Compute Rates
                FishERRate = 0
                If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                    FishERRate = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * FisheryScaler(Fish, TStep) * (1 + IncidentalRate(Fish, TStep))
                    NSStkERRate(Stk, Fish) += FishERRate
                    NSStkERRateTilde(Stk, Fish) += FishERRate
                End If

                'If FishERRate <> 0 Then
                '   PrnLine = String.Format("NS!{0,8}", StockName(Stk))
                '   PrnLine &= String.Format("{0,11}", FisheryName(Fish))
                '   PrnLine &= String.Format("{0,2}", TStep.ToString(" 0"))
                '   PrnLine &= String.Format("{0,10}", FishERRate.ToString(" #0.000000"))
                '   sw.WriteLine(PrnLine)
                'End If

                FishERRate = 0
                If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                    FishERRate = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep) * (1 + MarkSelectiveIncRate(Fish, TStep))
                    MSFStkERRate(Stk, Fish) += FishERRate
                    FishERRate = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep) * (1 + MarkSelectiveIncRate(Fish, TStep)) * EncounterMort(Stk, Fish)
                    MSFStkERRateTilde(Stk, Fish) += FishERRate
                End If

                '' ''If FishERRate <> 0 Then
                '   PrnLine = String.Format("MSF{0,8}", StockName(Stk))
                '   PrnLine &= String.Format("{0,11}", FisheryName(Fish))
                '   PrnLine &= String.Format("{0,2}", TStep.ToString(" 0"))
                '   PrnLine &= String.Format("{0,10}", FishERRate.ToString(" #0.000000"))
                '   sw.WriteLine(PrnLine)
                'End If

                '- Sum over all fisheries (Non-Selective and Selective)
                StkERRate(Stk) += (MSFStkERRate(Stk, Fish) + NSStkERRate(Stk, Fish))
                StkERRateTilde(Stk) += (MSFStkERRateTilde(Stk, Fish) + NSStkERRateTilde(Stk, Fish))
                If StkERRateTilde(Stk) > 1 Then
                    Jim = 1
                End If
                If StkERRate(Stk) > 1 And MSFBiasCount > 5 Then
                    MsgBox("Stock " & StockName(Stk) & "TStep " & TStep & " may produce negative escapements. Please finish the run and look for negative escapements in the PopStat report. Do not use this run for official results!")
                    Exit Sub
                End If
NextERateFish:
            Next
      Next
        
      Dim c1, c2, c3, c4 As Double

      '- Compute Bias Corrected Time Step ER & UnBiased Time Step ER
      For Stk% = 1 To NumStk%
         If StkERRate(Stk) = 0 Then
            Alpha(Stk) = 0
            Meeew(Stk) = 0

            '###################################################Pete-12/17/12.
         ElseIf StkERRate(Stk) > 1 Then
            ERgtrOne(TStep, Stk) = True
            '*************************************Pete-02/22/13
            If Cohort(Stk, Age, TerminalType, TStep) = 0 Then ERgtrOne(TStep, Stk) = False ' statement required to handle issue for stocks with zero abundance but fishery combos with ER>1
            '*************************************Pete-02/22/13
                StkERRate(Stk) = 0.9999
            '###################################################Pete-12/17/12.

         Else
            Alpha(Stk) = StkERRateTilde(Stk) / StkERRate(Stk)
            c1 = (1 - StkERRate(Stk))
            c2 = (1 - StkERRate(Stk)) ^ Alpha(Stk)
            c3 = c1 ^ Alpha(Stk)
                c4 = 1 - c3
            Meeew(Stk) = 1 - ((1 - StkERRate(Stk)) ^ Alpha(Stk))
         End If
      Next



      'sw.WriteLine("Stock Step StkERrate  Tilde       Delta     Meeew")
      'For Stk As Integer = 1 To NumStk
      '   'For Fish = 1 To NumFish
      '   If StkERRate(Stk) <> 0 Then
      '      PrnLine = String.Format("{0,8}", StockName(Stk))
      '      'PrnLine &= String.Format("{0,11}", FisheryName(Fish))
      '      PrnLine &= String.Format("{0,2}", TStep.ToString(" 0"))
      '      PrnLine &= String.Format("{0,10}", StkERRate(Stk).ToString(" #0.000000"))
      '      PrnLine &= String.Format("{0,10}", StkERRateTilde(Stk).ToString(" #0.000000"))
      '      PrnLine &= String.Format("{0,10}", Alpha(Stk).ToString(" #0.000000"))
      '      PrnLine &= String.Format("{0,10}", Meeew(Stk).ToString(" #0.000000"))
      '      sw.WriteLine(PrnLine)
      '   End If
      '   'Next
      'Next




      '- Zero Totals Array
      For Fish As Integer = 0 To NumFish
         '########################################################.Pete 1/2/13 
         'Must Bypass according to terminal type due to dependency of TAMM calculations on TotalLandedCatch
         If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
            '########################################################.Pete 1/2/13 
            TotalNSLanded(Fish) = 0
            TotalMSFLanded(Fish) = 0
            TotalLandedCatch(Fish, TStep) = 0
            TotalEncounters(Fish, TStep) = 0
         End If
      Next

      '- Compute Mortalities
      For Stk As Integer = 1 To NumStk
            '- Test for Zero
            If Stk = 33 Then
                Jim = 1
            End If
            
         If Cohort(Stk, Age, TerminalType, TStep) = 0 Or Meeew(Stk) = 0 Then
            StkMort(Stk) = 0
            GoTo NextStkMort
         End If
         '- Compute TimeStep Mortality for a Stock
         StkMort(Stk) = Cohort(Stk, Age, TerminalType, TStep) * Meeew(Stk)
         '- Compute Fishery/Time-Step Mortality for a Stock
            For Fish As Integer = 1 To NumFish
                
                If StkERRateTilde(Stk) = 0 Then
                    NSFishWeight(Stk, Fish) = 0
                    MSFishWeight(Stk, Fish) = 0
                Else
                    NSFishWeight(Stk, Fish) = NSStkERRateTilde(Stk, Fish) / StkERRateTilde(Stk)
                    MSFishWeight(Stk, Fish) = MSFStkERRateTilde(Stk, Fish) / StkERRateTilde(Stk)
                End If

                If NSFishWeight(Stk, Fish) = 0 Then
                    NSFishMort(Stk, Fish) = 0
                Else
                    NSFishMort(Stk, Fish) = StkMort(Stk) * NSFishWeight(Stk, Fish)
                End If

                If MSFishWeight(Stk, Fish) = 0 Then
                    MSFishMort(Stk, Fish) = 0
                Else
                    MSFishMort(Stk, Fish) = StkMort(Stk) * MSFishWeight(Stk, Fish)
                End If

                'If Fish = 187 And NSFishMort(Stk, Fish) <> 0 Then

                'End If
                'Jim = 1

                If NSFishMort(Stk, Fish) = 0 Then
                    LandedCatch(Stk, Age, Fish, TStep) = 0
                    Encounters(Stk, Age, Fish, TStep) = 0
                    DropOff(Stk, Age, Fish, TStep) = 0
                Else
                    '- Compute Non-Selective Landed and Non-Landed Mortality by Stock, Fishery, Time-Step
                    LandedCatch(Stk, Age, Fish, TStep) = NSFishMort(Stk, Fish) / (1 + IncidentalRate(Fish, TStep))
                    'If Stk = 1 And Fish = 1 And TStep = 4 Then
                    'Debug.Print(LandedCatch(Stk, Age, Fish, TStep))
                    'End If
                    Encounters(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
                    DropOff(Stk, Age, Fish, TStep) = NSFishMort(Stk, Fish) - LandedCatch(Stk, Age, Fish, TStep)
                    TotalNSLanded(Fish) += LandedCatch(Stk, Age, Fish, TStep)
                    TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                    TotalEncounters(Fish, TStep) += Encounters(Stk, Age, Fish, TStep)
                End If
                If MSFishMort(Stk, Fish) = 0 Then
                    MSFLandedCatch(Stk, Age, Fish, TStep) = 0
                    MSFNonRetention(Stk, Age, Fish, TStep) = 0
                    MSFEncounters(Stk, Age, Fish, TStep) = 0
                    MSFDropOff(Stk, Age, Fish, TStep) = 0
                Else
                    '- Compute Non-Selective Landed and Non-Landed Mortality by Stock, Fishery, Time-Step
                    MSFLandedCatch(Stk, Age, Fish, TStep) = MSFishMort(Stk, Fish) * PPNLandedCat(Stk, Fish)
                    MSFNonRetention(Stk, Age, Fish, TStep) = MSFishMort(Stk, Fish) * PPNNonRetion(Stk, Fish)
                    MSFEncounters(Stk, Age, Fish, TStep) = Cohort(Stk, Age, TerminalType, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * MSFFisheryScaler(Fish, TStep)
                    MSFDropOff(Stk, Age, Fish, TStep) = MSFishMort(Stk, Fish) * PPNIncidental(Stk, Fish)
                    TotalMSFLanded(Fish) += MSFLandedCatch(Stk, Age, Fish, TStep)
                    TotalLandedCatch(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                    TotalEncounters(Fish, TStep) += MSFEncounters(Stk, Age, Fish, TStep)
                End If
            Next
NextStkMort:
      Next

      '-debug
      'Fish = 136
      'sw.WriteLine("Stock    Fishery   TS   LCatch   NSFishMo  NSFshWght  Cohort     Meew       Alpha   NSStkRtTl StkRtTlde  quota      scalar   IncRate   BPER     ")
      'sw.WriteLine("-------- ---------- -  --------- --------- --------- --------- ---------- --------- --------- --------- --------- --------- --------- --------- ")
      'For Stk = 1 To NumStk
      '   If LandedCatch(Stk, Age, Fish, TStep) <> 0 Then
      '      PrnLine = String.Format("{0,8}", StockName(Stk))
      '      PrnLine &= String.Format("{0,11}", FisheryName(Fish))
      '      PrnLine &= String.Format("{0,2}", TStep.ToString(" 0"))
      '      PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString(" ######0.0"))
      '      PrnLine &= String.Format("{0,10}", NSFishMort(Stk, Fish).ToString(" ######0.0"))
      '      PrnLine &= String.Format("{0,10}", NSFishWeight(Stk, Fish).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", Cohort(Stk, Age, TerminalType, TStep).ToString(" ######0.0"))
      '      PrnLine &= String.Format("{0,10}", Meeew(Stk).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", Alpha(Stk).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", NSStkERRateTilde(Stk, Fish).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", StkERRateTilde(Stk).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", FisheryQuota(Fish, TStep).ToString(" ######0.0"))
      '      PrnLine &= String.Format("{0,10}", FisheryScaler(Fish, TStep).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", IncidentalRate(Fish, TStep).ToString(" ##0.00000"))
      '      PrnLine &= String.Format("{0,10}", BaseExploitationRate(Stk, Age, Fish, TStep).ToString(" ##0.00000"))
      '      sw.WriteLine(PrnLine)
      '   End If
      'Next


      '- Compute Fishery Scalers for Next Iteration and Check for Convergence Tolerance
        For Fish As Integer = 1 To NumFish
            If Fish = 112 And TStep = 4 And MSFBiasCount = 475 Then
                Jim = 1
            End If
            If TerminalFisheryFlag(Fish, TStep) <> TerminalType Then GoTo NextTolerCheck
            If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                If TotalNSLanded(Fish) = 0 Then
                    'If Fish = 144 Then
                    'MessageBox.Show("Fish" & Fish & "  TS" & TStep & "  tt" & TerminalType)
                    'End If
                    FisheryScaler(Fish, TStep) = 0
                Else
                    If FisheryQuota(Fish, TStep) = 0 Then
                        FisheryScaler(Fish, TStep) = 0
                    Else
                        'Try
                        FisheryScaler(Fish, TStep) *= (FisheryQuota(Fish, TStep) / TotalNSLanded(Fish))


                        'If FisheryScaler(Fish, TStep) > 200000 Then
                        '    MsgBox("Not enough abundance to fill the quota in fishery " & Fish & " Time Step " & TStep & ". Try to manually increase the abundance of the main stock.")
                        '    End
                        'End If
                        'Catch ex As Exception
                        '    MsgBox("Fishery " & Fish & " TimeStep " & TStep & "do not have enough fish to meet the quota.")
                        'End Try
                    End If
                    MSFTestTolerance = FisheryQuota(Fish, TStep) / TotalNSLanded(Fish)
                    'If MSFTestTolerance > MSFtolerance Then MSFBiasIter = True
                    If MSFTestTolerance > (1 + MSFtolerance) Or MSFTestTolerance < (1 - MSFtolerance) Then MSFBiasIter = True
                End If
            End If
            If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                If TotalMSFLanded(Fish) = 0 Then
                    MSFFisheryScaler(Fish, TStep) = 0
                Else
                    If MSFFisheryQuota(Fish, TStep) = 0 Then
                        MSFFisheryScaler(Fish, TStep) = 0
                    Else
                        MSFFisheryScaler(Fish, TStep) *= (MSFFisheryQuota(Fish, TStep) / TotalMSFLanded(Fish))
                    End If
                    MSFTestTolerance = MSFFisheryQuota(Fish, TStep) / TotalMSFLanded(Fish)
                    If MSFTestTolerance > (1 + MSFtolerance) Or MSFTestTolerance < (1 - MSFtolerance) Then MSFBiasIter = True
                End If
            End If
NextTolerCheck:
        Next

      SecondPass = True

        If MSFBiasIter = True Then
            If MSFBiasCount = 233 Then
                Jim = 1
            End If
            MSFBiasCount += 1
            If MSFBiasCount = 475 Then
                Jim = 1
            End If
            GoTo SecondPassEntry
        End If

   End Sub

    '======================================================================================

    Sub CompCNR(ByVal Fish, ByVal TerminalType, ByVal EncRate)
        '**************************************************************************
        'Subroutine computes CNR mortality in fisheries using one of four methods
        ' depending upon value of NonRetentionFlag parameter;
        ' NonRetentionFlag = 1 - Computed from fishery scale factor
        ' NonRetentionFlag = 2 - Ratio of CNR days to normal days
        ' NonRetentionFlag = 3 - External estimate of legal & sublegal encounters
        ' NonRetentionFlag = 4 - External estimate of total encounters
        ' Note: These are +1 Flag Values from old Fram Program
        '**************************************************************************

        ReDim PropLegCatch(NumStk, MaxAge)
        Dim CNRScale, CNREncounter As Double
        Dim LegProp, SubLegProp As Double
        If TotalLandedCatch(Fish, TStep) = 0 And (NonRetentionFlag(Fish, TStep) = 1 Or NonRetentionFlag(Fish, TStep) = 2) Then
            MsgBox("CNR Method 1 or 2 cannot be used in fishery with no catch" & vbCrLf & "Fishery= " & FisheryName(Fish) & " TStep= " & TStep.ToString & vbCrLf & "Must Fix this Error before Continuing!", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        '--------------- Four METHODS OF COMPUTING CNR CATCH --------------

        Select Case NonRetentionFlag(Fish, TStep)
            Case 1                                   '...Computed CNR
                If FisheryScaler(Fish, TStep) < 1 Then
                    For Stk = 1 To NumStk
                  For Age As Integer = MinAge To MaxAge
                     NonRetention(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                     NonRetention(Stk, Age, Fish, TStep) += TotalLandedCatch(Fish, TStep) * EncRate * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * PropSubPop(Stk, Age) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 3)
                     TotalNonRetention(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep)
                  Next Age
                    Next Stk
                End If

            Case 2                  '...Ratio of CNR days to normal days
                For Stk = 1 To NumStk
               For Age As Integer = MinAge To MaxAge
                  NonRetention(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                  NonRetention(Stk, Age, Fish, TStep) += Shakers(Stk, Age, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * NonRetentionInput(Fish, TStep, 3)
                  TotalNonRetention(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep)
               Next Age
                Next Stk

            Case 3                 '...External estimate of encounters

                Call CompPropCatch(Fish, TerminalType)
                For Stk = 1 To NumStk
               For Age As Integer = MinAge To MaxAge
                  LegProp = PropLegCatch(Stk, Age)
                  SubLegProp = PropSubPop(Stk, Age)
                  '- PS Sport legal size rel mort rate set now to 50% of shaker release rate (10 vs 20)
                  If Fish >= 36 And InStr(FisheryTitle(Fish), "Sport") > 0 Then
                     NonRetention(Stk, Age, Fish, TStep) = (LegProp * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * (ShakerMortRate(Fish, TStep) / 2))
                            NRLegal(1, Stk, Age, Fish, TStep) = LegProp * NonRetentionInput(Fish, TStep, 1)
                        Else
                            NonRetention(Stk, Age, Fish, TStep) = (LegProp * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep))
                            NRLegal(1, Stk, Age, Fish, TStep) = LegProp * NonRetentionInput(Fish, TStep, 1)
                        End If
                  NonRetention(Stk, Age, Fish, TStep) += (SubLegProp * NonRetentionInput(Fish, TStep, 2) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep))
                        NRLegal(2, Stk, Age, Fish, TStep) = SubLegProp * NonRetentionInput(Fish, TStep, 2)
                        TotalNonRetention(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep)
               Next Age
                Next Stk

            Case 4
                '--- Total Encounters Estimate (Legal + SubLegal)
                '--- Selective Fishery Sampling Estimates

                Dim CNREncStkAge(NumStk, MaxAge), PreSubCNR, SubCNR As Double
                If NonRetentionInput(Fish, TStep, 1) = 0 Then Exit Sub
                CNREncounter = 0

                For Stk = 1 To NumStk
               For Age As Integer = MinAge To MaxAge
                  'Call CompLegProp(Stk, Age, Fish, TerminalType, SubLegalProportion, LegalProportion)
                  Call CompLegProp(Stk, Age, Fish, TerminalType)


                  ''****************************************************************************************
                  ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                  ''Given that Puget Sound Chinook CNR inputs are roughly 'calibrated' to a 22" scenario AND that,
                  ''CNR impacts should remain constant regardless of size limits modeled during retention periods, 
                  ''prefer to use the Legal/Sublegal fractions based on this 'size limit' 
                  'If SizeLimitScenario = True Then
                  '   Select Case Fish
                  '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                  '         LegalProportion = CNRLegalProp
                  '         SubLegalProportion = CNRSublegalProp
                  '   End Select
                  'End If
                  ''############################# END NEW CODE ############################ Pete-Jan. 2013
                  ''****************************************************************************************



                  '- Zero Time 1 Yearling Shakers - Fish not Recruited Yet
                  If NumStk > 50 Then '- Sel.Fish Version Stock Numbers
                     If Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                        SubLegalPop = 0
                     Else
                        SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion
                     End If
                  Else
                     If Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
                        SubLegalPop = 0
                     Else
                        SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion
                     End If
                  End If
                  '- Legal Size Encounters and Mortality
                  CNREncStkAge(Stk, Age) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                  CNREncounter += CNREncStkAge(Stk, Age)
                  '- PS Sport legal size rel mort rate set now to 50 of shaker release rate (10 vs 20)
                  If Fish >= 36 And InStr(FisheryTitle(Fish), "Sport") > 0 Then
                            NonRetention(Stk, Age, Fish, TStep) += (CNREncStkAge(Stk, Age) * (ShakerMortRate(Fish, TStep) / 2))
                            NRLegal(1, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep) / (ShakerMortRate(Fish, TStep) / 2)
                  Else
                            NonRetention(Stk, Age, Fish, TStep) += (CNREncStkAge(Stk, Age) * ShakerMortRate(Fish, TStep))
                            NRLegal(1, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep) / ShakerMortRate(Fish, TStep)
                  End If
                  PreSubCNR = NonRetention(Stk, Age, Fish, TStep)
                  '- SubLegal Size Encounters and Mortality - PFMC Mar 2006 ... Added StkHRScale
                  CNREncStkAge(Stk, Age) += (SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep))
                  CNREncounter += (SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep))
                  NonRetention(Stk, Age, Fish, TStep) += (SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep))
                        'NRLegal(1, Stk, Age, Fish, TStep) += LegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                        NRLegal(2, Stk, Age, Fish, TStep) += SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                        SubCNR = (SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep))
               Next Age
                Next Stk
                '- Calculate scaler so the Total Encounters will equal Input Value
                If CNREncounter <> 0 Then
                    CNRScale = NonRetentionInput(Fish, TStep, 1) / CNREncounter
                    PrnLine = "CNRScale ft=" & String.Format("{0,2}", Fish) & String.Format("{0,2}", TStep) & " scaler=" & String.Format("{0,12}", CNRScale.ToString(" ###0.0000")) & " Input=" & String.Format("{0,10}", NonRetentionInput(Fish, TStep, 1).ToString(" #####0.0") & " TotEnc=" & String.Format("{0,9}", CNREncounter.ToString("#####0.000")))
                    sw.WriteLine(PrnLine)
                End If
                '- Apply Fishery Scaler to Morts & Encounters
                CNREncounter = 0
            For Stk As Integer = 1 To NumStk
               For Age As Integer = MinAge To MaxAge
                  If CNREncStkAge(Stk, Age) <> 0 Then
                     CNREncStkAge(Stk, Age) = CNREncStkAge(Stk, Age) * CNRScale * ModelStockProportion(Fish)
                     NonRetention(Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep) * CNRScale * ModelStockProportion(Fish)
                            NRLegal(1, Stk, Age, Fish, TStep) = NRLegal(1, Stk, Age, Fish, TStep) * CNRScale
                            NRLegal(2, Stk, Age, Fish, TStep) = NRLegal(2, Stk, Age, Fish, TStep) * CNRScale
                            CNREncounter += Encounters(Stk, Age, Fish, TStep)
                     TotalNonRetention(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep)
                  End If
               Next Age
            Next Stk
                PrnLine = "Total Model Stock Encounters/Morts=" & String.Format("{0,14}", CNREncounter.ToString(" #########0.00")) & String.Format("{0,14}", TotalNonRetention(Fish, TStep).ToString(" ######0.000"))
                sw.WriteLine(PrnLine)
                PrnLine = "Total Fishery     Encounters/Morts=" & String.Format("{0,14}", (CNREncounter / ModelStockProportion(Fish)).ToString(" #########0.00")) & String.Format("{0,14}", (TotalNonRetention(Fish, TStep) / ModelStockProportion(Fish)).ToString(" ######0.000"))
                sw.WriteLine(PrnLine)
            Case Else
        End Select

    End Sub

    Sub CompCohoCNR(ByVal Fish, ByVal TerminalType)
        '**************************************************************************
        'Subroutine computes CNR mortality in fisheries using one method
        ' NonRetentionFlag = 4 - Total CNR Mortality for Fishery
        '**************************************************************************

        Dim PropLandedCatch(NumStk) As Double
        Dim TempLandedCatch As Double

        TempLandedCatch = 0
        Age = 3                      'SUM THE CATCH OVER ALL STOCKS
        For Stk = 1 To NumStk
            PropLandedCatch(Stk) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, Age, Fish, TStep) * Cohort(Stk, Age, TerminalType, TStep)
            TempLandedCatch = TempLandedCatch + PropLandedCatch(Stk)
        Next Stk

        'COMPUTE PROPORTION OF CATCH WHICH EACH STOCK COMPRISES

        If TempLandedCatch <> 0.0 Then
            For Stk = 1 To NumStk
                PropLandedCatch(Stk) = PropLandedCatch(Stk) / TempLandedCatch
            Next Stk
            For Stk = 1 To NumStk
                NonRetention(Stk, Age, Fish, TStep) = NonRetentionInput(Fish, TStep, 1) * PropLandedCatch(Stk) * ModelStockProportion(Fish)
            Next Stk
            TotalNonRetention(Fish, TStep) = NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish)
        Else
            If AnyBaseRate(Fish, TStep) = 0 Then
                MsgBox("ERROR - NonRetention Catch for Zero Base Period Catch" & vbCrLf & "TIME,FISH = " & TStep.ToString & " " & Fish.ToString & " " & FisheryName(Fish), MsgBoxStyle.OkOnly)
            End If
            TotalNonRetention(Fish, TStep) = 0.0
        End If

    End Sub

    Sub CompEscape()
        If RunTAMMIter = 1 And TStep = 5 Then
            AnyNegativeEscapement = 0 ' reset negative escapement flag to zero if escapements are recomputed during TAMM iterations
        End If
        '- COMPUTE ESCAPE BY SUBTRACTING CATCH AND INCIDENTAL
        '- MORTALITY FROM THE MATURE POPULATION
        For Stk As Integer = 1 To NumStk
            
            For Age As Integer = MinAge To MaxAge
                If Age = 3 Then
                    Jim = 1
                End If
                Escape(Stk, Age, TStep) = Cohort(Stk, Age, Term, TStep)
                For Fish As Integer = 1 To NumFish
                    If TerminalFisheryFlag(Fish, TStep) = Term Then
                        Escape(Stk, Age, TStep) = Escape(Stk, Age, TStep) - LandedCatch(Stk, Age, Fish, TStep) - Shakers(Stk, Age, Fish, TStep) - NonRetention(Stk, Age, Fish, TStep) - DropOff(Stk, Age, Fish, TStep) - MSFLandedCatch(Stk, Age, Fish, TStep) - MSFShakers(Stk, Age, Fish, TStep) - MSFNonRetention(Stk, Age, Fish, TStep) - MSFDropOff(Stk, Age, Fish, TStep)
                        If Escape(Stk, Age, TStep) < -1 Then
                            AnyNegativeEscapement = 1
                        End If
                    End If
                Next Fish
            Next Age

        Next Stk
        If TStep = 5 Then
            Jim = 1
        End If
        '- CHINOOK TAMM Escapement Arrays
        If (RunTAMMIter = 1 And SpeciesName = "CHINOOK") Then
            If NumStk < 50 Then
                '------ Nooksack Fall Chinook
                TammEscape(1, TStep) = Escape(1, 3, TStep) + Escape(1, 4, TStep) + Escape(1, 5, TStep)
                '------ Skagit Fall Fingerling and Yearling Chinook
                TammEscape(2, TStep) = Escape(4, 3, TStep) + Escape(4, 4, TStep) + Escape(4, 5, TStep) + Escape(5, 3, TStep) + Escape(5, 4, TStep) + Escape(5, 5, TStep)
                '------ Snoh Fingerling and Yearling, Stillag., and Tulalip
                TammEscape(3, TStep) = Escape(7, 3, TStep) + Escape(7, 4, TStep) + Escape(7, 5, TStep) + Escape(8, 3, TStep) + Escape(8, 4, TStep) + Escape(8, 5, TStep) + Escape(9, 3, TStep) + Escape(9, 4, TStep) + Escape(9, 5, TStep) + Escape(10, 3, TStep) + Escape(10, 4, TStep) + Escape(10, 5, TStep)
                '------ Tulalip Rates use Tulalip ETRS
                TammEscape(4, TStep) = Escape(10, 3, TStep) + Escape(10, 4, TStep) + Escape(10, 5, TStep)
                '------ Hood Canal Fall Fingerlings and Yearlings
                TammEscape(5, TStep) = Escape(16, 3, TStep) + Escape(16, 4, TStep) + Escape(16, 5, TStep) + Escape(17, 3, TStep) + Escape(17, 4, TStep) + Escape(17, 5, TStep)
                '------ Nooksack Native Spring Chinook
                TammEscape(6, TStep) = Escape(2, 3, TStep) + Escape(2, 4, TStep) + Escape(2, 5, TStep) + Escape(3, 3, TStep) + Escape(3, 4, TStep) + Escape(3, 5, TStep)
                '------ South Sound Spring Yearling (White River at Minter Crk)
                TammEscape(7, TStep) = Escape(15, 3, TStep) + Escape(15, 4, TStep) + Escape(15, 5, TStep)
            Else
                '------ Nooksack Fall Chinook
                TammEscape(1, TStep) = Escape(1, 3, TStep) + Escape(1, 4, TStep) + Escape(1, 5, TStep) + Escape(2, 3, TStep) + Escape(2, 4, TStep) + Escape(2, 5, TStep)
                '------ Skagit Fall Fingerling and Yearling Chinook
                TammEscape(2, TStep) = Escape(7, 3, TStep) + Escape(7, 4, TStep) + Escape(7, 5, TStep) + Escape(8, 3, TStep) + Escape(8, 4, TStep) + Escape(8, 5, TStep) + Escape(9, 3, TStep) + Escape(9, 4, TStep) + Escape(9, 5, TStep) + Escape(10, 3, TStep) + Escape(10, 4, TStep) + Escape(10, 5, TStep)
                '------ Snoh Fingerling and Yearling, Stillag., and Tulalip
                TammEscape(3, TStep) = Escape(13, 3, TStep) + Escape(13, 4, TStep) + Escape(13, 5, TStep) + Escape(14, 3, TStep) + Escape(14, 4, TStep) + Escape(14, 5, TStep) + Escape(15, 3, TStep) + Escape(15, 4, TStep) + Escape(15, 5, TStep) + Escape(16, 3, TStep) + Escape(16, 4, TStep) + Escape(16, 5, TStep)
                TammEscape(3, TStep) = TammEscape(3, TStep) + Escape(17, 3, TStep) + Escape(17, 4, TStep) + Escape(17, 5, TStep) + Escape(18, 3, TStep) + Escape(18, 4, TStep) + Escape(18, 5, TStep) + Escape(19, 3, TStep) + Escape(19, 4, TStep) + Escape(19, 5, TStep) + Escape(20, 3, TStep) + Escape(20, 4, TStep) + Escape(20, 5, TStep)
                '------ Tulalip Rates use Tulalip ETRS
                TammEscape(4, TStep) = Escape(19, 3, TStep) + Escape(19, 4, TStep) + Escape(19, 5, TStep) + Escape(20, 3, TStep) + Escape(20, 4, TStep) + Escape(20, 5, TStep)
                '------ Hood Canal Fall Fingerlings and Yearlings
                TammEscape(5, TStep) = Escape(31, 3, TStep) + Escape(31, 4, TStep) + Escape(31, 5, TStep) + Escape(32, 3, TStep) + Escape(32, 4, TStep) + Escape(32, 5, TStep) + Escape(33, 3, TStep) + Escape(33, 4, TStep) + Escape(33, 5, TStep) + Escape(34, 3, TStep) + Escape(34, 4, TStep) + Escape(34, 5, TStep)
                '------ Nooksack Native Spring Chinook
                TammEscape(6, TStep) = Escape(3, 3, TStep) + Escape(3, 4, TStep) + Escape(3, 5, TStep) + Escape(4, 3, TStep) + Escape(4, 4, TStep) + Escape(4, 5, TStep) + Escape(5, 3, TStep) + Escape(5, 4, TStep) + Escape(5, 5, TStep) + Escape(6, 3, TStep) + Escape(6, 4, TStep) + Escape(6, 5, TStep)
                '------ South Sound Spring Yearling (White River at Minter Crk)
                TammEscape(7, TStep) = Escape(29, 3, TStep) + Escape(29, 4, TStep) + Escape(29, 5, TStep) + Escape(30, 3, TStep) + Escape(30, 4, TStep) + Escape(30, 5, TStep) + Escape(65, 3, TStep) + Escape(65, 4, TStep) + Escape(65, 5, TStep) + Escape(66, 3, TStep) + Escape(66, 4, TStep) + Escape(66, 5, TStep)
            End If
        End If

    End Sub

    Sub CompLegProp(ByVal Stk, ByVal Age, ByVal Fish, ByVal TerminalType)
        '********************************************************************************************
        '- Compute Legal Size Proportion by Stock, Age, Maturity and Fishery (i.e. Size Limit)
        '********************************************************************************************
        Dim KTime, MeanSize, SizeStdDev As Double

        '--------- ALL 3 YR Old COHO considered legal ---
        If SpeciesName = "COHO" Then
            LegalProportion = 1.0
            SubLegalProportion = 0.0
            Exit Sub
        End If
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        If Stk = 1 And Age = 2 And Fish = 36 And TStep = 2 Then 'used to break code
            Jim = 1
        End If
        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
        'COMPUTE MEAN LENGTH OF FISH AND SD
        KTime = (Age - 1) * 12 + MidTimeStep(TStep)
        'KTime = Age * 12 + MidTimeStep(TStep) 'for Pete's new growth functions
        MeanSize = VonBertL(Stk, TerminalType) * (1.0 - Exp(-VonBertK(Stk, TerminalType) * (KTime - VonBertT(Stk, TerminalType))))
        SizeStdDev = VonBertCV(Stk, Age, TerminalType) * MeanSize



      ''****************************************************************************************
      ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
      'Dim TempMinLimit As Double 'Storage place for database size limit value, for restoring after run

      'If SizeLimitScenario = True Then
      '   Select Case Fish
      '      'Only invoke this code for PS Sport at this stage
      '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals 
      '         TempMinLimit = MinSizeLimit(Fish, TStep) 'Place database value here for temporary storage
      '         If ShakerFlagMSF(Fish, TStep) >= 1 Or ShakerFlagNS(Fish, TStep) >= 1 Then
      '            If (AltFlag(Fish, TStep) <= 8 And AltFlag(Fish, TStep) > 0) Or AltLimitMSF(Fish, TStep) = AltLimitNS(Fish, TStep) Then
      '               If AltFlag(Fish, TStep) > 2 And AltFlag(Fish, TStep) <= 8 Then
      '                  MinSizeLimit(Fish, TStep) = AltLimitMSF(Fish, TStep) 'Replace with MSF input
      '               Else
      '                  MinSizeLimit(Fish, TStep) = AltLimitNS(Fish, TStep) 'Replace with NS input (doesn't matter if NS or MSF if combo with equal limits)
      '               End If
      '            End If
      '         End If
      '   End Select
      'End If
      ''############################# END NEW CODE ############################ Pete-Jan. 2013
      ''****************************************************************************************










      'COMPUTE PROPORTION OF COHORT LARGER THAN SIZE LIMIT for Current Size Limit
        If ChinookBaseLegProp = False Then
            If (MinSizeLimit(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
                LegalProportion = 1
            End If
            If (MinSizeLimit(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
                LegalProportion = 0
            End If
            If ((MinSizeLimit(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (MinSizeLimit(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
                LegalProportion = (1 - NormlDistr(MinSizeLimit(Fish, TStep), MeanSize, SizeStdDev))
            End If
            'COMPUTE SUBLEGAL PROPORTION  (AGE 2 ADJUSTED TO ACCOUNT FOR UNRECRUITED PROPORTION AND DISTRIBUTION)
            SubLegalProportion = EncounterRateAdjustment(Age, Fish, TStep) * (1 - LegalProportion)
        Else
            '- Chinook Legal Proportion in Base Period - Used in Shaker Calculations
            If (ChinookBaseSizeLimit(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
                BaseLegalProportion = 1
            End If
            If (ChinookBaseSizeLimit(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
                BaseLegalProportion = 0
            End If
            If ((ChinookBaseSizeLimit(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (ChinookBaseSizeLimit(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
                BaseLegalProportion = (1 - NormlDistr(ChinookBaseSizeLimit(Fish, TStep), MeanSize, SizeStdDev))
            End If
            'COMPUTE SUBLEGAL PROPORTION  (AGE 2 ADJUSTED TO ACCOUNT FOR UNRECRUITED PROPORTION AND DISTRIBUTION)
            BaseSubLegalProportion = EncounterRateAdjustment(Age, Fish, TStep) * (1 - BaseLegalProportion)
      End If








      ''****************************************************************************************
      ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
      'If SizeLimitScenario = True Then

      '   Select Case Fish
      '      'Only invoke this code for PS Sport at this stage
      '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals 


      '         'Compute CNR Legal and Sublegal proportions based on the existing 22 inch (= 520 mm fork length) size limit
      '         'that CNR encounters are calibrated to for Puget Sound recreational fisheries
      '         If (520 < MeanSize - 3 * SizeStdDev) Then
      '            CNRLegalProp = 1
      '         End If
      '         If (520 > MeanSize + 3 * SizeStdDev) Then
      '            CNRLegalProp = 0
      '         End If
      '         If ((520 >= MeanSize - 3 * SizeStdDev) And (520 <= MeanSize + 3 * SizeStdDev)) Then
      '            CNRLegalProp = (1 - NormlDistr(520, MeanSize, SizeStdDev))
      '         End If
      '         CNRSublegalProp = EncounterRateAdjustment(Age, Fish, TStep) * (1 - CNRLegalProp) 'BEWARE OF EncounterRateAdjustments...


      '         If AltFlag(Fish, TStep) > 8 And AltLimitMSF(Fish, TStep) <> AltLimitNS(Fish, TStep) _
      '            And AltLimitMSF(Fish, TStep) > 0 And AltLimitNS(Fish, TStep) > 0 Then
      '            'Compute Legal & Sublegal Proportions under MSF size limit
      '            If (AltLimitMSF(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
      '               MSFLegalProp = 1
      '            End If
      '            If (AltLimitMSF(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
      '               MSFLegalProp = 0
      '            End If
      '            If ((AltLimitMSF(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (AltLimitMSF(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
      '               MSFLegalProp = (1 - NormlDistr(AltLimitMSF(Fish, TStep), MeanSize, SizeStdDev))
      '            End If
      '            MSFSublegalProp = EncounterRateAdjustment(Age, Fish, TStep) * (1 - MSFLegalProp) 'BEWARE OF EncounterRateAdjustments...

      '            'Compute Legal & Sublegal Proportions under NS size limit
      '            If (AltLimitNS(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
      '               NSLegalProp = 1
      '            End If
      '            If (AltLimitNS(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
      '               NSLegalProp = 0
      '            End If
      '            If ((AltLimitNS(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (AltLimitNS(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
      '               NSLegalProp = (1 - NormlDistr(AltLimitNS(Fish, TStep), MeanSize, SizeStdDev))
      '            End If
      '            NSSublegalProp = EncounterRateAdjustment(Age, Fish, TStep) * (1 - NSLegalProp) 'BEWARE OF EncounterRateAdjustments...

      '         End If
      '         MinSizeLimit(Fish, TStep) = TempMinLimit 'Restore original value to database
      '   End Select
      'End If
      ''############################# END NEW CODE ############################ Pete-Jan. 2013
      ''****************************************************************************************




    End Sub

    Sub CompOthMort(ByVal Fish)
        '**************************************************************************
        'Subroutine computes dropoff, dropout, and predation mortality in each
        ' fishery based upon an input proportion of the catch.  
        '**************************************************************************

        'COMPUTE OTHER MORTALITY IF CATCH OCCURRED IN FISHERY
        If TotalLandedCatch(Fish, TStep) > 0 Then
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               '- DropOff Already Done for Selective Fisheries
               If FisheryFlag(Fish, TStep) < 6 Or FisheryFlag(Fish, TStep) > 10 Then
                  DropOff(Stk, Age, Fish, TStep) = IncidentalRate(Fish, TStep) * LandedCatch(Stk, Age, Fish, TStep)
                  TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + IncidentalRate(Fish, TStep) * LandedCatch(Stk, Age, Fish, TStep)
               End If
            Next Age
         Next Stk
        End If

    End Sub

    Sub CompPropCatch(ByVal Fish, ByVal TerminalType)
        '**************************************************************************
        'Subroutine computes proportion of legal and sublegal population which
        ' which each stock, and class comprises.  The proportion is
        ' used to allocate mortality.
        '**************************************************************************
        Dim TempCatch, TotalSubCNR As Double
        Dim Stock, ComAge As Integer

        SubLegalPop = 0
        TempCatch = 0

        'SUM THE CATCH OVER ALL STOCKS

        For Stock = 1 To NumStk
            For ComAge = MinAge To MaxAge
                'Call CompLegProp(Stock, ComAge, Fish, TerminalType, SubLegalProportion, LegalProportion)
            Call CompLegProp(Stock, ComAge, Fish, TerminalType)

                If Stock = 36 Then
                    Jim = 1
                End If
            ''****************************************************************************************
            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
            ''Given that Puget Sound Chinook CNR inputs are roughly 'calibrated' to a 22" scenario AND that,
            ''CNR impacts should remain constant regardless of size limits modeled during retention periods, 
            ''prefer to use the Legal/Sublegal fractions based on this 'size limit' 
            'If SizeLimitScenario = True Then
            '   Select Case Fish
            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals 
            '         LegalProportion = CNRLegalProp
            '   End Select
            'End If
            ''############################# END NEW CODE ############################ Pete-Jan. 2013
            '****************************************************************************************

                PropLegCatch(Stock, ComAge) = StockFishRateScalers(Stock, Fish, TStep) * BaseExploitationRate(Stock, ComAge, Fish, TStep) * Cohort(Stock, ComAge, TerminalType, TStep) * LegalProportion
                TempCatch = TempCatch + PropLegCatch(Stock, ComAge)
            Next ComAge
        Next Stock
        If TempCatch = 0 Then Exit Sub

        'COMPUTE PROPORTION OF CATCH WHICH EACH STOCK COMPRISES

        For Stock = 1 To NumStk
            For ComAge = MinAge To MaxAge
                PropLegCatch(Stock, ComAge) = PropLegCatch(Stock, ComAge) / TempCatch
            Next ComAge
        Next Stock

        'SUM UP SUBLEGAL POPULATION

        ReDim PropSubPop(NumStk, MaxAge)
        Dim CNRShakers(NumStk, MaxAge) As Double

        TotalSubCNR = 0
        For Stock = 1 To NumStk
            For ComAge = MinAge To MaxAge
                'Call CompLegProp(Stock, ComAge, Fish, TerminalType, SubLegalProportion, LegalProportion)
            Call CompLegProp(Stock, ComAge, Fish, TerminalType)


            ''****************************************************************************************
            ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
            ''Given that Puget Sound Chinook CNR inputs are roughly 'calibrated' to a 22" scenario AND that,
            ''CNR impacts should remain constant regardless of size limits modeled during retention periods, 
            ''prefer to use the Legal/Sublegal fractions based on this 'size limit' 
            'If SizeLimitScenario = True Then
            '   Select Case Fish
            '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals 
            '         SubLegalProportion = CNRSublegalProp
            '   End Select
            'End If
            ''############################# END NEW CODE ############################ Pete-Jan. 2013
            ''****************************************************************************************

                '- Zero Time 1 Yearling Shakers ...
                '- Fish not Recruited Yet - Temp Fix 1/3/2000 JFP
                If NumStk > 50 Then '- Sel.Fish Version Stock Numbers
                    If ComAge = 2 And (TStep = 1 Or TStep = 4) And (Stock = 9 Or Stock = 10 Or Stock = 11 Or Stock = 12 Or Stock = 15 Or Stock = 16 Or Stock = 27 Or Stock = 28 Or Stock = 33 Or Stock = 34 Or Stock = 49 Or Stock = 50) Then
                        SubLegalPop = 0
                    Else
                        SubLegalPop = Cohort(Stock, ComAge, TerminalType, TStep) * SubLegalProportion
                    End If
                Else
                    If ComAge = 2 And (TStep = 1 Or TStep = 4) And (Stock = 5 Or Stock = 6 Or Stock = 8 Or Stock = 14 Or Stock = 17 Or Stock = 25) Then
                        SubLegalPop = 0
                    Else
                        SubLegalPop = Cohort(Stock, ComAge, TerminalType, TStep) * SubLegalProportion
                    End If
                End If
                '- PFMC Mar 2006 ... Added StkHRScale
                CNRShakers(Stock, ComAge) = SubLegalPop * BaseSubLegalRate(Stock, ComAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stock, Fish, TStep)
                TotalSubCNR = TotalSubCNR + CNRShakers(Stock, ComAge)
            Next ComAge
        Next Stock

        'COMPUTE PROPORTION OF SUBLEGAL POPULATION WHICH EACH STOCK COMPRISES

        For Stock = 1 To NumStk
            For ComAge = MinAge To MaxAge
                If TotalSubCNR <> 0 Then
                    PropSubPop(Stock, ComAge) = CNRShakers(Stock, ComAge) / TotalSubCNR
                Else
                    PropSubPop(Stock, ComAge) = 0
                End If
            Next ComAge
        Next Stock

    End Sub

    Sub CompShakers(ByVal Fish, ByVal TerminalType, ByVal EncounterRate)
        '**************************************************************************
        'Subroutine computes shaker mortality in each fishery by multiplying the
        ' base period encounter rate by the sublegal population and the release
        ' mortality rate.
        '**************************************************************************
        Dim LegalPopulation, SubLegalPopulation

        'COMPUTE SHAKER MORTALITY IF CATCH OCCURRED IN FISHERY
        ' New shaker method for chinook using SubLegal Encounter Rate
        ' 2/13/98
        'Replace SubLegal Encounter Rate (EncRate) Calculations for CNR Method-0 "Computed CNR"
        ' 1/29/2004 JFP

        EncounterRate = 0
        LegalPopulation = 0
        SubLegalPopulation = 0
        If Fish = 20 And TStep = 2 Then
            TStep = 2
        End If

        If TotalLandedCatch(Fish, TStep) > 0 Then
            
            If SizeLimitFix = True And MinSizeLimit(Fish, TStep) > ChinookBaseSizeLimit(Fish, TStep) Then
                SizeLimitFixShaker(Fish, TerminalType, EncounterRate)
            Else
                For Stk As Integer = 1 To NumStk
                    For Age As Integer = MinAge To MaxAge
                        If Fish = 63 And Age = 2 And TStep = 3 And Stk = 13 Then
                            Jim = 1
                        End If
                        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                        If TStep = 3 And Fish = 42 And Stk = 21 And Age = 3 Then 'used to break code
                            Stk = 21
                        End If
                        'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

                        Call CompLegProp(Stk, Age, Fish, TerminalType)
                        LegalPopulation = LegalPopulation + Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                        LegalPop = Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                        '- Zero Time 1 Yearling Shakers ...
                        '- Fish not Recruited Yet - Temp Fix 1/3/2000 JFP
                        If NumStk < 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
                            '- Regular Chinook FRAM
                            SubLegalPop = 0
                        ElseIf NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                            '- Selective Fishery Version
                            SubLegalPop = 0
                        Else
                            SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion
                        End If

                        SubLegalPopulation = SubLegalPopulation + SubLegalPop

                        '                    '-=======  NEW SIZE LIMIT Calcs  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        '                    '- Check if New Size Limit is different from Base Period Size Limit
                        '                    If MinSizeLimit(Fish, TStep) <> ChinookBaseSizeLimit(Fish, TStep) Then
                        '                        Dim BaseLegalPop, BaseSubLegalPop, BaseShakers, BaseCatch, NewShakers, NewCatch As Double
                        '                        Dim BaseSubEncounters, NewSubEncounters, SubEncDiff As Double
                        '                        ChinookBaseLegProp = True
                        '                        Call CompLegProp(Stk, Age, Fish, TerminalType)
                        '                        ChinookBaseLegProp = False
                        '                        BaseLegalPop = Cohort(Stk, Age, TerminalType, TStep) * BaseLegalProportion
                        '                        BaseSubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * BaseSubLegalProportion
                        '                        '- PS Yearling Fish not yet released or recruited to fishery
                        '                        If NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                        '                            BaseSubLegalPop = 0
                        '                        End If

                        '                        BaseCatch = _
                        '                           Cohort(Stk, Age, TerminalType, TStep) * _
                        '                           BaseExploitationRate(Stk, Age, Fish, TStep) * _
                        '                           FisheryScaler(Fish, TStep) * _
                        '                           StockFishRateScalers(Stk, Fish, TStep) * _
                        '                           BaseLegalProportion

                        '                        BaseSubEncounters = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                        '                        NewSubEncounters = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                        '                        BaseShakers = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                        '                        NewShakers = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                        '                        If MinSizeLimit(Fish, TStep) < ChinookBaseSizeLimit(Fish, TStep) Then

                        '                            SubEncDiff = (BaseSubEncounters - NewSubEncounters)

                        '                            '- Debug Print to check calculations
                        '                            If BaseShakers > 0 Or NewShakers > 0 Then
                        '                                PrnLine = String.Format("{0,3}{1,3}{2,3}{3,3}", Stk, Age, Fish, TStep)
                        '                                PrnLine &= String.Format("{0,10}", BaseShakers.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", NewShakers.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", BaseSubEncounters.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", NewSubEncounters.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", SubEncDiff.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", BaseCatch.ToString("####0.0000"))
                        '                                PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("####0.0000"))
                        '                                PrnLine &= " " & FisheryName(Fish)
                        '                                PrnLine &= " " & StockName(Stk)
                        '                                sw.WriteLine(PrnLine)
                        '                            End If

                        '                            '- Redo Total Fishery Arrays before New SizeLimit Calculations
                        '                            TotalEncounters(Fish, TStep) -= Encounters(Stk, Age, Fish, TStep)
                        '                            TotalLandedCatch(Fish, TStep) -= LandedCatch(Stk, Age, Fish, TStep)
                        '                            '- Only UnMarked (Wild) in Fisheries NumFish+1 to NumFish*2
                        '                            If (Stk Mod 2) <> 0 Then
                        '                                TotalLandedCatch(NumFish + Fish, TStep) -= LandedCatch(Stk, Age, Fish, TStep)
                        '                            End If

                        '                            '- When SizeLimit is less than Base SizeLimit use difference in BaseEncounters and NewEncounters
                        '                            '  to increase Landed Catch 
                        '                            LandedCatch(Stk, Age, Fish, TStep) = BaseCatch + SubEncDiff

                        '                            Encounters(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
                        '                            TotalEncounters(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)

                        '                            '- Normal Shaker Calculation
                        '                            Shakers(Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                        '                            '- Recalculate MSF Impacts after LandedCatch changes
                        '                            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Then
                        '                                '--- Use Selective Incidental Rate on ALL fish encountered
                        '                                TotalDropOff(Fish, TStep) -= DropOff(Stk, Age, Fish, TStep)
                        '                                DropOff(Stk, Age, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * LandedCatch(Stk, Age, Fish, TStep)
                        '                                TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + DropOff(Stk, Age, Fish, TStep)
                        '                                '- All Stocks in Marked/UnMarked pairs
                        '                                TotalLegalShakers(Fish, TStep) = TotalLegalShakers(Fish, TStep) - LegalShakers(Stk, Age, Fish, TStep)
                        '                                If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                        '                                    LegalShakers(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                        '                                    LandedCatch(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                        '                                    TotalLegalShakers(Fish, TStep) = TotalLegalShakers(Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep)
                        '                                Else           '--- UnMarked (Wild) in Selective
                        '                                    LegalShakers(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                        '                                    LandedCatch(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                        '                                    TotalLegalShakers(Fish, TStep) = TotalLegalShakers(Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep)
                        '                                End If
                        '                                TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                        '                                '- Only UnMarked (Wild) in Fisheries NumFish+1 to NumFish*2
                        '                                If (Stk Mod 2) <> 0 Then
                        '                                    TotalLandedCatch(NumFish + Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                        '                                End If
                        '                            Else
                        '                                '- Non-MSF Calculations (note: DropOff Done in IncMort Routine)
                        '                                TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                        '                                '- Only UnMarked (Wild) in Fisheries NumFish+1 to NumFish*2
                        '                                If (Stk Mod 2) <> 0 Then
                        '                                    TotalLandedCatch(NumFish + Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                        '                                End If
                        '                            End If
                        'SkipMSFAdj:
                        '                        Else
                        '                            '- When SizeLimit is greater than Base SizeLimit use difference in BaseCatch and New Catch for Shakers
                        '                            '  This NewCatch does not have the MSF Impacts
                        '                            NewCatch = _
                        '                               Cohort(Stk, Age, TerminalType, TStep) * _
                        '                               BaseExploitationRate(Stk, Age, Fish, TStep) * _
                        '                               FisheryScaler(Fish, TStep) * _
                        '                               StockFishRateScalers(Stk, Fish, TStep) * _
                        '                               LegalProportion
                        '                            Shakers(Stk, Age, Fish, TStep) = BaseShakers + (BaseCatch - NewCatch) * ShakerMortRate(Fish, TStep)
                        '                        End If
                        '                    Else
                        '                        '- Normal Shaker Calculation
                        '                        Shakers(Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                        '                    End If


                        '                    '-=======  END of NEW SIZE LIMIT Calcs  ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                        'If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                        '   BYLandedCatch(BY, Stk, BYAge, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, BYAge, Fish, TStep) * BYCohort(BY, Stk, BYAge, TerminalType, TStep) * FisheryScaler(Fish, TStep) * LegalProportion
                        'End If
                        ''- MSF Fishery Scaler & Quota
                        'If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                        '- Retention Fishery Shaker Calculation
                        If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                            Shakers(Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                            TotalShakers(Fish, TStep) = TotalShakers(Fish, TStep) + Shakers(Stk, Age, Fish, TStep)
                        End If

                        '- MSF Shaker Calculation
                        If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                            MSFShakers(Stk, Age, Fish, TStep) = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                            TotalShakers(Fish, TStep) = TotalShakers(Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        End If
                    Next Age
                Next Stk

                If LegalPopulation > 0 Then
                    EncounterRate = SubLegalPopulation / LegalPopulation
                Else
                    EncounterRate = 0
                End If

                ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013
                'If SizeLimitScenario = True Then
                '   'Only invoke the subroutine if external shaker inputs are provided and only for Puget Sound Sport at this stage
                '   'Ditch the select case business if this is something that's desired outside of Puget Sound
                '   Select Case Fish
                '      Case 36, 42, 45, 53, 54, 56, 57, 64, 67 'Ignore 8D, 10A, and 10E (48,60,62) given assumption of zero base sublegals
                '         If ShakerFlagNS(Fish, TStep) > 1 Or ShakerFlagMSF(Fish, TStep) > 1 Then
                '            Call CompExternalChinookShakers(TerminalType, Fish)
                '         End If
                '   End Select
                'End If
                ''############################# BEGIN NEW CODE ############################ Pete-Jan. 2013


                ''- Check if MSF Quota with New LandedCatch equals Target Quota
                'If MinSizeLimit(Fish, TStep) < ChinookBaseSizeLimit(Fish, TStep) And FisheryFlag(Fish, TStep) = 8 Then
                '   If FisheryQuota(Fish, TStep) <> (TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish)) Then
                '      Dim QuotaScaler As Double
                '      '- Adjust Fishery LandedCatch & IncMort to Quota Value with New FisheryScaler Vvalue
                '      FisheryScaler(Fish, TStep) = FisheryScaler(Fish, TStep) * (FisheryQuota(Fish, TStep) / (TotalLandedCatch(Fish, TStep) * ModelStockProportion(Fish)))
                '      QuotaScaler = FisheryQuota(Fish, TStep) / (TotalLandedCatch(Fish, TStep) * ModelStockProportion(Fish))
                '      TotalLandedCatch(Fish, TStep) = 0
                '      TotalLandedCatch(NumFish + Fish, TStep) = 0
                '      For Stk = 1 To NumStk
                '         For Age = MinAge To MaxAge
                '            LandedCatch(Stk, Age, Fish, TStep) = QuotaScaler * LandedCatch(Stk, Age, Fish, TStep)
                '            '--- Adjust CNR and Shakers to Quota Scalar
                '            TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) - DropOff(Stk, Age, Fish, TStep)
                '            DropOff(Stk, Age, Fish, TStep) = DropOff(Stk, Age, Fish, TStep) * QuotaScaler
                '            TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + DropOff(Stk, Age, Fish, TStep)
                '            TotalLegalShakers(Fish, TStep) = TotalLegalShakers(Fish, TStep) - LegalShakers(Stk, Age, Fish, TStep)
                '            LegalShakers(Stk, Age, Fish, TStep) = LegalShakers(Stk, Age, Fish, TStep) * QuotaScaler
                '            TotalLegalShakers(Fish, TStep) = TotalLegalShakers(Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep)
                '            TotalEncounters(Fish, TStep) = TotalEncounters(Fish, TStep) - Encounters(Stk, Age, Fish, TStep)
                '            Encounters(Stk, Age, Fish, TStep) = Encounters(Stk, Age, Fish, TStep) * QuotaScaler
                '            TotalEncounters(Fish, TStep) = TotalEncounters(Fish, TStep) + Encounters(Stk, Age, Fish, TStep)
                '            TotalLandedCatch(Fish, TStep) = TotalLandedCatch(Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                '            If (Stk Mod 2) <> 0 Then      '--- UNMarked Fish
                '               TotalLandedCatch(NumFish + Fish, TStep) = TotalLandedCatch(NumFish + Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep)
                '            End If

                '            '- DEBUG Code to Check CompCatch Calculations
                '            'If SkipJim = 1 And Fish = 39 And TStep = 3 And LandedCatch(Stk, Age, Fish, TStep) <> 0 Then
                '            '   PrnLine = String.Format("{0,3}{1,4}{2,4}{3,4}", Stk.ToString, Age.ToString, Fish.ToString, TStep.ToString)
                '            '   PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
                '            '   PrnLine &= String.Format("{0,10}", Cohort(Stk, Age, TerminalType, TStep).ToString("#######0"))
                '            '   PrnLine &= String.Format("{0,11}", BaseExploitationRate(Stk, Age, Fish, TStep).ToString("0.00000000"))
                '            '   PrnLine &= String.Format("{0,8}", FisheryScaler(Fish, TStep).ToString("##0.000"))
                '            '   PrnLine &= String.Format("{0,8}", StockFishRateScalers(Stk, Fish, TStep).ToString("##0.000"))
                '            '   PrnLine &= String.Format("{0,8}", LegalProportion.ToString("##0.000"))
                '            '   sw.WriteLine(PrnLine)
                '            'End If

                '         Next Age
                '      Next Stk

                '   End If
                'End If

            End If
        End If
    End Sub
    Sub SizeLimitFixLanded(ByVal Fish As Integer, ByVal TerminalType As Integer)
        Dim BaseLegalPop, BaseSubLegalPop, BaseShakers, BaseCatch, NewShakers As Double
        Dim BaseSubEncounters, NewSubEncounters, SubEncDiff As Double
        'Dim NSFQuotaTotal(NumFish, TStep), MSFQuotaTotal(NumFish, TStep) As Double

        For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
                If Fish = 22 And TStep = 2 And Stk = 47 And Age = 3 Then
                    TStep = 2
                End If

                '- Zero Calculation Arrays for TAMM Iteration Calculations
                LandedCatch(Stk, Age, Fish, TStep) = 0
                DropOff(Stk, Age, Fish, TStep) = 0
                Encounters(Stk, Age, Fish, TStep) = 0
                NonRetention(Stk, Age, Fish, TStep) = 0
                MSFLandedCatch(Stk, Age, Fish, TStep) = 0
                MSFDropOff(Stk, Age, Fish, TStep) = 0
                MSFEncounters(Stk, Age, Fish, TStep) = 0
                MSFNonRetention(Stk, Age, Fish, TStep) = 0


                ChinookBaseLegProp = False  'computes legal prop at modeled size limit
                Call CompLegProp(Stk, Age, Fish, TerminalType)
                ChinookBaseLegProp = True 'computes legal prop at BP size limit
                Call CompLegProp(Stk, Age, Fish, TerminalType)
                ChinookBaseLegProp = False

                ' deal with cases where there is a quota and a quota flag but scalar = 0, first pass scalar = 1
                ' deal with cases where there is a quota flag but quota = 0

                If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                    If FisheryQuota(Fish, TStep) > 0 Then
                        FisheryScaler(Fish, TStep) = 1
                    Else
                        FisheryScaler(Fish, TStep) = 0
                    End If
                End If

                If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                    If MSFFisheryQuota(Fish, TStep) > 0 Then
                        MSFFisheryScaler(Fish, TStep) = 1
                    Else
                        MSFFisheryScaler(Fish, TStep) = 0
                    End If
                End If

                BaseCatch = _
                  Cohort(Stk, Age, TerminalType, TStep) * _
                  BaseExploitationRate(Stk, Age, Fish, TStep) * _
                  FisheryScaler(Fish, TStep) * _
                  StockFishRateScalers(Stk, Fish, TStep) * _
                  BaseLegalProportion

                MSFBaseLegalEncounters = _
                   Cohort(Stk, Age, TerminalType, TStep) * _
                   BaseExploitationRate(Stk, Age, Fish, TStep) * _
                   MSFFisheryScaler(Fish, TStep) * _
                   StockFishRateScalers(Stk, Fish, TStep) * _
                   BaseLegalProportion

                

                BaseSubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * BaseSubLegalProportion
                SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion

                BaseSubEncounters = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                NewSubEncounters = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                BaseShakers = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                NewShakers = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                MSFBaseSubEncounters = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                MSFNewSubEncounters = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                MSFBaseShakers = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                MSFNewShakers = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                Select Case FisheryFlag(Fish, TStep)
                    Case 1, 17, 18, 2, 27, 28
                        LandedCatch(Stk, Age, Fish, TStep) = BaseCatch + (BaseSubEncounters - NewSubEncounters)
                        TotalEncounters(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                End Select

                '- MSF Fishery Scaler
                Select Case FisheryFlag(Fish, TStep)
                    Case 7, 17, 27, 8, 18, 28
                        MSFLandedCatch(Stk, Age, Fish, TStep) = MSFBaseLegalEncounters + (MSFBaseSubEncounters - MSFNewSubEncounters)
                        MSFEncounters(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep)
                        MSFDropOff(Stk, Age, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * MSFLandedCatch(Stk, Age, Fish, TStep)
                        TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        TotalEncounters(Fish, TStep) += Encounters(Stk, Age, Fish, TStep) + MSFEncounters(Stk, Age, Fish, TStep)
                        '- All Stocks in Marked/UnMarked pairs
                        If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                            MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                            MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                            TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                        Else           '--- UnMarked (Wild) in Selective
                            MSFNonRetention(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                            MSFLandedCatch(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                            TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                        End If
                End Select

                Encounters(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
                'TotalEncounters(Fish, TStep) += Encounters(Stk, Age, Fish, TStep) + MSFEncounters(Stk, Age, Fish, TStep)
                TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                'TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                MSFQuotaTotal(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                NSFQuotaTotal(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
                If (Stk Mod 2) <> 0 Then
                    TotalLandedCatch(NumFish + Fish, TStep) = TotalLandedCatch(NumFish + Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                End If

            Next Age
        Next Stk
        

    End Sub
    Sub SizeLimitFixShaker(ByVal Fish, ByVal TerminalType, ByVal EncounterRate)
        '**************************************************************************
        'Subroutine computes shaker mortality in a size limit corrected run (hold total fisheries encounters constant)
        'in each fishery by multiplying the
        ' base period encounter rate by the sublegal population and the release
        ' mortality rate.
        '**************************************************************************

        Dim LegalPopulation, SubLegalPopulation
        Dim NewCatch As Double

        EncounterRate = 0
        LegalPopulation = 0
        SubLegalPopulation = 0
        NewCatch = 0

        For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge

                Call CompLegProp(Stk, Age, Fish, TerminalType)

                LegalPopulation = LegalPopulation + Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                LegalPop = Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
                '- Zero Time 1 Yearling Shakers ...
                '- Fish not Recruited Yet - Temp Fix 1/3/2000 JFP
                If NumStk < 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
                    '- Regular Chinook FRAM
                    SubLegalPop = 0
                ElseIf NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                    '- Selective Fishery Version
                    SubLegalPop = 0
                Else
                    SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion
                End If

                SubLegalPopulation = SubLegalPopulation + SubLegalPop
                '- Normal Shaker Calculation
                'TotalEncounters(Fish, TStep) -= Encounters(Stk, Age, Fish, TStep) - MSFEncounters(Stk, Age, Fish, TStep)


                'If TStep = 2 And Fish = 22 And Stk = 2 And Age = 4 Then
                '    TStep = 2
                'End If

                ChinookBaseLegProp = True
                Call CompLegProp(Stk, Age, Fish, TerminalType)
                ChinookBaseLegProp = False

                Dim BaseLegalPop, BaseSubLegalPop, BaseShakers, BaseCatch, NewShakers As Double
                Dim BaseSubEncounters, NewSubEncounters, SubEncDiff As Double

                'ChinookBaseLegProp = True

                'Call CompLegProp(Stk, Age, Fish, TerminalType)
                'ChinookBaseLegProp = False
                BaseLegalPop = Cohort(Stk, Age, TerminalType, TStep) * BaseLegalProportion
                BaseSubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * BaseSubLegalProportion
                '- PS Yearling Fish not yet released or recruited to fishery
                

                BaseCatch = _
                   Cohort(Stk, Age, TerminalType, TStep) * _
                   BaseExploitationRate(Stk, Age, Fish, TStep) * _
                   FisheryScaler(Fish, TStep) * _
                   StockFishRateScalers(Stk, Fish, TStep) * _
                   BaseLegalProportion

                MSFBaseLegalEncounters = _
                   Cohort(Stk, Age, TerminalType, TStep) * _
                   BaseExploitationRate(Stk, Age, Fish, TStep) * _
                   MSFFisheryScaler(Fish, TStep) * _
                   StockFishRateScalers(Stk, Fish, TStep) * _
                   BaseLegalProportion

                NewCatch = _
                                Cohort(Stk, Age, TerminalType, TStep) * _
                                BaseExploitationRate(Stk, Age, Fish, TStep) * _
                                FisheryScaler(Fish, TStep) * _
                                StockFishRateScalers(Stk, Fish, TStep) * _
                                LegalProportion

                MSFNewCatch = _
                                Cohort(Stk, Age, TerminalType, TStep) * _
                                BaseExploitationRate(Stk, Age, Fish, TStep) * _
                                MSFFisheryScaler(Fish, TStep) * _
                                StockFishRateScalers(Stk, Fish, TStep) * _
                                LegalProportion

                BaseSubEncounters = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                NewSubEncounters = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                BaseShakers = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                NewShakers = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                MSFBaseSubEncounters = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                MSFNewSubEncounters = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

                MSFBaseShakers = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                MSFNewShakers = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                '- Redo Total Fishery Arrays before New SizeLimit Calculations

                'TotalShakers(Fish, TStep) -= Shakers(Stk, Age, Fish, TStep) - MSFShakers(Stk, Age, Fish, TStep)

                Select Case FisheryFlag(Fish, TStep)
                    Case 1, 17, 18, 2, 27, 28
                        Shakers(Stk, Age, Fish, TStep) = BaseShakers + (BaseCatch - NewCatch) * ShakerMortRate(Fish, TStep)
                End Select
                Select Case FisheryFlag(Fish, TStep)
                    Case 7, 8, 17, 18, 27, 28
                        MSFShakers(Stk, Age, Fish, TStep) = MSFBaseShakers + (MSFBaseLegalEncounters - MSFNewCatch) * ShakerMortRate(Fish, TStep)
                End Select

                TotalShakers(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)


            Next Age
        Next Stk

        
        If LegalPopulation > 0 Then
            EncounterRate = SubLegalPopulation / LegalPopulation
        Else
            EncounterRate = 0
        End If

    End Sub
    Sub BYSizeLimitFixLanded(ByVal Fish As Integer, ByVal TerminalType As Integer)
        Dim BaseLegalPop, BaseSubLegalPop, BaseShakers, BaseCatch, NewShakers As Double
        Dim BaseSubEncounters, NewSubEncounters, SubEncDiff As Double
        'Dim NSFQuotaTotal(NumFish, TStep), MSFQuotaTotal(NumFish, TStep) As Double

        'For Stk As Integer = 1 To NumStk
        '    For Age As Integer = MinAge To MaxAge
        '        If Fish = 22 And TStep = 2 And Stk = 47 And Age = 3 Then
        '            TStep = 2
        '        End If

        '        '- Zero Calculation Arrays for TAMM Iteration Calculations
        '        LandedCatch(Stk, Age, Fish, TStep) = 0
        '        DropOff(Stk, Age, Fish, TStep) = 0
        '        Encounters(Stk, Age, Fish, TStep) = 0
        '        NonRetention(Stk, Age, Fish, TStep) = 0
        '        MSFLandedCatch(Stk, Age, Fish, TStep) = 0
        '        MSFDropOff(Stk, Age, Fish, TStep) = 0
        '        MSFEncounters(Stk, Age, Fish, TStep) = 0
        '        MSFNonRetention(Stk, Age, Fish, TStep) = 0


        '        ChinookBaseLegProp = False  'computes legal prop at modeled size limit
        '        Call CompLegProp(Stk, Age, Fish, TerminalType)
        ChinookBaseLegProp = True 'computes legal prop at BP size limit
        Call CompLegProp(Stk, BYAge, Fish, TerminalType)
        ChinookBaseLegProp = False

        ' deal with cases where there is a quota and a quota flag but scalar = 0, first pass scalar = 1
        ' deal with cases where there is a quota flag but quota = 0

        'If FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
        '    If FisheryQuota(Fish, TStep) > 0 Then
        '        FisheryScaler(Fish, TStep) = 1
        '    Else
        '        FisheryScaler(Fish, TStep) = 0
        '    End If
        'End If

        'If FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
        '    If MSFFisheryQuota(Fish, TStep) > 0 Then
        '        MSFFisheryScaler(Fish, TStep) = 1
        '    Else
        '        MSFFisheryScaler(Fish, TStep) = 0
        '    End If
        'End If

        BaseCatch = _
          BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
          BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
          FisheryScaler(Fish, TStep) * _
          StockFishRateScalers(Stk, Fish, TStep) * _
          BaseLegalProportion

        MSFBaseLegalEncounters = _
           BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
           BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
           MSFFisheryScaler(Fish, TStep) * _
           StockFishRateScalers(Stk, Fish, TStep) * _
           BaseLegalProportion



        BaseSubLegalPop = BYCohort(BY, Stk, BYAge, TerminalType, TStep) * BaseSubLegalProportion
        SubLegalPop = BYCohort(BY, Stk, BYAge, TerminalType, TStep) * SubLegalProportion

        BaseSubEncounters = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        NewSubEncounters = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        BaseShakers = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        NewShakers = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        MSFBaseSubEncounters = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        MSFNewSubEncounters = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        MSFBaseShakers = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        MSFNewShakers = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        If Fish = 17 And Stk = 26 And TStep = 2 And BYAge = 4 Then
            TStep = 2
        End If



        Select Case FisheryFlag(Fish, TStep)
            Case 1, 17, 18, 2, 27, 28
                BYLandedCatch(BY, Stk, BYAge, Fish, TStep) = BaseCatch + (BaseSubEncounters - NewSubEncounters)
                'TotalEncounters(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
        End Select

        '- MSF Fishery Scaler
        Select Case FisheryFlag(Fish, TStep)
            Case 7, 17, 27, 8, 18, 28
                BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = MSFBaseLegalEncounters + (MSFBaseSubEncounters - MSFNewSubEncounters)
                'MSFEncounters(Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep)
                BYMSFDropOff(BY, Stk, BYAge, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep)
                'TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                'TotalEncounters(Fish, TStep) += Encounters(Stk, Age, Fish, TStep) + MSFEncounters(Stk, Age, Fish, TStep)
                '- All Stocks in Marked/UnMarked pairs
                If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                    BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                    BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                    'TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                Else           '--- UnMarked (Wild) in Selective
                    BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                    BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                    'TotalNonRetention(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                End If
        End Select

        'Encounters(Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
        'TotalEncounters(Fish, TStep) += Encounters(Stk, Age, Fish, TStep) + MSFEncounters(Stk, Age, Fish, TStep)
        'TotalLandedCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
        'TotalDropOff(Fish, TStep) = TotalDropOff(Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
        'MSFQuotaTotal(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
        'NSFQuotaTotal(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep)
        'If (Stk Mod 2) <> 0 Then
        '    TotalLandedCatch(NumFish + Fish, TStep) = TotalLandedCatch(NumFish + Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
        'End If

        'Next Age
        'Next Stk
    End Sub
    Sub BYSizeLimitFixShaker(ByVal Fish, ByVal TerminalType, ByVal EncounterRate)
        '**************************************************************************
        ''Subroutine computes shaker mortality in a size limit corrected run (hold total fisheries encounters constant)
        ''in each fishery by multiplying the
        '' base period encounter rate by the sublegal population and the release
        '' mortality rate.
        '**************************************************************************

        'Dim LegalPopulation, SubLegalPopulation
        Dim NewCatch As Double

        'EncounterRate = 0
        'LegalPopulation = 0
        'SubLegalPopulation = 0
        NewCatch = 0

        'LegalPopulation = LegalPopulation + Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
        'LegalPop = Cohort(Stk, Age, TerminalType, TStep) * LegalProportion
        ''- Zero Time 1 Yearling Shakers ...
        ''- Fish not Recruited Yet - Temp Fix 1/3/2000 JFP
        'If NumStk < 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
        '    '- Regular Chinook FRAM
        '    SubLegalPop = 0
        'ElseIf NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
        '    '- Selective Fishery Version
        '    SubLegalPop = 0
        'Else
        '    SubLegalPop = Cohort(Stk, Age, TerminalType, TStep) * SubLegalProportion
        'End If

        'SubLegalPopulation = SubLegalPopulation + SubLegalPop
        ''- Normal Shaker Calculation
        ''TotalEncounters(Fish, TStep) -= Encounters(Stk, Age, Fish, TStep) - MSFEncounters(Stk, Age, Fish, TStep)


        'If TStep = 2 And Fish = 22 And Stk = 2 And Age = 4 Then
        '    TStep = 2
        'End If

        ChinookBaseLegProp = True
        Call CompLegProp(Stk, BYAge, Fish, TerminalType)
        ChinookBaseLegProp = False

        Dim BaseLegalPop, BaseSubLegalPop, BaseShakers, BaseCatch, NewShakers As Double
        Dim BaseSubEncounters, NewSubEncounters, SubEncDiff As Double

        
        BaseLegalPop = BYCohort(BY, Stk, BYAge, TerminalType, TStep) * BaseLegalProportion
        BaseSubLegalPop = BYCohort(BY, Stk, BYAge, TerminalType, TStep) * BaseSubLegalProportion
        '- PS Yearling Fish not yet released or recruited to fishery


        BaseCatch = _
           BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
           BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
           FisheryScaler(Fish, TStep) * _
           StockFishRateScalers(Stk, Fish, TStep) * _
           BaseLegalProportion

        MSFBaseLegalEncounters = _
           BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
           BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
           MSFFisheryScaler(Fish, TStep) * _
           StockFishRateScalers(Stk, Fish, TStep) * _
           BaseLegalProportion

        NewCatch = _
                        BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
                        BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
                        FisheryScaler(Fish, TStep) * _
                        StockFishRateScalers(Stk, Fish, TStep) * _
                        LegalProportion

        MSFNewCatch = _
                        BYCohort(BY, Stk, BYAge, TerminalType, TStep) * _
                        BaseExploitationRate(Stk, BYAge, Fish, TStep) * _
                        MSFFisheryScaler(Fish, TStep) * _
                        StockFishRateScalers(Stk, Fish, TStep) * _
                        LegalProportion

        BaseSubEncounters = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        NewSubEncounters = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        BaseShakers = FisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        NewShakers = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        MSFBaseSubEncounters = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        MSFNewSubEncounters = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)

        MSFBaseShakers = MSFFisheryScaler(Fish, TStep) * BaseSubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        MSFNewShakers = MSFFisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
        '- Redo Total Fishery Arrays before New SizeLimit Calculations

        'TotalShakers(Fish, TStep) -= Shakers(Stk, Age, Fish, TStep) - MSFShakers(Stk, Age, Fish, TStep)

        Select Case FisheryFlag(Fish, TStep)
            Case 1, 17, 18, 2, 27, 28
                BYShakers(BY, Stk, BYAge, Fish, TStep) = BaseShakers + (BaseCatch - NewCatch) * ShakerMortRate(Fish, TStep)
        End Select
        Select Case FisheryFlag(Fish, TStep)
            Case 7, 8, 17, 18, 27, 28
                BYMSFShakers(BY, Stk, BYAge, Fish, TStep) = MSFBaseShakers + (MSFBaseLegalEncounters - MSFNewCatch) * ShakerMortRate(Fish, TStep)
        End Select

        'TotalShakers(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)

        'If LegalPopulation > 0 Then
        '    EncounterRate = SubLegalPopulation / LegalPopulation
        'Else
        '    EncounterRate = 0
        'End If

    End Sub
    Sub IncMort(ByVal TerminalType)
        Dim EncRate As Double

        '**************************************************************************
        'Subroutine calls other subroutines to compute incidental mortality
        ' (shaker, CNR, and other nonlanded catch mortality).
        '**************************************************************************

        For Fish As Integer = 1 To NumFish
            If Fish = 20 Then
                Fish = 20
            End If
            ReDim PropSubPop(NumStk, MaxAge)
            If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
                If (FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8) And SpeciesName = "COHO" Then GoTo SelctFsh
                '- No sub-legal shaker calculations for COHO
                If SpeciesName = "CHINOOK" Then
                    Call CompShakers(Fish, TerminalType, EncRate)
                End If
                '--- Dropoff Calculated for Selective Fisheries in CompCatch

                If FisheryFlag(Fish, TStep) < 6 Or FisheryFlag(Fish, TStep) > 10 Then
                    Call CompOthMort(Fish)
                End If
SelctFsh:
                '- Non-Retention
                If NonRetentionFlag(Fish, TStep) <> 0 Then
                    If SpeciesName = "COHO" Then
                        Call CompCohoCNR(Fish, TerminalType)
                    Else
                        Call CompCNR(Fish, TerminalType, EncRate)
                    End If
                End If
            End If

        Next Fish

    End Sub

    Sub Mature()
        '**************************************************************************
        'Subroutine computes mature and immature components of cohort by
        ' multiplying the total cohort by the maturation rate.
        '**************************************************************************

        'COMPUTE COHORT SIZE AT END OF PRETERMINAL FISHING
        ' BY SUBTRACTING FISHING MORTALITY

        Dim JimD As Double

      For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
                If Stk = 19 And Age = 3 And TStep = 3 Then
                    Jim = 1
                End If

                For Fish As Integer = 1 To NumFish
                    If Fish = 198 Then
                        Jim = 1
                    End If
                    If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                        Cohort(Stk, Age, PTerm, TStep) = Cohort(Stk, Age, PTerm, TStep) - LandedCatch(Stk, Age, Fish, TStep) - Shakers(Stk, Age, Fish, TStep) - NonRetention(Stk, Age, Fish, TStep) - DropOff(Stk, Age, Fish, TStep) _
                                                                   - MSFLandedCatch(Stk, Age, Fish, TStep) - MSFShakers(Stk, Age, Fish, TStep) - MSFNonRetention(Stk, Age, Fish, TStep) - MSFDropOff(Stk, Age, Fish, TStep)

                        If Double.IsNaN(Cohort(Stk, Age, PTerm, TStep)) Then
                            MsgBox("Invalid Cohort Size for Stk " & Stk & ", PTerm " & PTerm & ", Time Step " & TStep & ".")
                        End If

                        'If TStep = 4 Then
                        JimD = LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) _
                                                                + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        'If JimD <> 0 Then
                        '   Jim = 1
                        '   PrnLine = String.Format("Mature-{0,3}{1,4}{2,4}{3,4}", Stk.ToString, Age.ToString, Fish.ToString, TStep.ToString)
                        '   PrnLine &= String.Format("{0,10}", Cohort(Stk, Age, PTerm, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", Shakers(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", NonRetention(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", DropOff(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", MSFLandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", MSFShakers(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", MSFNonRetention(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   PrnLine &= String.Format("{0,10}", MSFDropOff(Stk, Age, Fish, TStep).ToString("#####0.00"))
                        '   sw.WriteLine(PrnLine)
                        'End If
                        'End If

                    End If
                Next Fish
                '- Save After-PreTerminal Cohort
                Cohort(Stk, Age, 2, TStep) = Cohort(Stk, Age, PTerm, TStep)
                '- Check for Negative Cohort Size
                If Cohort(Stk, Age, PTerm, TStep) < 0.0 Then
                    Jim = 1
                    'Print #10, "ERROR - Negative Cohort Size for STOCK, AGE = " + SmlStockName$(Stk) + " " + Str(age)
                    'MsgBox("ERROR - Negative Cohort Size for STOCK, AGE = " & StockName(Stk).ToString & " " & Age.ToString, MsgBoxStyle.OkOnly)
                End If
            Next Age
      Next Stk

        'COMPUTE MATURE COMPONENT OF EACH COHORT
      For Stk As Integer = 1 To NumStk
         For Age As Integer = MinAge To MaxAge
            Cohort(Stk, Age, Term, TStep) = Cohort(Stk, Age, PTerm, TStep) * MaturationRate(Stk, Age, TStep)
            Cohort(Stk, Age, PTerm, TStep) = Cohort(Stk, Age, PTerm, TStep) - Cohort(Stk, Age, Term, TStep)
         Next Age
      Next Stk

    End Sub

    Sub NatMort()
        '**************************************************************************
        'Subroutine computes cohort abundance after natural mortality.
        '**************************************************************************

        'COMPUTE ABUNDANCE AFTER NATURAL MORTALITY
      For Stk As Integer = 1 To NumStk
         For Age As Integer = MinAge To MaxAge
            If TStep = 4 And SpeciesName = "COHO" And Age = 3 And RunTAMMIter = 1 And TammIteration = 0 And NumSteps = 5 Then
               CohoTime4Cohort(Stk) = Cohort(Stk, Age, PTerm, TStep)
            End If
            '- Save Pre-Natural-Mortality Cohort
            Cohort(Stk, Age, 4, TStep) = Cohort(Stk, Age, PTerm, TStep)
            '- Subtract Natural Mortality
            Cohort(Stk, Age, PTerm, TStep) = Cohort(Stk, Age, PTerm, TStep) * (1 - NaturalMortality(Age, TStep))
            '- Save Working PreTerminal Cohort size
            Cohort(Stk, Age, 3, TStep) = Cohort(Stk, Age, PTerm, TStep)
         Next Age
      Next Stk

    End Sub

    Function NormlDistr(ByVal MinSize As Integer, ByVal MeanSize As Double, ByVal SizeStdDev As Double)
        'Obtained from WDF/NBS program, Subroutine NORMFN.
        Dim Z, ABSZ, A1, A2, A3

        Z = (MinSize - MeanSize) / SizeStdDev
        ABSZ = Abs(Z)
        A1 = (ABSZ * (0.000005383 * ABSZ + 0.0000488906) + 0.0000380036)
        A2 = (ABSZ * (ABSZ * A1 + 0.0032776263) + 0.0211410061)
        A3 = 1 / (1 + ABSZ * (ABSZ * A2 + 0.049867347))
        A3 = 1 - 0.5 * A3 ^ 16
        If Z < 0 Then
            A3 = 1 - A3
        End If

        NormlDistr = A3

    End Function

    Sub SaveDat()

        '- Microsoft Access DataBase Table Updates

        Dim CmdStr As String
        Dim TimeDiff1, TimeStep As Integer
        Dim MortSum As Double
        Dim StartTime, EndTime As DateTime
        Dim myPoint As Point = FVS_RunModel.RunProgressLabel.Location

        '- Label Update
        FVS_RunModel.RunProgressLabel.Text = " Saving MORTALITY Table to Database ... Please Wait "
        myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
        FVS_RunModel.RunProgressLabel.Location = myPoint
        FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
        FVS_RunModel.RunProgressLabel.Refresh()
        '- Progress Bar
        FVS_RunModel.MRProgressBar.Minimum = 1
        'FVS_ModelRun.MRProgressBar.Maximum = NumStk * MaxAge * NumSteps * NumFish
        FVS_RunModel.MRProgressBar.Maximum = NumStk
        FVS_RunModel.MRProgressBar.Step = 1
        FVS_RunModel.MRProgressBar.Value = 1
        FVS_RunModel.BringToFront()
        FVS_RunModel.Refresh()

        Dim FishMortDA As New System.Data.OleDb.OleDbDataAdapter
        Dim NonRetEnc As String
        '*************************************Produce Output of NR Encounters************************************
        NonRetEnc = FVSdatabasepath & "\" & RunIDYearSelect & "NonRetention.txt"
        FileOpen(53, NonRetEnc, OpenMode.Output)
        Print(53, "Nonretention Encounters by legal(1) and sublegal (2) divided by model stock proportion" & vbCrLf)
        Print(53, "Year" & "," & "Tstep" & "," & "Stk" & "," & "Fish" & "," & "Age" & "," & "SizeStatus" & "," & "#Encounters" & vbCrLf)


        CmdStr = "SELECT * FROM Mortality WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, FisheryID, TimeStep"
        Dim FMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        FishMortDA.SelectCommand = FMcm

        CmdStr = "DELETE * FROM Mortality WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim FMDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        FishMortDA.DeleteCommand = FMDcm

        Dim FMcb As New OleDb.OleDbCommandBuilder
        FMcb = New OleDb.OleDbCommandBuilder(FishMortDA)

        StartTime = DateTime.Now

        '- MORTALITY DataBase Table Save --------
        Dim FramTrans As OleDb.OleDbTransaction
        Dim FIC As New OleDbCommand
        Dim RCount As Integer
        FramDB.Open()
        FishMortDA.DeleteCommand.ExecuteScalar()
        FramTrans = FramDB.BeginTransaction
        FIC.Connection = FramDB
        FIC.Transaction = FramTrans
        RCount = 0
        For Stk = 1 To NumStk
            FVS_RunModel.MRProgressBar.PerformStep()
            FVS_RunModel.BringToFront()
            FVS_RunModel.Refresh()
            FVS_RunModel.MRProgressBar.Refresh()
         For Age As Integer = 1 To MaxAge
            For Fish As Integer = 1 To NumFish
               For TimeStep = 1 To NumSteps
                  MortSum = LandedCatch(Stk, Age, Fish, TimeStep) + NonRetention(Stk, Age, Fish, TimeStep) + Shakers(Stk, Age, Fish, TimeStep) + DropOff(Stk, Age, Fish, TimeStep) + MSFLandedCatch(Stk, Age, Fish, TimeStep) + MSFNonRetention(Stk, Age, Fish, TimeStep) + MSFShakers(Stk, Age, Fish, TimeStep) + MSFDropOff(Stk, Age, Fish, TimeStep)
                  If MortSum <> 0 Then
                     RCount += 1
                     FIC.CommandText = "INSERT INTO Mortality (PrimaryKey,RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                     "VALUES(" & RCount.ToString & "," & _
                     RunIDSelect.ToString & "," & _
                     Stk.ToString & "," & _
                     Age.ToString & "," & _
                     Fish.ToString & "," & _
                     TimeStep.ToString & "," & _
                     LandedCatch(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     NonRetention(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     Shakers(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     DropOff(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     Encounters(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     MSFLandedCatch(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     MSFNonRetention(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     MSFShakers(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     MSFDropOff(Stk, Age, Fish, TimeStep).ToString("######0.000000") & "," & _
                     MSFEncounters(Stk, Age, Fish, TimeStep).ToString("######0.000000") & ")"
                            FIC.ExecuteNonQuery()
                            If NRLegal(1, Stk, Age, Fish, TimeStep) > 0 Then
                                Print(53, RunIDYearSelect & "," & TimeStep & "," & Stk & "," & Fish & "," & Age & "," & 1 & "," & NRLegal(1, Stk, Age, Fish, TimeStep) & vbCrLf)
                            End If
                            If NRLegal(2, Stk, Age, Fish, TimeStep) > 0 Then
                                Print(53, RunIDYearSelect & "," & TimeStep & "," & Stk & "," & Fish & "," & Age & "," & 2 & "," & NRLegal(2, Stk, Age, Fish, TimeStep) & vbCrLf)
                            End If
                        End If
                    Next
            Next
         Next
        Next
        FileClose(53)
        FramTrans.Commit()
        FramDB.Close()

        FishMortDA = Nothing
        FVS_RunModel.MRProgressBar.Visible = False

        '- COHORT DataBase Table Save --------
        FVS_RunModel.RunProgressLabel.Text = " Saving COHORT Table to Database ... Please Wait "
        myPoint = FVS_RunModel.RunProgressLabel.Location
        myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
        FVS_RunModel.RunProgressLabel.Location = myPoint
        FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
        FVS_RunModel.RunProgressLabel.Refresh()
        FVS_RunModel.Refresh()

        CmdStr = "SELECT * FROM Cohort WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, TimeStep"
        Dim COHcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim CohortDA As New System.Data.OleDb.OleDbDataAdapter
        CohortDA.SelectCommand = COHcm

        CmdStr = "DELETE * FROM Cohort WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim COHDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        CohortDA.DeleteCommand = COHDcm

        Dim COHcb As New OleDb.OleDbCommandBuilder
        COHcb = New OleDb.OleDbCommandBuilder(CohortDA)
        'CohortDA.Fill(FramDataSet, "CohortSizes")

        Dim CohTrans As OleDb.OleDbTransaction
        Dim FCC As New OleDbCommand

        FramDB.Open()
        CohortDA.DeleteCommand.ExecuteScalar()
        CohTrans = FramDB.BeginTransaction
        FCC.Connection = FramDB
        FCC.Transaction = CohTrans

      For Stk As Integer = 1 To NumStk
         For Age As Integer = MinAge To MaxAge
            For TimeStep = 1 To NumSteps
               If Cohort(Stk, Age, 3, TimeStep) <> 0 Or Cohort(Stk, Age, 1, TimeStep) <> 0 Then
                  FCC.CommandText = "INSERT INTO Cohort (RunID,StockID,Age,TimeStep,Cohort,MatureCohort,StartCohort,WorkingCohort,MidCohort) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Stk.ToString & "," & _
                  Age.ToString & "," & _
                  TimeStep.ToString & "," & _
                  Cohort(Stk, Age, 0, TimeStep).ToString("######0.000000") & "," & _
                  Cohort(Stk, Age, 1, TimeStep).ToString("######0.000000") & "," & _
                  Cohort(Stk, Age, 4, TimeStep).ToString("######0.000000") & "," & _
                  Cohort(Stk, Age, 3, TimeStep).ToString("######0.000000") & "," & _
                  Cohort(Stk, Age, 2, TimeStep).ToString("######0.000000") & ")"
                  FCC.ExecuteNonQuery()

               End If
            Next
         Next
      Next
        CohTrans.Commit()
        FramDB.Close()

        CohortDA = Nothing

        '- ESCAPEMENT DataBase Table Save --------
        FVS_RunModel.RunProgressLabel.Text = " Saving ESCAPEMENT Table to Database ... Please Wait "
        myPoint = FVS_RunModel.RunProgressLabel.Location
        myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
        FVS_RunModel.RunProgressLabel.Location = myPoint
        FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
        FVS_RunModel.RunProgressLabel.Refresh()

        CmdStr = "SELECT * FROM Escapement WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, TimeStep"
        Dim ESCcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim EscapeDA As New System.Data.OleDb.OleDbDataAdapter
        EscapeDA.SelectCommand = ESCcm

        CmdStr = "DELETE * FROM Escapement WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim ESCDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        EscapeDA.DeleteCommand = ESCDcm

        Dim ESCcb As New OleDb.OleDbCommandBuilder
        ESCcb = New OleDb.OleDbCommandBuilder(EscapeDA)
        'EscapeDA.Fill(FramDataSet, "Escapement")

        Dim ESCTrans As OleDb.OleDbTransaction
        Dim FEC As New OleDbCommand

        FramDB.Open()
        EscapeDA.DeleteCommand.ExecuteScalar()
        ESCTrans = FramDB.BeginTransaction
        FEC.Connection = FramDB
        FEC.Transaction = ESCTrans
      For Stk As Integer = 1 To NumStk
         For Age As Integer = MinAge To MaxAge
            For TimeStep = 1 To NumSteps
               If Escape(Stk, Age, TimeStep) <> 0 Then
                  FEC.CommandText = "INSERT INTO Escapement (RunID,StockID,Age,TimeStep,Escapement) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Stk.ToString & "," & _
                  Age.ToString & "," & _
                  TimeStep.ToString & "," & _
                  Escape(Stk, Age, TimeStep).ToString("######0.000000") & ")"
                  FEC.ExecuteNonQuery()
               End If
            Next
         Next
      Next
        ESCTrans.Commit()
        FramDB.Close()

        EscapeDA = Nothing

        '- UPDATE RunID RunModified Field

        CmdStr = "SELECT * FROM RunID WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim RIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim RunIDDA As New System.Data.OleDb.OleDbDataAdapter
        RunIDDA.SelectCommand = RIDcm

        Dim RIDcb As New OleDb.OleDbCommandBuilder
        RIDcb = New OleDb.OleDbCommandBuilder(RunIDDA)
        If FramDataSet.Tables.Contains("RunID") Then
            FramDataSet.Tables("RunID").Clear()
        End If
        RunIDDA.Fill(FramDataSet, "RunID")
        Dim NumRID As Integer
        NumRID = FramDataSet.Tables("RunID").Rows.Count
        If NumRID <> 1 Then
            MsgBox("ERROR in RunID Table of Database ... Duplicate Record", MsgBoxStyle.OkOnly)
        End If
        FramDataSet.Tables("RunID").Rows(0)(9) = DateTime.Now

      '*********************Begin Pete 2/27/13 BC Flag Change to Run Name

        'If SpeciesName = "COHO" Then
        FramDataSet.Tables("RunID").Rows(0)(3) = RunIDNameSelect
        'End If

        '*********************End Pete 2/27/13 BC Flag Change to Run Name

        FramDataSet.Tables("RunID").Rows(0)(12) = TAMMName
        FramDataSet.Tables("RunID").Rows(0)(13) = CoastalIter
        FramDataSet.Tables("RunID").Rows(0)(14) = FRAMVers

        RunIDDA.Update(FramDataSet, "RunID")
        RunIDDA = Nothing
        '- UpDate Memory Variable for Run Time!
        RunIDRunTimeDateSelect = DateTime.Now



        EndTime = DateTime.Now
        TimeDiff1 = (EndTime.Ticks - StartTime.Ticks) / 10000000
        Jim = 1

        '- Update FisheryScaler Field in FisheryScalers Table for Quota Fisheries

        CmdStr = "SELECT * FROM FisheryScalers WHERE RunID = " & RunIDSelect.ToString & " ORDER BY FisheryID, TimeStep"
        Dim FSHcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim FishDA As New System.Data.OleDb.OleDbDataAdapter
        FishDA.SelectCommand = FSHcm

        CmdStr = "DELETE * FROM FisheryScalers WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim FSHDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        FishDA.DeleteCommand = FSHDcm

        Dim FSHcb As New OleDb.OleDbCommandBuilder
        FSHcb = New OleDb.OleDbCommandBuilder(FishDA)

        Dim FSHTrans As OleDb.OleDbTransaction
        Dim FSH As New OleDbCommand

        FramDB.Open()
        FishDA.DeleteCommand.ExecuteScalar()
        FSHTrans = FramDB.BeginTransaction
        FSH.Connection = FramDB
        FSH.Transaction = FSHTrans
        For Fish As Integer = 1 To NumFish
            For TimeStep = 1 To NumSteps
                FSH.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                   "VALUES(" & RunIDSelect.ToString & "," & _
                   Fish.ToString & "," & _
                   TimeStep.ToString & "," & _
                   FisheryFlag(Fish, TimeStep).ToString & "," & _
                   FisheryScaler(Fish, TimeStep).ToString("0.0000") & "," & _
                   FisheryQuota(Fish, TimeStep).ToString("######0.00000") & "," & _
                   MSFFisheryScaler(Fish, TimeStep).ToString("0.0000") & "," & _
                   MSFFisheryQuota(Fish, TimeStep).ToString("######0.00000") & "," & _
                   MarkSelectiveMortRate(Fish, TimeStep).ToString("0.0000") & "," & _
                   MarkSelectiveMarkMisID(Fish, TimeStep).ToString("0.0000") & "," & _
                   MarkSelectiveUnMarkMisID(Fish, TimeStep).ToString("0.0000") & "," & _
                   MarkSelectiveIncRate(Fish, TimeStep).ToString("0.0000") & ")"
                FSH.ExecuteNonQuery()
            Next
        Next
        FSHTrans.Commit()
        FramDB.Close()

        '- Save Total FisheryMortality Table 
        CmdStr = "SELECT * FROM FisheryMortality WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, TimeStep"
        Dim TFMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim TFMDA As New System.Data.OleDb.OleDbDataAdapter
        TFMDA.SelectCommand = TFMcm

        CmdStr = "DELETE * FROM FisheryMortality WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim TFMDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        TFMDA.DeleteCommand = TFMDcm

        Dim TFMcb As New OleDb.OleDbCommandBuilder
        TFMcb = New OleDb.OleDbCommandBuilder(EscapeDA)

        Dim TFMTrans As OleDb.OleDbTransaction
        Dim TFM As New OleDbCommand

        FramDB.Open()
        TFMDA.DeleteCommand.ExecuteScalar()
        TFMTrans = FramDB.BeginTransaction
        TFM.Connection = FramDB
        TFM.Transaction = TFMTrans
        Dim TotFM As Double
        For Fish As Integer = 1 To NumFish
            For TimeStep = 1 To NumSteps
                TotFM = TotalLandedCatch(Fish, TimeStep) + TotalLandedCatch(NumFish + Fish, TimeStep) + TotalEncounters(Fish, TimeStep) + TotalShakers(Fish, TimeStep) + TotalDropOff(Fish, TimeStep) + TotalNonRetention(Fish, TimeStep)
                If TotFM <> 0 Then
                    TFM.CommandText = "INSERT INTO FisheryMortality (RunID,FisheryID,TimeStep,TotalLandedCatch,TotalUnMarkedCatch,TotalNonRetention,TotalShakers,TotalDropOff,TotalEncounters) " & _
                    "VALUES(" & RunIDSelect.ToString & "," & _
                    Fish.ToString & "," & _
                    TimeStep.ToString & "," & _
                    TotalLandedCatch(Fish, TimeStep).ToString("#######0.000000") & "," & _
                    TotalLandedCatch(NumFish + Fish, TimeStep).ToString("#######0.000000") & "," & _
                    TotalNonRetention(Fish, TimeStep).ToString("#######0.000000") & "," & _
                    TotalShakers(Fish, TimeStep).ToString("#######0.000000") & "," & _
                    TotalDropOff(Fish, TimeStep).ToString("#######0.000000") & "," & _
                    TotalEncounters(Fish, TimeStep).ToString("#######0.000000") & ")"
                    TFM.ExecuteNonQuery()
                End If
            Next
        Next
        TFMTrans.Commit()
        FramDB.Close()

        '===================================================================================
        'Pete 12/13 -- Save recordset code for SLRatio and RunEncounterRateAdjustment Tables
        'After final pass of update run, go ahead and replace the RunEncounterRateAdjustment values
        'THIS CODE DOESN'T RUN UNLESS IT'S AN Update Run.

        If (SpeciesName = "CHINOOK" And UpdateRunEncounterRateAdjustment = True And FinalUpdatePass = True) Then

            CmdStr = "DELETE * FROM SLRatio WHERE RunID = " & RunIDSelect.ToString & ";"
            Dim SLRatDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim SLRatDA As New System.Data.OleDb.OleDbDataAdapter
            SLRatDA.DeleteCommand = SLRatDcm


            '- SLRatio DataBase Table Save --------
            Dim SLRatC As New OleDbCommand
            FramDB.Open()
            SLRatDA.DeleteCommand.ExecuteScalar()
            FramTrans = FramDB.BeginTransaction
            SLRatC.Connection = FramDB
            SLRatC.Transaction = FramTrans
            For Fish As Integer = 1 To NumFish
                For Age As Integer = MinAge To MaxAge
                    For TimeStep = 1 To NumSteps
                        'Uncomment the If end if content if cluttering with 1.00 is undesired; for now, testing, leave it in...
                        If TargetRatio(Fish, Age, TimeStep) <> -1 Then
                            SLRatC.CommandText = "INSERT INTO SLRatio (RunID,FisheryID,Age,TimeStep,TargetRatio,RunEncounterRateAdjustment,UpdateWhen,UpdateBy) " & _
                            "VALUES(" & RunIDSelect.ToString & "," & _
                            Fish.ToString & "," & _
                            Age.ToString & "," & _
                            TimeStep.ToString & "," & _
                            TargetRatio(Fish, Age, TimeStep).ToString & "," & _
                            RunEncounterRateAdjustment(Fish, Age, TimeStep).ToString & "," & _
                            "'" & DateTime.Now.ToString & "'" & "," & _
                            "'" & WhoUpdated.ToString & "'" & ")"
                            SLRatC.ExecuteNonQuery()
                        End If
                    Next
                Next
            Next
            FramTrans.Commit()
            FramDB.Close()

        End If
        '===================================================================================





        TFMDA = Nothing

    End Sub

    Sub ScaleCohort()
        '**************************************************************************
        'Subroutine computes cohort abundance by multiplying the base period cohort
        ' by the abundance scale factor for the current year.
        '**************************************************************************

        '- Reset Starting Cohort Size to Base Period Value for Backwards Coho FRAM
        If SpeciesName = "COHO" And RunBackFramFlag <> 0 Then
            Age = 3
            ' prevent ER from exceeding 100% otherwise MSF bias corrected equation produce error
            If BackFRAMIteration < 8 Then 'don't start bias calculations until target escapemetns are sufficiently close
                MSFBiasFlag = False
            Else
                MSFBiasFlag = SaveInitialFlag
            End If
            If BackFRAMIteration < 2 Then
                'start with a recruit scalar on first iteration that is sufficiently large to hold potentially huge catch inputs
                'without producing negative escapements and ER>100%
                For Stk = 1 To NumStk
                    If Stk = 5 Then
                        Jim = 1
                    End If
                    If BackwardsFlag(Stk) > 0 Or RunBackwardsFlag(Stk) > 0 Then
                        Cohort(Stk, Age, PTerm, 1) = 1000 * BaseCohortSize(Stk, Age)
                    End If
                Next Stk
                Exit Sub
            End If
        End If
        '- Reset TIME-1 AGE 3-5 Cohort Sizes to Initial Value for Backwards Chinook FRAM
        '- Must do this because TIME-4 Ages Cohort Sizes for "Next Year"
        If SpeciesName = "CHINOOK" And BackwardsFRAMFlag = 1 Then
            For Stk As Integer = 1 To NumStk
                For Age As Integer = 3 To 5
                    Cohort(Stk, Age, PTerm, 1) = BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)
                Next Age
            Next Stk
            Exit Sub
        End If

        '- Apply Stock Recruit Scaler to Base Period Cohort Size TIME-1 ... All Species
        Dim Jim1, Jim2, Jim3 As Double
        Dim Trm, TP As Integer
        '- Zero Cohort Array
        For Stk As Integer = 0 To NumStk
            For Age As Integer = 0 To MaxAge
                For Trm = 0 To 2
                    For TP = 0 To NumSteps
                        Cohort(Stk, Age, Trm, TP) = 0
                    Next
                Next
            Next Age
        Next Stk
        For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
                Cohort(Stk, Age, PTerm, 1) = BaseCohortSize(Stk, Age) * StockRecruit(Stk, Age, 1)
                Jim1 = Cohort(Stk, Age, PTerm, 1)
                Jim2 = BaseCohortSize(Stk, Age)
                Jim3 = StockRecruit(Stk, Age, 1)
            Next Age
        Next Stk

    End Sub


    'Here I Am ... Changing program to be able to handle NSF + MSF same time step


    Sub TCHNComp(ByVal TammIteration)       'CHINOOK TAMM Comparision Routine

        '**************************************************************************
        ' TAMM Variable Processing - Iteratively Solve for CAM Effort Scalars that
        '                            equal the BaseExploitationRate catch from PS Net Estimates
        '**************************************************************************
        Dim StartStk(7), StopStk(7), Area As Integer
        Dim TammDiff, TammLoop As Integer

        'SET STOCK AGGREGATES FOR TERMINAL RUNS
        StartStk(1) = 1        '--- Nooksack/Samish Summer/Fall
        StopStk(1) = 1
        StartStk(2) = 4        '--- Skagit Summer/Fall
        StopStk(2) = 4
        StartStk(3) = 7        '--- Stillaguamish/Snohomish Summer/Fall
        StopStk(3) = 10
        StartStk(4) = 10       '--- Tulalip Fall
        StopStk(4) = 10
        StartStk(5) = 16       '--- Hood Canal Fall
        StopStk(5) = 17
        StartStk(6) = 2        '--- Nooksack Spring
        StopStk(6) = 3
        StartStk(7) = 15       '--- White River (13A) Spring
        StopStk(7) = 15

        '----------------- Compute Total Terminal Run in PS Net Scale of Reference ---

        'INITIALIZE TERMINAL RUNS     1 = Nooksack Fall
        '                             2 = Skagit Fall
        '                             3 = Still./Snohomish/Tulalip Fall
        '                             4 = Tulalip Fall
        '                             5 = Hood Canal Fall
        '                             6 = Nooksack Spring
        '                             7 = White River (13A) Spring

        For Area = 1 To 7
            TammTermRun(Area) = 0
        Next Area

        For Area = 1 To 7             '--- SUM ESCAPEMENT OVER TIME STEPS
            For TStep = 1 To NumSteps
                TammTermRun(Area) = TammTermRun(Area) + TammEscape(Area, TStep)
            Next TStep
        Next Area

        '-ADD IN FRESHWATER-NET and FRESHWATER-SPORT CATCH TO GET EXTREME TERMINAL RUN.

      For Fish As Integer = 72 To 73
         For Area = 1 To 7
            For Stk As Integer = StartStk(Area) To StopStk(Area)
               For TStep As Integer = 1 To NumSteps
                  For Age As Integer = 3 To MaxAge   '--- Only age 3-5 Fish in Freshwater
                     TammTermRun(Area) = TammTermRun(Area) + LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  Next Age
               Next TStep
            Next Stk
         Next Area
      Next Fish

        '----- ADD IN Time Period 3 CATCHES FOR TERMINAL AREA NET FISHERIES -----

        For Area = 1 To 5
            TammTermRun(Area) = TammTermRun(Area) + TammCatch(Area * 2 - 1, 3) + TammCatch(Area * 2, 3)
        Next Area
        For Area = 6 To 7
            TammTermRun(Area) = TammTermRun(Area) + TammCatch(Area * 2 + 1, 3) + TammCatch(Area * 2 + 2, 3)
        Next Area

        '--- SKAGIT SUMMER/FALL, ADD AREA 8 TREATY AND NONTREATY CATCH, TIME STEP 2
        TammTermRun(2) = TammTermRun(2) + TammCatch(3, 2) + TammCatch(4, 2)

        '--- SUBTRACT FRESHWATER SPORT AND ADJUST MARINE SPORT
        'TammTermRun(1) = TammTermRun(1) - TNkFWSpt! + TNkMSA!
        'TammTermRun(2) = TammTermRun(2) - TSkFWSpt! + TSkMSA!
        'TammTermRun(3) = TammTermRun(3) - TSnFWSpt! + TSnMSA!
        'TammTermRun(5) = TammTermRun(5) - THCFWSpt!

        '--- 3/23/99 Tulalip Rates now use  Tulalip ETRS
        ' TammTermRun(4) = TammTermRun(3)

        For Area = 1 To 8         '--- Initialize TAMM Estimate Array
            For TStep = 1 To 4
                TammEstimate(Area * 2 - 1, TStep) = 0.0
                TammEstimate(Area * 2, TStep) = 0.0
            Next TStep
        Next Area

        '------------------------ NOOKSACK FALL --------------------------
        Area = 1
        TStep = 3
        '--- NonTreaty 7B-D Net
        If TammPSER(1, 3) = -88 Then
            GoTo SkipNoSant2
        End If
        If TammPSER(1, 3) < 1 Then          '--- Use Harvest Rate
            TammEstimate(1, 3) = TammTermRun(1) * TammPSER(1, 3)
        Else
            If TammPSER(1, 3) > 1 Then       '--- Use Target Quota
                TammEstimate(1, 3) = TammPSER(1, 3)
            End If
        End If
SkipNoSant2:
        '--- Treaty 7B-D Net
        If TammPSER(2, 3) = -88 Then
            GoTo SkipNoSnt2
        End If
        If TammPSER(2, 3) < 1 Then          '--- Use Harvest Rate
            TammEstimate(2, 3) = TammTermRun(1) * TammPSER(2, 3)
        Else
            If TammPSER(2, 3) > 1 Then       '--- Use Target Quota
                TammEstimate(2, 3) = TammPSER(2, 3)
            End If
        End If
SkipNoSnt2:

        '--- NOOKSACK NATIVE (SPRING)
        TammEstimate(13, 2) = TammTermRun(6) * TammPSER(21, 2)
        TammEstimate(14, 2) = TammTermRun(6) * TammPSER(22, 2)
        TammEstimate(13, 3) = TammTermRun(6) * TammPSER(21, 3)
        TammEstimate(14, 3) = TammTermRun(6) * TammPSER(22, 3)

        '- WhRvrSpr not used in 13A Net
        '--- WHITE RIVER SPRING
        'TammEstimate(15, 2) = TammTermRun(7) * TammPSER(23, 2)
        'TammEstimate(16, 2) = TammTermRun(7) * TammPSER(24, 2)
        'TammEstimate(15, 3) = TammTermRun(7) * TammPSER(23, 3)
        'TammEstimate(16, 3) = TammTermRun(7) * TammPSER(24, 3)

        '========================================================================='
        '--- Calculate Estimated Catches for Rate Fisheries in Time Period 3
        '--- Other Time Periods are Target Catches except Skagit Time 2
        '--- 1/19/96  Allow for Target Quotas in Time Step 3 for Validation Runs JFP
        TStep = 3
        For Area = 2 To 5
            If TammPSER(Area * 2 - 1, 3) > 1 Then           '--- Target
                TammEstimate(Area * 2 - 1, 3) = TammPSER(Area * 2 - 1, 3)
            Else                                              '--- Harvest Rate
                TammEstimate(Area * 2 - 1, 3) = TammTermRun(Area) * TammPSER(Area * 2 - 1, 3)
            End If
            If TammPSER(Area * 2, 3) > 1 Then               '--- Target
                TammEstimate(Area * 2, 3) = TammPSER(Area * 2, 3)
            Else                                              '--- Harvest Rate
                TammEstimate(Area * 2, 3) = TammTermRun(Area) * TammPSER(Area * 2, 3)
            End If
        Next Area

        '--- 13A Net Target Fisheries
        TammEstimate(11, 2) = TammPSER(19, 2)
        TammEstimate(12, 2) = TammPSER(20, 2)
        TammEstimate(11, 3) = TammPSER(19, 3)
        TammEstimate(12, 3) = TammPSER(20, 3)
        '--- Nooksack Fall Time 2 Target
        TammEstimate(1, 2) = TammPSER(1, 2)
        TammEstimate(2, 2) = TammPSER(2, 2)
        '---- Skagit Time Period 2
        If TammPSER(3, 2) > 1.0 Then
            TammEstimate(3, 2) = TammPSER(3, 2)               '--- Target Quota!
        Else
            TammEstimate(3, 2) = TammTermRun(2) * TammPSER(3, 2)  '--- Terminal Rate
        End If
        If TammPSER(4, 2) > 1.0 Then
            TammEstimate(4, 2) = TammPSER(4, 2)               '--- Target Quota!
        Else
            TammEstimate(4, 2) = TammTermRun(2) * TammPSER(4, 2)  '--- Terminal Rate
        End If

        '--- Special Case Tulalip Bay HR=.99 ... Catch All Mature Fish
        '--- for Chinook TAMM   3/26/99 JFP
        '--- NT Time 3 Net Rate = Sport HR of ETRS  3/30/99
        '--- Note: The Tulalip Bay Net fishery cannot be MSF with this code
        If FisheryScaler(52, 3) = 0.99 Then
            TammEstimate(7, 3) = 0
            TammEstimate(8, 3) = 0
         For Age As Integer = MinAge To 5
            TotalLandedCatch(51, 3) = TotalLandedCatch(51, 3) - LandedCatch(10, Age, 51, 3) - MSFLandedCatch(10, Age, 51, 3)
            TotalDropOff(51, 3) = TotalDropOff(51, 3) - DropOff(10, Age, 51, 3) - MSFDropOff(10, Age, 51, 3)
            TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) - LandedCatch(10, Age, 52, 3) - MSFLandedCatch(10, Age, 52, 3)
            TotalDropOff(52, 3) = TotalDropOff(52, 3) - DropOff(10, Age, 52, 3) - MSFDropOff(10, Age, 52, 3)
            DropOff(10, Age, 52, 3) = (Escape(10, Age, 3) * (1 - FisheryScaler(51, 3))) * IncidentalRate(52, 3)
            DropOff(10, Age, 51, 3) = (Escape(10, Age, 3) * FisheryScaler(51, 3)) * IncidentalRate(51, 3)
            TotalDropOff(52, 3) = TotalDropOff(52, 3) + DropOff(10, Age, 52, 3) + MSFDropOff(10, Age, 52, 3)
            TotalDropOff(51, 3) = TotalDropOff(51, 3) + DropOff(10, Age, 51, 3) + MSFDropOff(10, Age, 51, 3)
            LandedCatch(10, Age, 52, 3) = (Escape(10, Age, 3) - DropOff(10, Age, 52, 3) - DropOff(10, Age, 51, 3) - MSFDropOff(10, Age, 52, 3) - MSFDropOff(10, Age, 51, 3)) * (1 - FisheryScaler(51, 3))
            LandedCatch(10, Age, 51, 3) = (Escape(10, Age, 3) - DropOff(10, Age, 52, 3) - DropOff(10, Age, 51, 3) - MSFDropOff(10, Age, 52, 3) - MSFDropOff(10, Age, 51, 3)) * FisheryScaler(51, 3)
            TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) + LandedCatch(10, Age, 52, 3) + MSFLandedCatch(10, Age, 52, 3)
            TotalLandedCatch(51, 3) = TotalLandedCatch(51, 3) + LandedCatch(10, Age, 51, 3) + MSFLandedCatch(10, Age, 51, 3)
            Escape(10, Age, 3) = 0
         Next Age
            TammEstimate(8, 3) = TotalLandedCatch(52, 3)
            TammEstimate(7, 3) = TotalLandedCatch(51, 3)
            TammCatch(7, 3) = TammEstimate(7, 3)
            TammCatch(8, 3) = TammEstimate(8, 3)
            TammTermRun(4) = TammTermRun(4) + TammCatch(8, 3) + TammCatch(7, 3)
        End If

        '----------------------- Compare TAMM and FRAM Catch by Time Period ---
        TammChinookConverge = 0

        '--- First Check Spring Chinook Impacts

        'Call To CHKSPRCH Parameters = TStep,Fish1,Fish2,Fish3,Stk
        '--- Fish1 = Spring TammCatch and TammEstimate
        '--- Fish2 = Fall TamkRate
        '--- Fish3 = Fall TammCatch

        NewStockFishRateScalers = 0
        Call CHKSPRCH(2, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 2
        Call CHKSPRCH(2, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 2
        Call CHKSPRCH(3, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 3
        Call CHKSPRCH(3, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 3
        '- 13A Not Used for WhRvr Spr
        'Call CHKSPRCH(2, 15, 70, 11, 14) 'White River 13A NT Net TStep 2
        'Call CHKSPRCH(2, 16, 71, 12, 14) 'White River 13A TR Net TStep 2
        'Call CHKSPRCH(3, 15, 70, 11, 14) 'White River 13A NT Net TStep 3
        'Call CHKSPRCH(3, 16, 71, 12, 14) 'White River 13A TR Net TStep 3

        '------------ Check for Time = 3 for All Fisheries then for other Time = 2
        TStep = 3
        For Area = 1 To 12
            If TammCatch(Area, TStep) = 0.0 Then
                If TammEstimate(Area, TStep) = 0.0 Then
                    TammScaler(Area, TStep) = 1.0
                Else
                    '-------------------------------------------------- Exception Case ----------
                    ' If Model Catch is 0.0 but the TAMM Estimated Catch is not 0.0 then the
                    ' effort scalar is most likely set to 0.0000 and the harvest rate or monthly
                    ' percent split for TAMM is not 0.0 ---- hmmm ---- To solve this problem the
                    ' effort scalar will be reset to 0.1000 and the next iteration will change it
                    ' to the appropriate value.  The only way this won't work is if there are no
                    ' base period harvest rates for this time period, which would then be an input
                    ' error by the person creating the TAMM input file.
                    If AnyBaseRate(TamkFish(Area), TStep) = 0 Then
                        TammScaler(Area, TStep) = 1.0
                    Else
                        TammScaler(Area, TStep) = 0.1
                        If FisheryFlag(TamkFish(Area), TStep) = 1 Then
                            FisheryScaler(TamkFish(Area), TStep) = 1.0
                        Else
                            FisheryQuota(TamkFish(Area), TStep) = 1000.0
                        End If
                    End If
                End If
            Else
                TammScaler(Area, TStep) = TammEstimate(Area, TStep) / TammCatch(Area, TStep)
            End If
            TammDiff = Abs(TammEstimate(Area, TStep) - TammCatch(Area, TStep))
            If ((TammScaler(Area, TStep) > 1.0001 Or TammScaler(Area, TStep) < 0.9999) And TammDiff > 1.0) Then
                TammChinookConverge = 1
            End If
            If TammPSER(Area, TStep) = -88 Then
                TammScaler(Area, TStep) = 1
            End If

            PrnLine = "Area,TS,TammEstimate,TammCatch " & Area.ToString & " " & TStep.ToString & " " & TammEstimate(Area, TStep).ToString("######0.0") & " " & TammCatch(Area, TStep).ToString("######0.0") & " " & TammScaler(Area, TStep).ToString("###0.0000") & " " & TammChinookConverge.ToString
            sw.WriteLine(PrnLine)
        Next Area
        '--------------------------------------- Time Period 2 Nooksack, Skagit, 13A ---
        TStep = 2
        For TammLoop = 1 To 6
            If TammLoop > 4 Then
                Area = TammLoop + 6
            Else
                Area = TammLoop
            End If
            If TammCatch(Area, TStep) = 0.0 Then
                If TammEstimate(Area, TStep) = 0.0 Then
                    TammScaler(Area, TStep) = 1.0
                Else
                    If (FisheryFlag(TamkFish(Area), TStep) <> 1) Then
                        FisheryScaler(TamkFish(Area), TStep) = 1.0
                    End If
                End If
            Else
                If TammPSER(Area, TStep) = -88 Then
                    TammScaler(Area, TStep) = 1
                Else
                    TammScaler(Area, TStep) = TammEstimate(Area, TStep) / TammCatch(Area, TStep)
                End If
            End If
            TammDiff = Abs(TammEstimate(Area, TStep) - TammCatch(Area, TStep))
            If ((TammScaler(Area, TStep) > 1.0001 Or TammScaler(Area, TStep) < 0.9999) And TammDiff > 1.0) Then
                TammChinookConverge = 1
            End If
            PrnLine = "Area,TS,TammEstimate,TammCatch " & Area.ToString & " " & TStep.ToString & " " & TammEstimate(Area, TStep).ToString("######0.0") & " " & TammCatch(Area, TStep).ToString("######0.0") & " " & TammScaler(Area, TStep).ToString("###0.0000") & " " & TammChinookConverge.ToString
            sw.WriteLine(PrnLine)
        Next TammLoop

        For TStep = 2 To 3
            For Area = 13 To 14
                PrnLine = "Area,TS,TammEstimate,TammCatch "
                PrnLine &= String.Format("{0,3}", Area.ToString)
                PrnLine &= String.Format("{0,3}", TStep.ToString)
                PrnLine &= String.Format("{0,10}", TammEstimate(Area, TStep).ToString("######0.0"))
                PrnLine &= String.Format("{0,10}", TammCatch(Area, TStep).ToString("######0.0"))
                sw.WriteLine(PrnLine)
            Next Area
        Next TStep

    End Sub
    Sub TCHNSFComp(ByVal TammIteration) 'CHINOOK Selective Fishery TAMM Comparision Routine

        '**************************************************************************
        ' TAMM Variable Processing - Iteratively Solve for FRAM Effort Scalars that
        '                            equal the HR catch from PS Net TAMM Estimates
        '**************************************************************************

        Dim Area, TammDiff, TammLoop As Integer
        Dim TamkFish(12) As Integer

        TamkFish(1) = 39         '--- Nooksack-Samish 7BCD Net
        TamkFish(2) = 40
        TamkFish(3) = 46         '--- Skagit Bay 8 Net
        TamkFish(4) = 47
        TamkFish(5) = 49         '--- Still/Snohomish 8A Net
        TamkFish(6) = 50
        TamkFish(7) = 51         '--- Tulalip Bay 8D Net
        TamkFish(8) = 52
        TamkFish(9) = 65         '--- Hood Canal 12-12D Net
        TamkFish(10) = 66
        TamkFish(11) = 70        '--- Carr Inlet 13A Net
        TamkFish(12) = 71

        'SET STOCK AGGREGATES FOR TERMINAL RUNS

        Dim StartStk(7), StopStk(7) As Integer

        StartStk(1) = 1        '--- Nooksack/Samish Summer/Fall
        StopStk(1) = 1
        StartStk(2) = 4        '--- Skagit Summer/Fall
        StopStk(2) = 4
        StartStk(3) = 7        '--- Stillaguamish/Snohomish Summer/Fall
        StopStk(3) = 10
        StartStk(4) = 10       '--- Tulalip Fall
        StopStk(4) = 10
        StartStk(5) = 16       '--- Hood Canal Fall
        StopStk(5) = 17
        StartStk(6) = 2        '--- Nooksack Spring
        StopStk(6) = 3
        StartStk(7) = 15       '--- White River (13A) Spring
        StopStk(7) = 15

        '----------------- Compute Total Terminal Run in PS Net Scale of Reference ---

        'INITIALIZE TERMINAL RUNS     1 = Nooksack Fall
        '                             2 = Skagit Fall
        '                             3 = Still./Snohomish/Tulalip Fall
        '                             4 = Tulalip Fall
        '                             5 = Hood Canal Fall
        '                             6 = Nooksack Spring
        '                             7 = White River (13A) Spring

        For Area = 1 To 7
            TammTermRun(Area) = 0
        Next Area

        For Area = 1 To 7             '--- SUM ESCAPEMENT OVER TIME STEPS
            For TStep = 1 To NumSteps
                TammTermRun(Area) = TammTermRun(Area) + TammEscape(Area, TStep)
            Next TStep
        Next Area

        '-ADD IN FRESHWATER-NET and FRESHWATER-SPORT CATCH TO GET EXTREME TERMINAL RUN.

        For Fish = 72 To 73
            For Area = 1 To 7
            For Stk = StartStk(Area) To StopStk(Area)
               For TStep = 1 To NumSteps
                  For Age = 3 To MaxAge   '--- Only age 3-5 Fish in Freshwater
                     TammTermRun(Area) = TammTermRun(Area) + LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep)
                     TammTermRun(Area) = TammTermRun(Area) + LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep)
                  Next Age
               Next TStep
            Next Stk
            Next Area
        Next Fish

        '----- ADD IN Time Period 3 CATCHES FOR TERMINAL AREA NET FISHERIES -----

        For Area = 1 To 5
            TammTermRun(Area) = TammTermRun(Area) + TammCatch(Area * 2 - 1, 3) + TammCatch(Area * 2, 3)
        Next Area
        For Area = 6 To 7
            TammTermRun(Area) = TammTermRun(Area) + TammCatch(Area * 2 + 1, 3) + TammCatch(Area * 2 + 2, 3)
        Next Area

        '--- SKAGIT SUMMER/FALL, ADD AREA 8 TREATY AND NONTREATY CATCH, TIME STEP 2
        TammTermRun(2) = TammTermRun(2) + TammCatch(3, 2) + TammCatch(4, 2)

        '--- SUBTRACT FRESHWATER SPORT AND ADJUST MARINE SPORT
        'TammTermRun(1) = TammTermRun(1) - TNkFWSpt! + TNkMSA!
        'TammTermRun(2) = TammTermRun(2) - TSkFWSpt! + TSkMSA!
        'TammTermRun(3) = TammTermRun(3) - TSnFWSpt! + TSnMSA!
        'TammTermRun(5) = TammTermRun(5) - THCFWSpt!
        '--- 3/23/99 Tulalip Rates now use  Tulalip ETRS
        ' TammTermRun(4) = TammTermRun(3)

        For Area = 1 To 8
            For TStep = 1 To 4      '--- Initialize Estimate Array
                TammEstimate(Area * 2 - 1, TStep) = 0.0
                TammEstimate(Area * 2, TStep) = 0.0
            Next TStep
        Next Area

        '------------------------ NOOKSACK FALL --------------------------
        Area = 1
        TStep = 3
        '--- NonTreaty 7B-D Net
        If TammPSER(1, 3) = -88 Then
            GoTo SkipNoSant
        End If
        If TammPSER(1, 3) < 1 Then          '--- Use Harvest Rate
            TammEstimate(1, 3) = TammTermRun(1) * TammPSER(1, 3)
        Else
            If TammPSER(1, 3) > 1 Then       '--- Use Target Quota
                TammEstimate(1, 3) = TammPSER(1, 3)
            End If
        End If
SkipNoSant:
        '--- Treaty 7B-D Net
        If TammPSER(2, 3) = -88 Then
            GoTo SkipNoSat
        End If
        If TammPSER(2, 3) < 1 Then          '--- Use Harvest Rate
            TammEstimate(2, 3) = TammTermRun(1) * TammPSER(2, 3)
        Else
            If TammPSER(2, 3) > 1 Then       '--- Use Target Quota
                TammEstimate(2, 3) = TammPSER(2, 3)
            End If
        End If
SkipNoSat:

        '--- NOOKSACK NATIVE (SPRING)
        TammEstimate(13, 2) = TammTermRun(6) * TammPSER(21, 2)
        TammEstimate(14, 2) = TammTermRun(6) * TammPSER(22, 2)
        TammEstimate(13, 3) = TammTermRun(6) * TammPSER(21, 3)
        TammEstimate(14, 3) = TammTermRun(6) * TammPSER(22, 3)

        '--- WHITE RIVER SPRING
        '... 13A Not Used for WhRvr Spr 9/20/2002
        TammEstimate(15, 2) = TammTermRun(7) * TammPSER(23, 2)
        TammEstimate(16, 2) = TammTermRun(7) * TammPSER(24, 2)
        TammEstimate(15, 3) = TammTermRun(7) * TammPSER(23, 3)
        TammEstimate(16, 3) = TammTermRun(7) * TammPSER(24, 3)

        '========================================================================='
        '--- Calculate Estimated Catches for Rate Fisheries in Time Period 3
        '--- Other Time Periods are Target Catches except Skagit Time 2
        '--- 1/19/96  Allow for Target Quotas in Time Step 3 for Validation Runs JFP
        TStep = 3
        For Area = 2 To 5
            If TammPSER(Area * 2 - 1, 3) >= 1 Then           '--- Target
                TammEstimate(Area * 2 - 1, 3) = TammPSER(Area * 2 - 1, 3)
            Else                                              '--- Harvest Rate
                TammEstimate(Area * 2 - 1, 3) = TammTermRun(Area) * TammPSER(Area * 2 - 1, 3)
            End If
            If TammPSER(Area * 2, 3) >= 1 Then               '--- Target
                TammEstimate(Area * 2, 3) = TammPSER(Area * 2, 3)
            Else                                              '--- Harvest Rate
                TammEstimate(Area * 2, 3) = TammTermRun(Area) * TammPSER(Area * 2, 3)
            End If
        Next Area

        '--- 13A Net Target Fisheries
        '... 13A Not Used for WhRvr Spr 9/20/2002
        TammEstimate(11, 2) = TammPSER(19, 2)
        TammEstimate(12, 2) = TammPSER(20, 2)
        TammEstimate(11, 3) = TammPSER(19, 3)
        TammEstimate(12, 3) = TammPSER(20, 3)
        '--- Nooksack Fall Time 2 Target
        TammEstimate(1, 2) = TammPSER(1, 2)
        TammEstimate(2, 2) = TammPSER(2, 2)
        '---- Skagit Time Period 2
        If TammPSER(3, 2) >= 1.0 Then
            TammEstimate(3, 2) = TammPSER(3, 2)               '--- Target Quota!
        Else
            TammEstimate(3, 2) = TammTermRun(2) * TammPSER(3, 2)  '--- Terminal Rate
        End If
        If TammPSER(4, 2) >= 1.0 Then
            TammEstimate(4, 2) = TammPSER(4, 2)               '--- Target Quota!
        Else
            TammEstimate(4, 2) = TammTermRun(2) * TammPSER(4, 2)  '--- Terminal Rate
        End If

        '--- Special Case Tulalip Bay HR=.99 ... Catch All Mature Fish
        '--- for Chinook TAMM   3/26/99 JFP
        '--- NT Time 3 Net Rate = Sport HR of ETRS  3/30/99
        '... UnMarked/Marked Change 9/20/2002
        '--- Note: The Tulalip Bay Net fishery cannot be MSF with this code
        If FisheryScaler(52, 3) = 0.99 Then
            TammEstimate(7, 3) = 0
            TammEstimate(8, 3) = 0
            For Age = MinAge To 5
                TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) - LandedCatch(19, Age, 52, 3) - MSFLandedCatch(19, Age, 52, 3)
                TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) - LandedCatch(20, Age, 52, 3) - MSFLandedCatch(20, Age, 52, 3)
                TotalDropOff(52, 3) = TotalDropOff(52, 3) - DropOff(19, Age, 52, 3) - MSFDropOff(19, Age, 52, 3)
                TotalDropOff(52, 3) = TotalDropOff(52, 3) - DropOff(20, Age, 52, 3) - MSFDropOff(20, Age, 52, 3)
                DropOff(19, Age, 52, 3) = Escape(19, Age, 3) * IncidentalRate(52, 3)
                DropOff(20, Age, 52, 3) = Escape(20, Age, 3) * IncidentalRate(52, 3)
                TotalDropOff(52, 3) = TotalDropOff(52, 3) + DropOff(19, Age, 52, 3) + MSFDropOff(19, Age, 52, 3)
                TotalDropOff(52, 3) = TotalDropOff(52, 3) + DropOff(20, Age, 52, 3) + MSFDropOff(19, Age, 51, 3)
                LandedCatch(19, Age, 52, 3) = Escape(19, Age, 3) - DropOff(19, Age, 52, 3) - MSFDropOff(19, Age, 52, 3)
                LandedCatch(20, Age, 52, 3) = Escape(20, Age, 3) - DropOff(20, Age, 52, 3) - MSFDropOff(20, Age, 52, 3)
                TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) + LandedCatch(19, Age, 52, 3) + MSFLandedCatch(19, Age, 52, 3)
                TotalLandedCatch(52, 3) = TotalLandedCatch(52, 3) + LandedCatch(20, Age, 52, 3) + MSFLandedCatch(20, Age, 52, 3)
                Escape(19, Age, 3) = 0
                Escape(20, Age, 3) = 0
            Next Age
            TammEstimate(8, 3) = TotalLandedCatch(52, 3)
            TammEstimate(7, 3) = TotalLandedCatch(51, 3)
            TammCatch(7, 3) = TammEstimate(7, 3)
            TammCatch(8, 3) = TammEstimate(8, 3)
            TammTermRun(4) = TammTermRun(4) + TammCatch(8, 3) + TammCatch(7, 3)
        End If

        '----------------------- Compare TAMM and FRAM Catch by Time Period ---
        TammChinookConverge = 0

        '--- First Check Spring Chinook Impacts

        'Call To CHKSPRCH Parameters = TStep,Fish1,Fish2,Fish3,Stk
        '--- Fish1 = Spring TammCatch and TammEstimate
        '--- Fish2 = Fall TamkRate
        '--- Fish3 = Fall TammCatch

        '- Changed Stk for SF Chin FRAM 9/20/2002
        NewStockFishRateScalers = 0

        If TammPSER(22, 3) <> -88 Then 'AHB 8/16/2017 added if statement for new BP
            'let FRAM's BPERs calculate Nooksack Spring impacts in B'ham Bay net rather than using TAMM rate AHB 3/15/17
            If NumStk > 50 Then
                Call CHKSPRCHSF(2, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 2
                Call CHKSPRCHSF(2, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 2
                Call CHKSPRCHSF(3, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 3
                Call CHKSPRCHSF(3, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 3
            Else
                Call CHKSPRCH(2, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 2
                Call CHKSPRCH(2, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 2
                Call CHKSPRCH(3, 13, 39, 1, 2)   'Nooksack Native NT Net TStep 3
                Call CHKSPRCH(3, 14, 40, 2, 2)   'Nooksack Native TR Net TStep 3
            End If
        End If

        '... 13A Not Used for WhRvr Spr 9/20/2002
        'Call CHKSPRCH(2, 15, 70, 11, 14) 'White River 13A NT Net TStep 2
        'Call CHKSPRCH(2, 16, 71, 12, 14) 'White River 13A TR Net TStep 2
        'Call CHKSPRCH(3, 15, 70, 11, 14) 'White River 13A NT Net TStep 3
        'Call CHKSPRCH(3, 16, 71, 12, 14) 'White River 13A TR Net TStep 3

        If TammIteration = 13 Then Jim = 1

        '------------ Check for Time = 3 for All Fisheries then for other Time = 2
        TStep = 3
        PrnLine = "Chinook TAMM Time Step 3"
        sw.WriteLine(PrnLine)
        For Area = 1 To 12
            If CLng(TammCatch(Area, TStep) * 100000) = 0.0 Then
                If TammEstimate(Area, TStep) = 0.0 Then
                    TammScaler(Area, TStep) = 1.0
                Else
                    '-------------------------------------------------- Exception Case ----------
                    ' If Model Catch is 0.0 but the TAMM Estimated Catch is not 0.0 then the
                    ' effort scalar is most likely set to 0.0000 and the harvest rate or monthly
                    ' percent split for TAMM is not 0.0 ---- hmmm ---- To solve this problem the
                    ' effort scalar will be reset to 0.1000 and the next iteration will change it
                    ' to the appropriate value.  The only way this won't work is if there are no
                    ' base period harvest rates for this time period, which would then be an input
                    ' error by the person creating the TAMM input file.
                    If AnyBaseRate(TamkFish(Area), TStep) = 0 Then
                        TammScaler(Area, TStep) = 1.0
                    Else
                        TammScaler(Area, TStep) = 0.1
                        If FisheryFlag(TamkFish(Area), TStep) = 1 Then
                            FisheryScaler(TamkFish(Area), TStep) = 1.0
                        Else
                            FisheryQuota(TamkFish(Area), TStep) = 1000.0
                        End If
                    End If
                End If
            Else
                If TammPSER(Area, TStep) = -88 Then
                    TammScaler(Area, TStep) = 1
                Else
                    TammScaler(Area, TStep) = TammEstimate(Area, TStep) / TammCatch(Area, TStep)
                End If
            End If

            TammDiff = Abs(TammEstimate(Area, TStep) - TammCatch(Area, TStep))
            If (TammScaler(Area, TStep) > 1.001 Or TammScaler(Area, TStep) < 0.999) Then
                TammChinookConverge = 1
            ElseIf TammDiff > 2.0 Then
                TammChinookConverge = 1
            End If
            PrnLine = "Area,TS,TammEstimate,TammCatch "
            PrnLine &= String.Format("{0,3}", Area.ToString)
            PrnLine &= String.Format("{0,3}", TStep.ToString)
            PrnLine &= String.Format("{0,10}", TammEstimate(Area, TStep).ToString("######0.0"))
            PrnLine &= String.Format("{0,10}", TammCatch(Area, TStep).ToString("######0.0"))
            PrnLine &= String.Format("{0,10}", TammScaler(Area, TStep).ToString("###0.0000"))
            PrnLine &= String.Format("{0,3}", TammChinookConverge.ToString)
            If Area < 8 Then
                PrnLine &= String.Format("{0,10}", TammTermRun(Area).ToString("#####0.00"))
            End If
            sw.WriteLine(PrnLine)
        Next Area
        '--------------------------------------- Time Period 2 Nooksack, Skagit, 13A ---
        TStep = 2
        PrnLine = "Chinook TAMM Time Step 2"
        sw.WriteLine(PrnLine)
        For TammLoop = 1 To 6
            If TammLoop > 4 Then
                Area = TammLoop + 6
            Else
                Area = TammLoop
            End If
            If TammCatch(Area, TStep) = 0.0 Then
                If TammEstimate(Area, TStep) = 0.0 Then
                    TammScaler(Area, TStep) = 1.0
                Else
                    If AnyBaseRate(TamkFish(Area), TStep) = 0 Then
                        TammScaler(Area, TStep) = 1.0
                    Else
                        TammScaler(Area, TStep) = 0.1
                        If FisheryFlag(TamkFish(Area), TStep) = 1 Then
                            FisheryScaler(TamkFish(Area), TStep) = 1.0
                        Else
                            FisheryQuota(TamkFish(Area), TStep) = 1000.0
                        End If
                    End If
                End If
            Else
                TammScaler(Area, TStep) = TammEstimate(Area, TStep) / TammCatch(Area, TStep)
            End If
            TammDiff = Abs(TammEstimate(Area, TStep) - TammCatch(Area, TStep))
            If ((TammScaler(Area, TStep) > 1.001 Or TammScaler(Area, TStep) < 0.999) And TammDiff > 2.0) Then
                TammChinookConverge = 1
            End If
            PrnLine = "Area,TS,TammEstimate,TammCatch "
            PrnLine &= String.Format("{0,3}", Area.ToString)
            PrnLine &= String.Format("{0,3}", TStep.ToString)
            PrnLine &= String.Format("{0,10}", TammEstimate(Area, TStep).ToString("######0.0"))
            PrnLine &= String.Format("{0,10}", TammCatch(Area, TStep).ToString("######0.0"))
            PrnLine &= String.Format("{0,10}", TammScaler(Area, TStep).ToString("###0.0000"))
            PrnLine &= String.Format("{0,3}", TammChinookConverge.ToString)
            sw.WriteLine(PrnLine)
        Next TammLoop

        For TStep = 2 To 3
            For Area = 13 To 14
                PrnLine = "Area,TS,TammEstimate,TammCatch "
                PrnLine &= String.Format("{0,3}", Area.ToString)
                PrnLine &= String.Format("{0,3}", TStep.ToString)
                PrnLine &= String.Format("{0,10}", TammEstimate(Area, TStep).ToString("######0.0"))
                PrnLine &= String.Format("{0,10}", TammCatch(Area, TStep).ToString("######0.0"))
                sw.WriteLine(PrnLine)
            Next Area
        Next TStep

    End Sub

    Sub CHKSPRCH(ByVal TStep, ByVal Fish1, ByVal Fish2, ByVal Fish3, ByVal Stk)
        Dim TammDiff As Integer
        Dim TamkSHRS As Double

        If TammPSER(1, 2) = -88 Or TammPSER(1, 3) = -88 Or TammPSER(2, 2) = -88 Or TammPSER(2, 3) = -88 Then Exit Sub
        If TammCatch(Fish1, TStep) <> TammEstimate(Fish1, TStep) Then
            '   ReDim NewSHRS(NumSteps)
            TammDiff = TammCatch(Fish1, TStep) - TammEstimate(Fish1, TStep)
            If Abs(TammDiff) > 1.0 Then
                TammChinookConverge = 1
                TamkSHRS = StockFishRateScalers(Stk, Fish2, TStep)
                If TamkSHRS = 0.0 Then
                    TamkSHRS = 1.0
                Else
                    If StockFishRateScalers(Stk, Fish1, TStep) = 0.0 Then
                        TamkSHRS = 0.0
                    Else
                        If TammCatch(Fish1, TStep) <> 0.0 Then
                            TamkSHRS = StockFishRateScalers(Stk, Fish2, TStep) * (TammEstimate(Fish1, TStep) / TammCatch(Fish1, TStep))
                        End If
                    End If
                End If
                StockFishRateScalers(Stk, Fish2, TStep) = TamkSHRS
                If TamkSHRS <> 1 Then
                    NewStockFishRateScalers = 1
                End If
                If Stk = 2 Then  '- SF Nooksack Spring
                    StockFishRateScalers(3, Fish2, TStep) = TamkSHRS
                    If TamkSHRS <> 1 Then
                        NewStockFishRateScalers = 1
                    End If
                End If
                TammCatch(Fish3, TStep) = TammCatch(Fish3, TStep) + TammDiff
            End If
        End If

    End Sub

    Sub CHKSPRCHSF(ByVal TStep, ByVal Fish1, ByVal Fish2, ByVal Fish3, ByVal Stk)

        Dim StkLp, TammDiff As Integer
        Dim TamkSHRS As Double

        If TammPSER(1, 2) = -88 Or TammPSER(1, 3) = -88 Or TammPSER(2, 2) = -88 Or TammPSER(2, 3) = -88 Then Exit Sub

        '- Changed Stk for SF Chin FRAM 9/20/2002
        If TammCatch(Fish1, TStep) <> TammEstimate(Fish1, TStep) Then
            TammDiff = TammCatch(Fish1, TStep) - TammEstimate(Fish1, TStep)
            If Abs(TammDiff) > 1.0 Then
                TammChinookConverge = 1
                For StkLp = 1 To 2
                    TamkSHRS = StockFishRateScalers(Stk * 2 - StkLp + 1, Fish2, TStep)
                    If TamkSHRS = 0.0 Then
                        TamkSHRS = 1.0
                    Else
                        If StockFishRateScalers(Stk * 2 - StkLp + 1, Fish1, TStep) = 0.0 Then
                            TamkSHRS = 0.0
                        Else
                            If TammCatch(Fish1, TStep) <> 0.0 Then
                                TamkSHRS = StockFishRateScalers(Stk * 2 - StkLp + 1, Fish2, TStep) * (TammEstimate(Fish1, TStep) / TammCatch(Fish1, TStep))
                            End If
                        End If
                    End If
                    StockFishRateScalers(Stk * 2 - StkLp + 1, Fish2, TStep) = TamkSHRS
                    If TamkSHRS <> 1 Then
                        NewStockFishRateScalers = 1
                    End If
                    If Stk = 2 Then  '- SF Nooksack Spring
                        StockFishRateScalers(StkLp + 4, Fish2, TStep) = TamkSHRS
                        If TamkSHRS <> 1 Then
                            NewStockFishRateScalers = 1
                        End If
                    End If
                    TammCatch(Fish3, TStep) = TammCatch(Fish3, TStep) + TammDiff
                Next StkLp
            End If
        End If

    End Sub

    Sub TCHNInit()
        Dim Area As Integer

        '---------- Initialization for Chinook TAMM Processing
        '---------- Values are very specific for CHINOOK.OUT file with
        '---------- Wash. State Treaty/Non-Treaty fisheries

        PrnLine = "Bham Bay NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(39, TStep).ToString(" ####0.0000") & FisheryFlag(39, TStep).ToString(" 0") & TammScaler(1, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Bham Bay TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(40, TStep).ToString(" ####0.0000") & FisheryFlag(40, TStep).ToString(" 0") & TammScaler(2, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8   NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(46, TStep).ToString(" ####0.0000") & FisheryFlag(46, TStep).ToString(" 0") & TammScaler(3, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8   TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(47, TStep).ToString(" ####0.0000") & FisheryFlag(47, TStep).ToString(" 0") & TammScaler(4, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8A  NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(49, TStep).ToString(" ####0.0000") & FisheryFlag(49, TStep).ToString(" 0") & TammScaler(5, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8A  TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(50, TStep).ToString(" ####0.0000") & FisheryFlag(50, TStep).ToString(" 0") & TammScaler(6, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8D  NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(51, TStep).ToString(" ####0.0000") & FisheryFlag(51, TStep).ToString(" 0") & TammScaler(7, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 8D  TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(52, TStep).ToString(" ####0.0000") & FisheryFlag(52, TStep).ToString(" 0") & TammScaler(8, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 12  NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(65, TStep).ToString(" ####0.0000") & FisheryFlag(65, TStep).ToString(" 0") & TammScaler(9, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 12  TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(66, TStep).ToString(" ####0.0000") & FisheryFlag(66, TStep).ToString(" 0") & TammScaler(10, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 13A NNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(70, TStep).ToString(" ####0.0000") & FisheryFlag(70, TStep).ToString(" 0") & TammScaler(11, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)
        PrnLine = "Area 13A TNET - TStep,QEff,qf,QScale = " & TStep.ToString & FisheryScaler(71, TStep).ToString(" ####0.0000") & FisheryFlag(71, TStep).ToString(" 0") & TammScaler(12, TStep).ToString(" ####0.0000")
        sw.WriteLine(PrnLine)

        '----------------- ReSet Effort Using TAMM Scalar ---
        For Area = 1 To 12
            If FisheryFlag(TamkFish(Area), TStep) = 1 Then
                FisheryScaler(TamkFish(Area), TStep) = FisheryScaler(TamkFish(Area), TStep) * TammScaler(Area, TStep)
            ElseIf FisheryFlag(TamkFish(Area), TStep) = 2 Then
                FisheryQuota(TamkFish(Area), TStep) = FisheryQuota(TamkFish(Area), TStep) * TammScaler(Area, TStep)
            End If
            For Stk = 1 To NumStk
                For Age = MinAge To MaxAge
                    Shakers(Stk, Age, TamkFish(Area), TStep) = 0
                    NonRetention(Stk, Age, TamkFish(Area), TStep) = 0
                    MSFShakers(Stk, Age, TamkFish(Area), TStep) = 0
                    MSFNonRetention(Stk, Age, TamkFish(Area), TStep) = 0
                Next Age
            Next Stk
        Next Area

    End Sub

    Sub TammChinookProc()       'CHINOOK TAMM Main Processing
        '**************************************************************************
        ' TAMM Variable Processing - Iteratively Solve for FRAM Effort Scalars that
        '                            equal the BaseExploitationRate catch from 
        '                            PS Net Estimates
        '**************************************************************************

        '- Define CHINOOK TAMM Terminal Fisheries
        ReDim TamkFish(12)
        TamkFish(1) = 39         '--- Nooksack-Samish 7BCD Net
        TamkFish(2) = 40
        TamkFish(3) = 46         '--- Skagit Bay 8 Net
        TamkFish(4) = 47
        TamkFish(5) = 49         '--- Still/Snohomish 8A Net
        TamkFish(6) = 50
        TamkFish(7) = 51         '--- Tulalip Bay 8D Net
        TamkFish(8) = 52
        TamkFish(9) = 65         '--- Hood Canal 12-12D Net
        TamkFish(10) = 66
        TamkFish(11) = 70        '--- Carr Inlet 13A Net
        TamkFish(12) = 71

        If TammChinookRunFlag > 1 Then
            '- Special Run Options
            If NumStk > 50 Then
                Call TCHNSFTran()
            Else
                Call TCHNTran()
            End If
            Exit Sub
        End If

        For TammIteration = 1 To 15

            If TammIteration = 5 Then
                Jim = 1
            End If

            If NumStk < 50 Then
                Call TCHNComp(TammIteration)
            Else
                Call TCHNSFComp(TammIteration)
            End If
            If TammChinookConverge = 0 Then
                Exit For
            End If
            PrnLine = "-------- TAMM Iter= " & TammIteration.ToString & " ------------------------"
            sw.WriteLine(PrnLine)
            FVS_RunModel.RunProgressLabel.Text = "TAMM Iteration - " & TammIteration.ToString
            FVS_RunModel.RunProgressLabel.Refresh()
            For TStep = 2 To 3
                If TStep = 3 Then Jim = 1
                Call TCHNInit()
                Call CompCatch(Term)
                For Fish = 1 To NumFish
                    If TerminalFisheryFlag(Fish, TStep) = Term Then
                        Call CompOthMort(Fish)
                    End If
                Next Fish
                Call CompEscape()
            Next TStep
        Next TammIteration

        PrnLine = "Total Number of TAMM Iterations = " & (TammIteration - 1).ToString
        sw.WriteLine(PrnLine)

        If TammIteration > 15 Then    '---- Did NOT Converge .... Print Error in FRAMMENU
            RunTAMMIter = -1
            Exit Sub
        Else
            RunTAMMIter = TammIteration - 1
            If TammTransferSave = False Then Exit Sub
            FVS_RunModel.RunProgressLabel.Text = "Transferring Data to TAMM SpreadSheet "
            FVS_RunModel.RunProgressLabel.Refresh()
            If NumStk > 50 Then
                Call TCHNSFTran()
            Else
                Call TCHNTran()
            End If
        End If

    End Sub

    Sub TCHNTran()
        '------------------ TAMM Transfer File for CHINOOK ---

        Dim StkNum As Integer
        Dim USPS0, UWACC, DSPS0, SPSYR, NONSS, Sps1011, SSETAC As Double
        Dim USSETRS, DSSETRS, SUMETRS As Double
        Dim TotalChinEsc(3, 22) As Double
        Dim TermChinAbun(7) As Double

        '- Terminal Area Escapements ---

        For TStep = 2 To NumSteps
            For Stk = 1 To 18
                If Stk > 2 Then
                    StkNum = Stk - 1
                Else
                    StkNum = Stk
                End If
                For Age = MinAge To MaxAge
                    If Age = 2 Then
                        TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + Escape(Stk, Age, TStep)
                    Else
                        TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + Escape(Stk, Age, TStep)
                    End If
                    If StkNum = 13 Then  '--- sps yearling split 'not used for TAMX
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * Escape(Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * Escape(Stk, Age, TStep)
                            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * Escape(Stk, Age, TStep)
                        End If
                    End If
                    If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + Escape(Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + Escape(Stk, Age, TStep)
                        End If
                    End If
                    If StkNum = 12 Then               '--- Deep SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + Escape(Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + Escape(Stk, Age, TStep)
                        End If
                    End If
                Next Age
            Next Stk
        Next TStep

        'ADD IN FRESHWATER NET and Sport CATCH TO GET EXTREME TERMINAL RUN

        For Fish = 72 To 73
            For Stk = 1 To 18
                If Stk > 2 Then
                    StkNum = Stk - 1
                Else
                    StkNum = Stk
                End If
                For TStep = 2 To NumSteps
                    For Age = MinAge To MaxAge
                        If Age = 2 Then
                            TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + LandedCatch(Stk, Age, Fish, TStep)
                        Else
                            TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + LandedCatch(Stk, Age, Fish, TStep)
                        End If
                        If StkNum = 13 Then  '---- sps yearling split
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            End If
                        End If
                        If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            End If
                        End If
                        If StkNum = 12 Then               '--- Deep SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Else
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            End If
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- New Section for TAA for 7B, 8, 8A, 10, and 12 plus TRS
        For Stk = 1 To 21
            TotalChinEsc(3, Stk) = TotalChinEsc(2, Stk) '--- Start with Age 3 ETRS local stock
        Next Stk

        '-------------------------------- NkSam TAA
        TermChinAbun(1) = TotalChinEsc(3, 1)
        TStep = 3 '- Only Time 3 by definition
        For Fish = 39 To 40  '---- B'Ham Bay Net 7B
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk = 1 Then
                        TotalChinEsc(3, 1) = TotalChinEsc(3, 1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        '--- TRS
                    End If
                    If Stk = 2 Or Stk = 3 Then
                        TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        '--- TRS
                    End If
                    TermChinAbun(1) = TermChinAbun(1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    '---------- TAA
                Next Age
            Next Stk
        Next Fish

        TStep = 2
        '--- B'Ham Bay Net 7B Nooksack Spring Chinook time step 2
        For Stk = 2 To 3
            For Fish = 39 To 40
                For Age = MinAge To MaxAge   '---- All Ages in ETRS marine catches
                    TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                Next Age
            Next Fish
        Next Stk

        '-------------------------------- Skagit TAA
        TermChinAbun(2) = TotalChinEsc(3, 3) + TotalChinEsc(3, 4)
        For Fish = 46 To 47  '--- Skagit Bay Net
            For Stk = 1 To 32
                For TStep = 2 To 3          '---- only Step 2 and 3 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk = 4 Or Stk = 5 Or Stk = 6 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            '--- TRS Falls and Springs
                        End If
                        TermChinAbun(2) = TermChinAbun(2) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        '---------- TAA
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '------------------------------------ Still/Snohomish 8A TAA
        TermChinAbun(3) = TotalChinEsc(3, 6) + TotalChinEsc(3, 7) + TotalChinEsc(3, 8) + TotalChinEsc(3, 9)
        TStep = 3 '- by definition
        For Fish = 49 To 50   '---- Area 8A Net
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk >= 7 And Stk <= 9 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        '--- Tulalip uses ETRS ... Don't add 8A Catch for Stock #10  3/24/99
                    End If
                    TermChinAbun(3) = TermChinAbun(3) + LandedCatch(Stk, Age, Fish, TStep)
                Next Age
            Next Stk
        Next Fish

        For Fish = 51 To 52     '--- Tulalip Bay Net
            For Stk = 1 To 32
                For TStep = 3 To 4 '---- only Step 3 and 4 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk >= 7 And Stk <= 9 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        End If
                        TermChinAbun(3) = TermChinAbun(3) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                            TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        Else
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '-------------------------------------------- South Sound TAA
        TermChinAbun(4) = TotalChinEsc(3, 10) + TotalChinEsc(3, 11) + TotalChinEsc(3, 12) + TotalChinEsc(3, 13)
        TStep = 3 '- by definition
        Sps1011 = 0.0
        USPS0 = 0.0
        UWACC = 0.0
        DSPS0 = 0.0
        SPSYR = 0.0
        NONSS = 0.0
        For Fish = 58 To 71
            If Fish > 63 And Fish < 68 Then GoTo NotFish
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS catches
                    TermChinAbun(4) = TermChinAbun(4) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    '---------- TAA
                    '-- 10/11 and 13 catches TRS for FRAM SS Stocks
                    '   both Falls and 13A Springs
                    If ((Stk > 10 And Stk < 16) Or Stk = 33) And (Fish <= 59 Or Fish = 68 Or Fish = 69) Then
                        Select Case Stk
                            Case 11
                                TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 12
                                TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 13
                                TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 14
                                TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 15  '--- WhRvrSpr Fing
                                TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                                'Case 33  '--- WhRvrSpr Fing
                                '   TotalChinEsc(3, 22) = TotalChinEsc(3, 22) + LandedCatch(Stk, Age, Fish, TStep)
                        End Select
                    End If
                    If Fish <= 59 Then       '--- 10/11 Net Catches for Split TAA
                        Sps1011 = Sps1011 + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    End If
                    '------ 10A,10E,13A,SPS Net ETRS Catches
                    If (Fish >= 60 And Fish <= 63) Or (Fish >= 65 And Fish <= 68) Then
                        Select Case Stk
                            Case 11
                                USPS0 = USPS0 + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 12
                                UWACC = UWACC + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 13
                                DSPS0 = DSPS0 + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case 14
                                SPSYR = SPSYR + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            Case Else
                                NONSS = NONSS + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        End Select
                    End If
                    'If Stk = 15 And Fish >= 70 Then   '--- Spring Yearling
                    '   TotalChinEsc(2, 14) = TotalChinEsc(2, 14) + LandedCatch(Stk, Age, Fish, TStep) '--- ETRS
                    '   TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + LandedCatch(Stk, Age, Fish, TStep) '--- ETRS
                    'End If
                Next Age
            Next Stk
NotFish:
        Next Fish

        SSETAC = USPS0 + UWACC + DSPS0 + SPSYR
        If SSETAC <> 0.0 Then
            '- Changed SS ETRS to NOT include Non-Local Stocks Feb 2011
            'TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0 + (NONSS * (USPS0 / SSETAC))
            'TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACC + (NONSS * (UWACC / SSETAC))
            'TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0 + (NONSS * (DSPS0 / SSETAC))
            'TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYR + (NONSS * (SPSYR / SSETAC))
            'TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0 + (NONSS * (USPS0 / SSETAC))
            'TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACC + (NONSS * (UWACC / SSETAC))
            'TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0 + (NONSS * (DSPS0 / SSETAC))
            'TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYR + (NONSS * (SPSYR / SSETAC))

            'TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0 + (NONSS * (USPS0 / SSETAC)) + _
            '   UWACC + (NONSS * (UWACC / SSETAC)) + _
            '   (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            'TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0 + (NONSS * (DSPS0 / SSETAC)) + _
            '   (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            'TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0 + (NONSS * (USPS0 / SSETAC)) + _
            '   UWACC + (NONSS * (UWACC / SSETAC)) + _
            '   (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            'TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0 + (NONSS * (DSPS0 / SSETAC)) + _
            '   (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            'TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0
            TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACC
            TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0
            TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYR
            TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0
            TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACC
            TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0
            TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYR

            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0 + UWACC + SPSYR * SpsYrSpl
            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0 + SPSYR * (1.0 - SpsYrSpl)
            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + SPSYR * SpsYrSpl
            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + SPSYR * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0 + UWACC + SPSYR * SpsYrSpl
            TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0 + SPSYR * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + SPSYR * SpsYrSpl
            TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + SPSYR * (1.0 - SpsYrSpl)
        End If

        '------ Area 10/11 Net LandedCatch Split between Upper and Deep South Sound TAA
        USSETRS = TotalChinEsc(3, 18)
        DSSETRS = TotalChinEsc(3, 19)
        SUMETRS = USSETRS + DSSETRS
        If SUMETRS <> 0.0 Then
            TermChinAbun(6) = USSETRS + ((USSETRS / SUMETRS) * Sps1011)
            TermChinAbun(7) = DSSETRS + ((DSSETRS / SUMETRS) * Sps1011)
        End If

        '-------------------------------------------- Hood Canal TAA
        TermChinAbun(5) = TotalChinEsc(3, 15) + TotalChinEsc(3, 16)
        TStep = 3 '- by definition
        For Fish = 65 To 66  '--- HC Net
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS
                    If Stk >= 16 And Stk <= 17 Then
                        '- TRS
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    End If
                    '- TAA
                    TermChinAbun(5) = TermChinAbun(5) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                Next Age
            Next Stk
        Next Fish

        '--- Subtract TAMM FW Sport and Add Marine Sport Savings to TAA's
        '- PS Chinook Run Reconstruction now includes FW Sport ... 1/3/2011 JFP
        'TermChinAbun(1) = TermChinAbun(1) - TNkFWSpt! + TNkMSA!
        'TermChinAbun(2) = TermChinAbun(2) - TSkFWSpt! + TSkMSA!
        'TermChinAbun(3) = TermChinAbun(3) - TSnFWSpt! + TSnMSA!
        'TermChinAbun(4) = TermChinAbun(4)
        'TermChinAbun(5) = TermChinAbun(5)
        'TermChinAbun(6) = TermChinAbun(6)
        'TermChinAbun(7) = TermChinAbun(7)

        '------- Transfer Data to TAMX Worksheet ---

        xlWorkSheet = xlWorkBook.Sheets("TAMX")
        xlWorkSheet.Range("B1").Value = FramVersion
        xlWorkSheet.Range("B2").Value = RunIDNameSelect
        xlWorkSheet.Range("E1").Value = "Date:"
        xlWorkSheet.Range("F1").Value = DateTime.Now.ToString

        Dim TamxLineNum(19) As Integer
        TamxLineNum(0) = 0
        TamxLineNum(1) = 5
        TamxLineNum(2) = 7
        TamxLineNum(3) = 8
        TamxLineNum(4) = 9
        TamxLineNum(5) = 11
        TamxLineNum(6) = 12
        TamxLineNum(7) = 13
        TamxLineNum(8) = 14
        TamxLineNum(9) = 15
        TamxLineNum(10) = 17
        TamxLineNum(11) = 18
        TamxLineNum(12) = 19
        TamxLineNum(13) = 20
        TamxLineNum(14) = 26
        TamxLineNum(15) = 27
        TamxLineNum(16) = 28
        TamxLineNum(17) = 30
        TamxLineNum(18) = 31
        TamxLineNum(19) = 32

        '----- Put Terminal and Extreme Terminal Run Sizes into TAMX WorkSheet---
        Dim RngVal1 As String
        For Stk = 1 To 17
            RngVal1 = "B" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, Stk).ToString("######0")
            RngVal1 = "C" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, Stk).ToString("######0")
            RngVal1 = "D" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, Stk).ToString("######0")
            '- Put Stock Specific Terminal Run Sizes into WorkSheet -
            Select Case Stk
                Case 1
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(1).ToString("######0")
                Case 4
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(2).ToString("######0")
                Case 9
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(3).ToString("######0")
                Case 13
                    '--- Upper South Sound Yr.
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 20).ToString("######0")
                    RngVal1 = "C" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 20).ToString("######0")
                    RngVal1 = "D" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 20).ToString("######0")
                    '--- Deep South Sound Yr.
                    RngVal1 = "B" & (TamxLineNum(Stk) + 2).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 21).ToString("######0")
                    RngVal1 = "C" & (TamxLineNum(Stk) + 2).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 21).ToString("######0")
                    RngVal1 = "D" & (TamxLineNum(Stk) + 2).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 21).ToString("######0")
                    '--- Upper South Sound Agg.
                    RngVal1 = "B" & (TamxLineNum(Stk) + 3).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(6).ToString("######0")
                    RngVal1 = "C" & (TamxLineNum(Stk) + 3).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 18).ToString("######0")
                    RngVal1 = "D" & (TamxLineNum(Stk) + 3).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 18).ToString("######0")
                    '--- Deep South Sound Agg.
                    RngVal1 = "B" & (TamxLineNum(Stk) + 4).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(7).ToString("######0")
                    RngVal1 = "C" & (TamxLineNum(Stk) + 4).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 19).ToString("######0")
                    RngVal1 = "D" & (TamxLineNum(Stk) + 4).ToString
                    xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 19).ToString("######0")
                    '--- Total TAA
                    RngVal1 = "B" & (TamxLineNum(Stk) + 5).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(4).ToString("######0")
                Case 16
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(5).ToString("######0")
                Case Else
            End Select
        Next
        '- Add White River Springs to List (Sum Both Components)
        Stk = 18
        RngVal1 = "B" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 22).ToString("######0")
        RngVal1 = "C" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 22).ToString("######0")
        RngVal1 = "D" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 22).ToString("######0")
        '- Note: Hoko Not Used in Older SpreadSheets
        ''- Add Hoko to Bottom of List (Sum Both Components)
        'Stk = 19
        'RngVal1 = "B" & (TamxLineNum(Stk)).ToString
        'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 23).ToString("######0")
        'RngVal1 = "C" & (TamxLineNum(Stk)).ToString
        'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 23).ToString("######0")
        'RngVal1 = "D" & (TamxLineNum(Stk)).ToString
        'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 23).ToString("######0")

        '--- GET Fishery Landed Catch and Total Mortality DATA ---

        '- Dimension Transfer Arrays to Accomodate WorkSheet Transfer
        Dim FishVal As Integer
        Dim LandCatch(1, 1)  '- Landed Catch
        Dim TotalMort(1, 1)  '- Total Mortality
        If TammChinookRunFlag = 1 Then
            ReDim LandCatch(NumFish - 3, NumSteps)
            ReDim TotalMort(NumFish - 3, NumSteps)
        Else
            ReDim LandCatch(NumFish - 2, NumSteps)
            ReDim TotalMort(NumFish - 2, NumSteps)
        End If
        For Fish = 1 To NumFish
            '- Determine Fishery Numbers consistent with TAMM SpreadSheet
            If TammChinookRunFlag = 1 Then
                Select Case Fish
                    Case 1 To 12
                        FishVal = Fish
                    Case 13 To 15
                        FishVal = 13
                    Case 16 To 55
                        FishVal = Fish - 2
                    Case 56 To 57
                        FishVal = 54
                    Case 58 To NumFish
                        FishVal = Fish - 3
                End Select
            Else
                Select Case Fish
                    Case 1 To 12
                        FishVal = Fish
                    Case 13 To 15
                        FishVal = 13
                    Case 16 To NumFish
                        FishVal = Fish - 2
                End Select
            End If
            For TStep = 2 To NumSteps
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        LandCatch(FishVal - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        TotalMort(FishVal - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        LandCatch(FishVal - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        TotalMort(FishVal - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                    Next
                Next
            Next
        Next
        '- Convert Arrays into Long Integer Values for WorkSheet
        If TammChinookRunFlag = 1 Then
            For Fish = 0 To NumFish - 3
                For TStep = 0 To NumSteps
                    LandCatch(Fish, TStep) = CLng(LandCatch(Fish, TStep))
                    TotalMort(Fish, TStep) = CLng(TotalMort(Fish, TStep))
                Next
            Next
        Else
            For Fish = 0 To NumFish - 2
                For TStep = 0 To NumSteps
                    LandCatch(Fish, TStep) = CLng(LandCatch(Fish, TStep))
                    TotalMort(Fish, TStep) = CLng(TotalMort(Fish, TStep))
                Next
            Next
        End If
        '- Put TAMX arrays into WorkSheet
        If TammChinookRunFlag = 1 Then
            xlWorkSheet.Range("B42").Resize(NumFish - 3, NumSteps + 1).Value = TotalMort
            xlWorkSheet.Range("P42").Resize(NumFish - 3, NumSteps + 1).Value = LandCatch
        Else
            xlWorkSheet.Range("B42").Resize(NumFish - 2, NumSteps + 1).Value = TotalMort
            xlWorkSheet.Range("P42").Resize(NumFish - 2, NumSteps + 1).Value = LandCatch
        End If

        '-------------------------- STOCK CATCH BY FISHERY ---

        Dim FishNum, StkVal As Integer
        For Stk = 1 To 18
            If TammChinookRunFlag = 1 Then
                ReDim LandCatch(NumFish - 3, NumSteps)
                ReDim TotalMort(NumFish - 3, NumSteps)
            Else
                ReDim LandCatch(NumFish - 2, NumSteps)
                ReDim TotalMort(NumFish - 2, NumSteps)
            End If
            '- Determine Stock Numbers and Sequence Values
            '- WhRvr Spring Yearling and Hoko Added (#33 & # 38)
            If Stk = 1 Then
                StkNum = 1
                StkVal = 1
            ElseIf Stk = 2 Then
                StkNum = 2
                StkVal = 2
            ElseIf Stk > 2 And Stk < 18 Then
                StkNum = Stk + 1
                StkVal = Stk
            ElseIf Stk = 18 Then
                StkNum = 33 '- WhRvr Spr Year
                StkVal = 18
                '- Note: Hoko Not Used in Older SpreadSheets
                'ElseIf Stk = 20 Then
                '   StkNum = 38 '- Hoko
                '   StkVal = 19
            End If
NooksackSpringReEntry:
            For Fish = 1 To NumFish
                '- Combined Fisheries
                If TammChinookRunFlag = 1 Then
                    '- Old Style Format (Area 10/11 Sport Combined)
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 55
                            FishNum = Fish - 2
                        Case 56, 57
                            FishNum = 54
                        Case 58 To 73
                            FishNum = Fish - 3
                    End Select
                Else
                    '- Current Chinook TAMM Transfer
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 73
                            FishNum = Fish - 2
                    End Select
                End If
                For TStep = 2 To NumSteps
                    For Age = MinAge To MaxAge
                        '- AEQ Value NOT used for Total Mortality in Terminal Fisheries
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                            TotalMort(FishNum - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                            TotalMort(FishNum - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                        Else
                            TotalMort(FishNum - 1, TStep - 1) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(StkNum, Age, TStep))
                            TotalMort(FishNum - 1, NumSteps) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(StkNum, Age, TStep))
                        End If
                        '- Landed Catch Arrays
                        LandCatch(FishNum - 1, TStep - 1) += (LandedCatch(StkNum, Age, Fish, TStep) + MSFLandedCatch(StkNum, Age, Fish, TStep))
                        LandCatch(FishNum - 1, NumSteps) += (LandedCatch(StkNum, Age, Fish, TStep) + MSFLandedCatch(StkNum, Age, Fish, TStep))
                    Next Age
                Next TStep
            Next Fish
            '- Test for Nooksack Spring Chinook (2 Stocks now)
            If Stk = 2 And StkNum = 2 Then
                StkNum = 3
                GoTo NooksackSpringReEntry
            End If
            '- Convert Arrays into Long Integer Values for WorkSheet
            If TammChinookRunFlag = 1 Then
                For Fish = 0 To NumFish - 3
                    For TStep = 0 To NumSteps
                        LandCatch(Fish, TStep) = CLng(LandCatch(Fish, TStep))
                        TotalMort(Fish, TStep) = CLng(TotalMort(Fish, TStep))
                    Next
                Next
            Else
                For Fish = 0 To NumFish - 2
                    For TStep = 0 To NumSteps
                        LandCatch(Fish, TStep) = CLng(LandCatch(Fish, TStep))
                        TotalMort(Fish, TStep) = CLng(TotalMort(Fish, TStep))
                    Next
                Next
            End If
            '- Put Stock Catch and Stock TotalMort Arrays into WorkSheet
            If TammChinookRunFlag = 1 Then
                RngVal1 = "B" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = TotalMort
                RngVal1 = "P" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = LandCatch
            Else
                RngVal1 = "B" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = TotalMort
                RngVal1 = "P" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = LandCatch
            End If
        Next

        '- Save WorkBook and Close Application if Necessary
        'xlApp.Application.DisplayAlerts = False
        'xlWorkBook.Save()
        'If WorkBookWasNotOpen = True Then
        '   xlWorkBook.Close()
        'End If
        'If ExcelWasNotRunning = True Then
        '   xlApp.Application.Quit()
        '   xlApp.Quit()
        'Else
        '   xlApp.Visible = True
        '   xlApp.WindowState = Excel.XlWindowState.xlMinimized
        'End If
        xlApp.Visible = True
        xlApp.Application.DisplayAlerts = True
        'xlApp = Nothing

    End Sub

    Sub TCHNSFTran()

        '--- TAMM Transfer File for Selective Fishery Style CHINOOK ---

        Dim PSStocks, AllStocks, SomeStocks As Integer
        Dim USPS0u, UWACCu, DSPS0u, SPSYRu, NONSSu, Sps1011u, SSETACu As Double
        Dim USPS0m, UWACCm, DSPS0m, SPSYRm, NONSSm, Sps1011m, SSETACm As Double
        Dim USSETRSu, DSSETRSu, SUMETRSu As Double
        Dim USSETRSm, DSSETRSm, SUMETRSm As Double
        Dim StkNum, StkVal As Integer
        Dim TotalChinEsc(6, 23) As Double
        Dim SptSave(14) As Double
        Dim TermChinAbun(14) As Double

        '- Test to include Elwha in TAMX
        If NumStk = 38 Or NumStk = 76 Then
            PSStocks = 20
            AllStocks = 38
            SomeStocks = 23
        ElseIf NumStk = 33 Or NumStk = 66 Then
            PSStocks = 19
            AllStocks = 33
            SomeStocks = 22
        Else
            PSStocks = 20
            AllStocks = NumStk / 2
            SomeStocks = 23
        End If







        '--- 18 stocks plus SPS Yearling splits Plus WhRvr Sprg and Hoko

        '------------------------ Terminal Area Escapements ---
        For TStep = 2 To NumSteps
            For Stk = 1 To PSStocks



                If Stk > 2 And Stk < 19 Then
                    StkNum = Stk - 1
                    StkVal = Stk
                Else
                    If Stk = 19 Then
                        StkVal = 33 '- WhRvr Spr Year
                        StkNum = 22
                    ElseIf Stk = 20 Then
                        StkVal = 38 '- Hoko
                        StkNum = 23
                    Else
                        StkVal = Stk
                        StkNum = Stk
                    End If
                End If

               

                For Age = MinAge To MaxAge
                    If Age = 2 Then
                        TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + Escape(StkVal * 2 - 1, Age, TStep)
                        TotalChinEsc(4, StkNum) = TotalChinEsc(4, StkNum) + Escape(StkVal * 2, Age, TStep)
                    Else
                        TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + Escape(StkVal * 2 - 1, Age, TStep)
                        TotalChinEsc(5, StkNum) = TotalChinEsc(5, StkNum) + Escape(StkVal * 2, Age, TStep)
                    End If
                    'If StkNum = 13 Then  '--- sps yearling split
                    '    If Age = 2 Then
                    '        TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + (SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + (1.0 - SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(4, 20) = TotalChinEsc(4, 20) + (SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(4, 21) = TotalChinEsc(4, 21) + (1.0 - SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '    Else
                    '        TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * Escape(StkVal * 2 - 1, Age, TStep)
                    '        TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + (SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + (1.0 - SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '        TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (1.0 - SpsYrSpl) * Escape(StkVal * 2, Age, TStep)
                    '    End If
                    'End If
                    'If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                    '    If Age = 2 Then
                    '        TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + Escape(Stk * 2 - 1, Age, TStep)
                    '        TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + Escape(Stk * 2, Age, TStep)
                    '    Else
                    '        TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + Escape(Stk * 2 - 1, Age, TStep)
                    '        TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + Escape(Stk * 2, Age, TStep)
                    '    End If
                    'End If
                    'If StkNum = 12 Then               '--- Deep SPS
                    '    If Age = 2 Then
                    '        TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + Escape(Stk * 2 - 1, Age, TStep)
                    '        TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + Escape(Stk * 2, Age, TStep)
                    '    Else
                    '        TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + Escape(Stk * 2 - 1, Age, TStep)
                    '        TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + Escape(Stk * 2, Age, TStep)
                    '    End If
                    'End If
                Next Age
            Next Stk
        Next TStep
        '*****************************************
        'AHB 4/2/18
        'Save Nooksack NOR when modeled in North Fork slot (stk = 2), before they get combined with South Fork (or hatchery); FW net and sport = 0
        Dim NookSprETRS_NOR As Double
        For Age = 3 To 5
            For TStep = 2 To 3
                NookSprETRS_NOR += Escape(3, Age, TStep) 'UM Nook NOR
            Next TStep
        Next Age
        '******************************************

        'ADD IN FRESHWATER NET and Sport CATCH TO GET EXTREME TERMINAL RUN

        For Fish = 72 To 73
            For Stk = 1 To PSStocks
                If Stk > 2 And Stk < 19 Then
                    StkNum = Stk - 1
                    StkVal = Stk
                Else
                    If Stk = 19 Then
                        StkVal = 33 '- WhRvr Spr Year
                        StkNum = 22
                    ElseIf Stk = 20 Then
                        StkVal = 38 '- Hoko
                        StkNum = 23
                    Else
                        StkVal = Stk
                        StkNum = Stk
                    End If
                End If
                For TStep = 2 To NumSteps
                    For Age = MinAge To MaxAge
                        If Age = 2 Then
                            TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(4, StkNum) = TotalChinEsc(4, StkNum) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        Else
                            TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(5, StkNum) = TotalChinEsc(5, StkNum) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        End If
                        'If StkNum = 13 Then  '---- sps yearling split
                        '    If Age = 2 Then
                        '        TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + (SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(4, 20) = TotalChinEsc(4, 20) + (SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(4, 21) = TotalChinEsc(4, 21) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    Else
                        '        TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + (SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '        TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (1.0 - SpsYrSpl) * (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    End If
                        'End If
                        'If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                        '    If Age = 2 Then
                        '        TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    Else
                        '        TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    End If
                        'End If
                        'If StkNum = 12 Then               '--- Deep SPS
                        '    If Age = 2 Then
                        '        TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    Else
                        '        TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (LandedCatch(StkVal * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2 - 1, Age, Fish, TStep))
                        '        TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + (LandedCatch(StkVal * 2, Age, Fish, TStep) + MSFLandedCatch(StkVal * 2, Age, Fish, TStep))
                        '    End If
                        'End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- New Section for TAA for 7B, 8, 8A, 10, and 12 plus TRS
        For Stk = 1 To SomeStocks
            TotalChinEsc(3, Stk) = TotalChinEsc(2, Stk) '--- Start with Age 3 ETRS local stock
            TotalChinEsc(6, Stk) = TotalChinEsc(5, Stk) '--- Start with Age 3 ETRS local stock
        Next Stk

        '-------------------------------- NkSam TAA
        TermChinAbun(1) = TotalChinEsc(3, 1)
        TermChinAbun(8) = TotalChinEsc(6, 1)
        TStep = 3 '- Only Time 3 by definition
        For Fish = 39 To 40  '---- B'Ham Bay Net 7B
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk = 1 Then 'NookSam SF
                        '- TRS
                        TotalChinEsc(3, 1) = TotalChinEsc(3, 1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TotalChinEsc(6, 1) = TotalChinEsc(6, 1) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                    End If
                    If Stk = 2 Or Stk = 3 Then 'NookSpr
                        '- TRS
                        TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TotalChinEsc(6, 2) = TotalChinEsc(6, 2) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                    End If
                    '- TAA
                    TermChinAbun(1) = TermChinAbun(1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                    TermChinAbun(8) = TermChinAbun(8) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                Next Age
            Next Stk
        Next Fish

        '--- B'Ham Bay Net 7B Nooksack Spring Chinook time step 2
        TStep = 2
        For Stk = 2 To 3
            For Fish = 39 To 40
                For Age = MinAge To MaxAge   '---- All Ages in ETRS marine catches
                    TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                    TotalChinEsc(6, 2) = TotalChinEsc(6, 2) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                Next Age
            Next Fish
        Next Stk

        '--- Skagit TAA
        TermChinAbun(2) = TotalChinEsc(3, 3) + TotalChinEsc(3, 4)
        TermChinAbun(9) = TotalChinEsc(6, 3) + TotalChinEsc(6, 4)
        For Fish = 46 To 47  '--- Skagit Bay Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For TStep = 2 To 3          '---- only Step 2 and 3 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk = 4 Or Stk = 5 Or Stk = 6 Then ' Skag SF Fing, Skag SF Yrl, Skag Spring
                            '- TRS Falls and Springs
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        End If
                        '- TAA
                        TermChinAbun(2) = TermChinAbun(2) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TermChinAbun(9) = TermChinAbun(9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- Still/Snohomish 8A TAA
        TermChinAbun(3) = TotalChinEsc(3, 6) + TotalChinEsc(3, 7) + TotalChinEsc(3, 8) + TotalChinEsc(3, 9)
        TermChinAbun(10) = TotalChinEsc(6, 6) + TotalChinEsc(6, 7) + TotalChinEsc(6, 8) + TotalChinEsc(6, 9)
        TStep = 3 '- by definition
        For Fish = 49 To 50   '---- Area 8A Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk >= 7 And Stk <= 9 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        '--- Tulalip uses ETRS ... Don't add 8A Catch for Stock #10  3/24/99
                    End If
                    TermChinAbun(3) = TermChinAbun(3) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                    TermChinAbun(10) = TermChinAbun(10) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                Next Age
            Next Stk
        Next Fish
        '--- Tulalip Bay Net
        '-     When 8D Sport is Term add to TRS
        If TerminalFisheryFlag(48, 3) = 1 Then
            For Stk = 1 To AllStocks '- NumStk / 2
                For TStep = 3 To 4 '---- only Step 3 and 4 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk >= 7 And Stk <= 9 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep)
                            TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep)
                        End If
                        TermChinAbun(3) = TermChinAbun(3) + LandedCatch(Stk * 2 - 1, Age, 48, TStep)
                        TermChinAbun(10) = TermChinAbun(10) + LandedCatch(Stk * 2, Age, 48, TStep)
                        If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                            TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + (LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep))
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep))
                            TotalChinEsc(4, 9) = TotalChinEsc(4, 9) + (LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep))
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + (LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep))
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + (LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep))
                        Else
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk * 2 - 1, Age, 48, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, 48, TStep))
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + (LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep))
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + (LandedCatch(Stk * 2, Age, 48, TStep) + MSFLandedCatch(Stk * 2, Age, 48, TStep))
                        End If
                    Next Age
                Next TStep
            Next Stk
        End If
        For Fish = 51 To 52
            For Stk = 1 To AllStocks '- NumStk / 2
                For TStep = 3 To 4 '---- only Step 3 and 4 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk >= 7 And Stk <= 9 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        End If
                        TermChinAbun(3) = TermChinAbun(3) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TermChinAbun(10) = TermChinAbun(10) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                            TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(4, 9) = TotalChinEsc(4, 9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        Else
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- South Sound TAA

        TermChinAbun(4) = TotalChinEsc(3, 10) + TotalChinEsc(3, 11) + TotalChinEsc(3, 12) '+ TotalChinEsc(3, 13) ' + TotalChinEsc(3, 22) 'AHB 12/14/2015 White should not be included
        TermChinAbun(11) = TotalChinEsc(6, 10) + TotalChinEsc(6, 11) + TotalChinEsc(6, 12) '+ TotalChinEsc(6, 13) '+ TotalChinEsc(6, 22)
        Sps1011u = 0.0
        USPS0u = 0.0
        UWACCu = 0.0
        DSPS0u = 0.0
        SPSYRu = 0.0
        NONSSu = 0.0

        Sps1011m = 0.0
        USPS0m = 0.0
        UWACCm = 0.0
        DSPS0m = 0.0
        SPSYRm = 0.0
        NONSSm = 0.0

        TStep = 3 '- by definition
        Dim UnSPSMort, MkSPSMort As Double
        For Fish = 58 To 71
            If Fish > 63 And Fish < 68 Then GoTo NotFish
            For Stk = 1 To AllStocks '- NumStk / 2



                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS catches

                    If FisheryFlag(Fish, TStep) > 6 Then
                        '- SPS MSF - TAMM Fix 2010
                        UnSPSMort = (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep)) + MSFNonRetention(Stk * 2 - 1, Age, Fish, TStep)
                        MkSPSMort = (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep)) + MSFNonRetention(Stk * 2, Age, Fish, TStep)
                    Else
                        '- Net Fisheries
                        UnSPSMort = (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        MkSPSMort = (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                    End If

                    TermChinAbun(4) = TermChinAbun(4) + UnSPSMort
                    TermChinAbun(11) = TermChinAbun(11) + MkSPSMort
                    '---------- TAA
                    '-- 10/11 and 13 catches TRS for FRAM SS Stocks
                    '   both Falls and 13A Springs

                    If Fish = 68 And Stk = 11 Then
                        Stk = 13
                    End If
                    'If ((Stk > 10 And Stk < 16) Or Stk = 33) And (Fish <= 59 Or Fish = 68 Or Fish = 69) Then
                    If ((Stk > 10 And Stk < 16) Or Stk = 33) And Fish <= 59 Then 'AHB 12/14/2015 Fish 68 and 69 (13+ Net)are part of ETRS and get added later
                        Select Case Stk
                            Case 11  '---USPS FF---
                                TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + UnSPSMort
                                TotalChinEsc(6, 10) = TotalChinEsc(6, 10) + MkSPSMort
                            Case 12  '---UW Acc---
                                TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UnSPSMort
                                TotalChinEsc(6, 11) = TotalChinEsc(6, 11) + MkSPSMort
                            Case 13  '---DSPS FF---
                                TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + UnSPSMort
                                TotalChinEsc(6, 12) = TotalChinEsc(6, 12) + MkSPSMort
                            Case 14  '---SPS Yrl--- 
                                TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + UnSPSMort
                                TotalChinEsc(6, 13) = TotalChinEsc(6, 13) + MkSPSMort
                                'Case 15  '--- WhRvr Spring Fing  ' for White use ETRS only - White matures in time 2 and 3 TRS not needed in FRAM/TAMM
                                '    TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + UnSPSMort
                                '    TotalChinEsc(6, 14) = TotalChinEsc(6, 14) + MkSPSMort
                                'Case 33  '--- WhRvr Spring Year
                                '    TotalChinEsc(3, 22) = TotalChinEsc(3, 22) + UnSPSMort
                                '    TotalChinEsc(6, 22) = TotalChinEsc(6, 22) + MkSPSMort
                                'Case 38  '--- Hoko
                                '   TotalChinEsc(3, 23) = TotalChinEsc(3, 23) + UnSPSMort
                                '   TotalChinEsc(6, 23) = TotalChinEsc(6, 23) + MkSPSMort
                        End Select
                    End If
                    If Fish <= 59 Then       '--- 10/11 Net Catches for Split TAA
                        Sps1011u = Sps1011u + UnSPSMort
                        Sps1011m = Sps1011m + MkSPSMort
                    End If
                    '------ 10A,10E,13A,SPS Net ETRS Catches
                    If (Fish >= 60 And Fish <= 63) Or (Fish >= 68 And Fish <= 71) Then
                        Select Case Stk
                            Case 11
                                USPS0u = USPS0u + UnSPSMort
                                USPS0m = USPS0m + MkSPSMort
                            Case 12
                                UWACCu = UWACCu + UnSPSMort
                                UWACCm = UWACCm + MkSPSMort
                            Case 13
                                DSPS0u = DSPS0u + UnSPSMort
                                DSPS0m = DSPS0m + MkSPSMort
                            Case 14
                                SPSYRu = SPSYRu + UnSPSMort
                                SPSYRm = SPSYRm + MkSPSMort
                                'Case 33
                                '   SPSYRu = SPSYRu + UnSPSMort
                                '   SPSYRm = SPSYRm + MkSPSMort
                            Case Else
                                NONSSu = NONSSu + UnSPSMort
                                NONSSm = NONSSm + MkSPSMort
                        End Select
                    End If
                Next Age
            Next Stk
NotFish:
        Next Fish

        SSETACu = USPS0u + UWACCu + DSPS0u + SPSYRu
        If SSETACu <> 0.0 Then
            'TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0u + (NONSSu * (USPS0u / SSETACu))
            'TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACCu + (NONSSu * (UWACCu / SSETACu))
            'TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0u + (NONSSu * (DSPS0u / SSETACu))
            'TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYRu + (NONSSu * (SPSYRu / SSETACu))
            'TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0u + (NONSSu * (USPS0u / SSETACu))
            'TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACCu + (NONSSu * (UWACCu / SSETACu))
            'TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0u + (NONSSu * (DSPS0u / SSETACu))
            'TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYRu + (NONSSu * (SPSYRu / SSETACu))

            'TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0u + (NONSSu * (USPS0u / SSETACu)) + _
            '   UWACCu + (NONSSu * (UWACCu / SSETACu)) + _
            '   (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            'TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0u + (NONSSu * (DSPS0u / SSETACu)) + _
            '   (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            'TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0u + (NONSSu * (USPS0u / SSETACu)) + _
            '   UWACCu + (NONSSu * (UWACCu / SSETACu)) + _
            '   (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            'TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0u + (NONSSu * (DSPS0u / SSETACu)) + _
            '   (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            'TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            '- AHB Change 2/4/2011
            TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0u
            TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACCu
            TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0u
            TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYRu
            TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0u
            TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACCu
            TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0u
            TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYRu

            'TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0u + UWACCu + (SPSYRu * SpsYrSpl)
            'TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0u + (SPSYRu * (1.0 - SpsYrSpl))
            'TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + SPSYRu * SpsYrSpl
            'TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + SPSYRu * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0u + UWACCu + (SPSYRu * SpsYrSpl)
            'TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0u + SPSYRu * (1.0 - SpsYrSpl)
            'TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + SPSYRu * SpsYrSpl
            'TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + SPSYRu * (1.0 - SpsYrSpl)
        End If

        SSETACm = USPS0m + UWACCm + DSPS0m + SPSYRm
        If SSETACm <> 0.0 Then
            'TotalChinEsc(5, 10) = TotalChinEsc(5, 10) + USPS0m + (NONSSm * (USPS0m / SSETACm))
            'TotalChinEsc(5, 11) = TotalChinEsc(5, 11) + UWACCm + (NONSSm * (UWACCm / SSETACm))
            'TotalChinEsc(5, 12) = TotalChinEsc(5, 12) + DSPS0m + (NONSSm * (DSPS0m / SSETACm))
            'TotalChinEsc(5, 13) = TotalChinEsc(5, 13) + SPSYRm + (NONSSm * (SPSYRm / SSETACm))
            'TotalChinEsc(6, 10) = TotalChinEsc(6, 10) + USPS0m + (NONSSm * (USPS0m / SSETACm))
            'TotalChinEsc(6, 11) = TotalChinEsc(6, 11) + UWACCm + (NONSSm * (UWACCm / SSETACm))
            'TotalChinEsc(6, 12) = TotalChinEsc(6, 12) + DSPS0m + (NONSSm * (DSPS0m / SSETACm))
            'TotalChinEsc(6, 13) = TotalChinEsc(6, 13) + SPSYRm + (NONSSm * (SPSYRm / SSETACm))

            'TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + USPS0m + (NONSSm * (USPS0m / SSETACm)) + _
            '   UWACCm + (NONSSm * (UWACCm / SSETACm)) + _
            '   (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            'TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + DSPS0m + (NONSSm * (DSPS0m / SSETACm)) + _
            '   (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            'TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(6, 18) = TotalChinEsc(6, 18) + USPS0m + (NONSSm * (USPS0m / SSETACm)) + _
            '   UWACCm + (NONSSm * (UWACCm / SSETACm)) + _
            '   (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            'TotalChinEsc(6, 19) = TotalChinEsc(6, 19) + DSPS0m + (NONSSm * (DSPS0m / SSETACm)) + _
            '   (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            'TotalChinEsc(6, 20) = TotalChinEsc(6, 20) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            'TotalChinEsc(6, 21) = TotalChinEsc(6, 21) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            '- AHB Change 2/4/2011
            TotalChinEsc(5, 10) = TotalChinEsc(5, 10) + USPS0m
            TotalChinEsc(5, 11) = TotalChinEsc(5, 11) + UWACCm
            TotalChinEsc(5, 12) = TotalChinEsc(5, 12) + DSPS0m
            TotalChinEsc(5, 13) = TotalChinEsc(5, 13) + SPSYRm
            TotalChinEsc(6, 10) = TotalChinEsc(6, 10) + USPS0m
            TotalChinEsc(6, 11) = TotalChinEsc(6, 11) + UWACCm
            TotalChinEsc(6, 12) = TotalChinEsc(6, 12) + DSPS0m
            TotalChinEsc(6, 13) = TotalChinEsc(6, 13) + SPSYRm

            'TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + USPS0m + UWACCm + (SPSYRm * SpsYrSpl)
            'TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + DSPS0m + SPSYRm * (1.0 - SpsYrSpl)
            'TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + SPSYRm * SpsYrSpl
            'TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + SPSYRm * (1.0 - SpsYrSpl)
            'TotalChinEsc(6, 18) = TotalChinEsc(6, 18) + USPS0m + UWACCm + SPSYRm * SpsYrSpl
            'TotalChinEsc(6, 19) = TotalChinEsc(6, 19) + DSPS0m + SPSYRm * (1.0 - SpsYrSpl)
            'TotalChinEsc(6, 20) = TotalChinEsc(6, 20) + SPSYRm * SpsYrSpl
            'TotalChinEsc(6, 21) = TotalChinEsc(6, 21) + SPSYRm * (1.0 - SpsYrSpl)
        End If

        '------ Area 10/11 Net Catch Split between Upper and Deep South Sound TAA
        'USSETRSu = TotalChinEsc(3, 18)
        'DSSETRSu = TotalChinEsc(3, 19)
        'SUMETRSu = USSETRSu + DSSETRSu
        'If SUMETRSu <> 0.0 Then
        '    TermChinAbun(6) = USSETRSu + ((USSETRSu / SUMETRSu) * Sps1011u)
        '    TermChinAbun(7) = DSSETRSu + ((DSSETRSu / SUMETRSu) * Sps1011u)
        'End If
        'USSETRSm = TotalChinEsc(6, 18)
        'DSSETRSm = TotalChinEsc(6, 19)
        'SUMETRSm = USSETRSm + DSSETRSm
        'If SUMETRSm <> 0.0 Then
        '    TermChinAbun(13) = USSETRSm + ((USSETRSm / SUMETRSm) * Sps1011m)
        '    TermChinAbun(14) = DSSETRSm + ((DSSETRSm / SUMETRSm) * Sps1011m)
        'End If

        '- NOTE: WhRvrSpr not used in 13A as previously done after re-introduction

        '--- Hood Canal TAA
        TermChinAbun(5) = TotalChinEsc(3, 15) + TotalChinEsc(3, 16)
        TermChinAbun(12) = TotalChinEsc(6, 15) + TotalChinEsc(6, 16)
        TStep = 3 '- by definition
        For Fish = 65 To 66  '--- HC Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS
                    If Stk >= 16 And Stk <= 17 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                        TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                        '--- TRS
                    End If
                    TermChinAbun(5) = TermChinAbun(5) + (LandedCatch(Stk * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(Stk * 2 - 1, Age, Fish, TStep))
                    TermChinAbun(12) = TermChinAbun(12) + (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep))
                    '---------- TAA

                    If (LandedCatch(Stk * 2, Age, Fish, TStep) + MSFLandedCatch(Stk * 2, Age, Fish, TStep)) <> 0 Then
                        PrnLine = "StkFish " & (Stk * 2).ToString & "," & Fish.ToString & "," & LandedCatch(Stk * 2, Age, Fish, TStep).ToString(" #####0.0000") & "," & MSFLandedCatch(Stk * 2, Age, Fish, TStep).ToString(" #####0.0000")
                    End If
                    sw.WriteLine(PrnLine)

                Next Age
            Next Stk
        Next Fish

        '--- Subtract TAMM FW Sport and Add Marine Sport Savings to TAA's
        '--- Split proportion by Unmarked + Marked
        '- PS Chinook Run Reconstruction now includes FW Sport ... 1/3/2011 JFP
        'SptSave(1) = ((TNkFWSpt! - TNkMSA!) * (TermChinAbun(1) / (TermChinAbun(1) + TermChinAbun(8))))
        'SptSave(2) = ((TSkFWSpt! - TSkMSA!) * (TermChinAbun(2) / (TermChinAbun(2) + TermChinAbun(9))))
        'SptSave(3) = ((TSnFWSpt! - TSnMSA!) * (TermChinAbun(3) / (TermChinAbun(3) + TermChinAbun(10))))
        'SptSave(4) = (TermChinAbun(4) / (TermChinAbun(4) + TermChinAbun(11)))
        'SptSave(5) = (THCFWSpt! * (TermChinAbun(5) / (TermChinAbun(5) + TermChinAbun(12))))
        'SptSave(6) = TermChinAbun(6) / (TermChinAbun(6) + TermChinAbun(13))
        'SptSave(7) = TermChinAbun(7) / (TermChinAbun(7) + TermChinAbun(14))
        'SptSave(8) = ((TNkFWSpt! - TNkMSA!) * (TermChinAbun(8) / (TermChinAbun(1) + TermChinAbun(8))))
        'SptSave(9) = ((TSkFWSpt! - TSkMSA!) * (TermChinAbun(9) / (TermChinAbun(2) + TermChinAbun(9))))
        'SptSave(10) = ((TSnFWSpt! - TSnMSA!) * (TermChinAbun(10) / (TermChinAbun(3) + TermChinAbun(10))))
        'SptSave(11) = (TermChinAbun(11) / (TermChinAbun(4) + TermChinAbun(11)))
        'SptSave(12) = (THCFWSpt! * (TermChinAbun(12) / (TermChinAbun(5) + TermChinAbun(12))))
        'SptSave(13) = TermChinAbun(13) / (TermChinAbun(6) + TermChinAbun(13))
        'SptSave(14) = TermChinAbun(14) / (TermChinAbun(7) + TermChinAbun(14))
        'For I = 1 To 14
        '   TermChinAbun(I) = TermChinAbun(I) - SptSave(I)
        'Next I

        '------- Print Version Number and Command File Number ---

       

        xlWorkSheet = xlWorkBook.Sheets("TAMX")
        xlWorkSheet.Range("B1").Value = FramVersion
        xlWorkSheet.Range("B2").Value = RunIDNameSelect
        xlWorkSheet.Range("E1").Value = "RunDate:"
        xlWorkSheet.Range("F1").Value = RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("E2").Value = "RepDate:"
        xlWorkSheet.Range("F2").Value = DateTime.Now.ToString

        Dim TamxLineNum(19) As Integer
        TamxLineNum(0) = 0
        TamxLineNum(1) = 5
        TamxLineNum(2) = 7
        TamxLineNum(3) = 8
        TamxLineNum(4) = 9
        TamxLineNum(5) = 11
        TamxLineNum(6) = 12
        TamxLineNum(7) = 13
        TamxLineNum(8) = 14
        TamxLineNum(9) = 15
        TamxLineNum(10) = 17
        TamxLineNum(11) = 18
        TamxLineNum(12) = 19
        TamxLineNum(13) = 20
        TamxLineNum(14) = 26
        TamxLineNum(15) = 27
        TamxLineNum(16) = 28
        TamxLineNum(17) = 30
        TamxLineNum(18) = 31
        TamxLineNum(19) = 32
        xlApp.Application.Interactive = False
        '----- Put Terminal and Extreme Terminal Run Sizes into TAMX WorkSheet---
        Dim RngVal1 As String
        For Stk = 1 To 17
            RngVal1 = "B" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, Stk).ToString("######0")
            RngVal1 = "C" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, Stk).ToString("######0")
            RngVal1 = "D" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, Stk).ToString("######0")
            RngVal1 = "G" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(6, Stk).ToString("######0")
            RngVal1 = "H" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, Stk).ToString("######0")
            RngVal1 = "I" & (TamxLineNum(Stk)).ToString
            xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, Stk).ToString("######0")
            '- Put Stock Specific Terminal Run Sizes into WorkSheet -
            Select Case Stk
                Case 1
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(1).ToString("######0")
                    RngVal1 = "G" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(8).ToString("######0")
                Case 4
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(2).ToString("######0")
                    RngVal1 = "G" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(9).ToString("######0")
                Case 9
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(3).ToString("######0")
                    RngVal1 = "G" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(10).ToString("######0")
                Case 13
                    '--- Upper South Sound Yr.
                    'RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 20).ToString("######0")
                    'RngVal1 = "C" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 20).ToString("######0")
                    'RngVal1 = "D" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 20).ToString("######0")
                    'RngVal1 = "G" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(6, 20).ToString("######0")
                    'RngVal1 = "H" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 20).ToString("######0")
                    'RngVal1 = "I" & (TamxLineNum(Stk) + 1).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 20).ToString("######0")
                    ''--- Deep South Sound Yr.
                    'RngVal1 = "B" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 21).ToString("######0")
                    'RngVal1 = "C" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 21).ToString("######0")
                    'RngVal1 = "D" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 21).ToString("######0")
                    'RngVal1 = "G" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(6, 21).ToString("######0")
                    'RngVal1 = "H" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 21).ToString("######0")
                    'RngVal1 = "I" & (TamxLineNum(Stk) + 2).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 21).ToString("######0")
                    ''--- Upper South Sound Agg.
                    'RngVal1 = "B" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TermChinAbun(6).ToString("######0")
                    'RngVal1 = "C" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 18).ToString("######0")
                    'RngVal1 = "D" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 18).ToString("######0")
                    'RngVal1 = "G" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TermChinAbun(13).ToString("######0")
                    'RngVal1 = "H" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 18).ToString("######0")
                    'RngVal1 = "I" & (TamxLineNum(Stk) + 3).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 18)
                    ''--- Deep South Sound Agg.
                    'RngVal1 = "B" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TermChinAbun(7).ToString("######0")
                    'RngVal1 = "C" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 19).ToString("######0")
                    'RngVal1 = "D" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 19).ToString("######0")
                    'RngVal1 = "G" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TermChinAbun(14).ToString("######0")
                    'RngVal1 = "H" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 19).ToString("######0")
                    'RngVal1 = "I" & (TamxLineNum(Stk) + 4).ToString
                    'xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 19)
                    '--- Total TAA
                    RngVal1 = "B" & (TamxLineNum(Stk) + 5).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(4).ToString("######0")
                    RngVal1 = "G" & (TamxLineNum(Stk) + 5).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(11).ToString("######0")
                Case 16
                    RngVal1 = "B" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(5).ToString("######0")
                    RngVal1 = "G" & (TamxLineNum(Stk) + 1).ToString
                    xlWorkSheet.Range(RngVal1).Value = TermChinAbun(12).ToString("######0")
                Case Else
            End Select
        Next
        '- Add White River Springs to List (Sum Both Components)
        Stk = 18
        RngVal1 = "B" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 22).ToString("######0")
        RngVal1 = "C" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 22).ToString("######0")
        RngVal1 = "D" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 22).ToString("######0")
        RngVal1 = "G" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(6, 22).ToString("######0")
        RngVal1 = "H" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 22).ToString("######0")
        RngVal1 = "I" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 22).ToString("######0")
        '- Add Hoko to Bottom of List (Sum Both Components)
        Stk = 19
        RngVal1 = "B" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(3, 23).ToString("######0")
        RngVal1 = "C" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(1, 23).ToString("######0")
        RngVal1 = "D" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(2, 23).ToString("######0")
        RngVal1 = "G" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(6, 23).ToString("######0")
        RngVal1 = "H" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(4, 23).ToString("######0")
        RngVal1 = "I" & (TamxLineNum(Stk)).ToString
        xlWorkSheet.Range(RngVal1).Value = TotalChinEsc(5, 23).ToString("######0")

        '--- GET Fishery Landed Catch and Total Mortality DATA ---

        '- Dimension Transfer Arrays to Accomodate WorkSheet Transfer
        Dim FishVal As Integer
        Dim UnMkCatch(1, 1)  '- UnMarked Landed Catch
        Dim UnMkTMort(1, 1)  '- UnMarked Total Mortality
        Dim MarkCatch(1, 1)  '- Marked Landed Catch
        Dim MarkTMort(1, 1)  '- Marked Total Mortality
        Dim FishTMort(NumFish, NumSteps + 1)  '- Fishery Total Mortality
        If TammChinookRunFlag = 1 Then
            ReDim UnMkCatch(NumFish - 3, NumSteps)
            ReDim UnMkTMort(NumFish - 3, NumSteps)
            ReDim MarkCatch(NumFish - 3, NumSteps)
            ReDim MarkTMort(NumFish - 3, NumSteps)
        Else
            ReDim UnMkCatch(NumFish - 2, NumSteps)
            ReDim UnMkTMort(NumFish - 2, NumSteps)
            ReDim MarkCatch(NumFish - 2, NumSteps)
            ReDim MarkTMort(NumFish - 2, NumSteps)
        End If
        For Fish = 1 To NumFish
            '- Determine Fishery Numbers consistent with TAMM SpreadSheet
            If TammChinookRunFlag = 1 Then
                Select Case Fish
                    Case 1 To 12
                        FishVal = Fish
                    Case 13 To 15
                        FishVal = 13
                    Case 16 To 55
                        FishVal = Fish - 2
                    Case 56 To 57
                        FishVal = 54
                    Case 58 To NumFish
                        FishVal = Fish - 3
                End Select
            Else
                Select Case Fish
                    Case 1 To 12
                        FishVal = Fish
                    Case 13 To 15
                        FishVal = 13
                    Case 16 To NumFish
                        FishVal = Fish - 2
                End Select
            End If
            For TStep = 2 To NumSteps
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        If (Stk Mod 2) <> 0 Then '- UnMarked Catch & Mortality
                            UnMkCatch(FishVal - 1, TStep - 1) += ((LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            UnMkTMort(FishVal - 1, TStep - 1) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            UnMkCatch(FishVal - 1, NumSteps) += ((LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            UnMkTMort(FishVal - 1, NumSteps) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                        Else                     '- Marked Catch & Mortality
                            MarkCatch(FishVal - 1, TStep - 1) += ((LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            MarkTMort(FishVal - 1, TStep - 1) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            MarkCatch(FishVal - 1, NumSteps) += ((LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            MarkTMort(FishVal - 1, NumSteps) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                        End If
                        FishTMort(Fish, TStep) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                        FishTMort(Fish, NumSteps + 1) += ((LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                    Next
                Next
            Next
        Next
        '- Convert Arrays into Long Integer Values for WorkSheet
        'If TammChinookRunFlag = 1 Then
        '   For Fish = 0 To NumFish - 3
        '      For TStep = 0 To NumSteps
        '         UnMkCatch(Fish, TStep) = CLng(UnMkCatch(Fish, TStep))
        '         MarkCatch(Fish, TStep) = CLng(MarkCatch(Fish, TStep))
        '         UnMkTMort(Fish, TStep) = CLng(UnMkTMort(Fish, TStep))
        '         MarkTMort(Fish, TStep) = CLng(MarkTMort(Fish, TStep))
        '      Next
        '   Next
        'Else
        '   For Fish = 0 To NumFish - 2
        '      For TStep = 0 To NumSteps
        '         UnMkCatch(Fish, TStep) = CLng(UnMkCatch(Fish, TStep))
        '         MarkCatch(Fish, TStep) = CLng(MarkCatch(Fish, TStep))
        '         UnMkTMort(Fish, TStep) = CLng(UnMkTMort(Fish, TStep))
        '         MarkTMort(Fish, TStep) = CLng(MarkTMort(Fish, TStep))
        '      Next
        '   Next
        'End If
        '- Put TAMX arrays into WorkSheet
        If TammChinookRunFlag = 1 Then
            xlWorkSheet.Range("B42").Resize(NumFish - 3, NumSteps + 1).Value = UnMkTMort
            xlWorkSheet.Range("H42").Resize(NumFish - 3, NumSteps + 1).Value = MarkTMort
            xlWorkSheet.Range("P42").Resize(NumFish - 3, NumSteps + 1).Value = UnMkCatch
            xlWorkSheet.Range("V42").Resize(NumFish - 3, NumSteps + 1).Value = MarkCatch
        Else
            xlWorkSheet.Range("B42").Resize(NumFish - 2, NumSteps + 1).Value = UnMkTMort
            xlWorkSheet.Range("H42").Resize(NumFish - 2, NumSteps + 1).Value = MarkTMort
            xlWorkSheet.Range("P42").Resize(NumFish - 2, NumSteps + 1).Value = UnMkCatch
            xlWorkSheet.Range("V42").Resize(NumFish - 2, NumSteps + 1).Value = MarkCatch
        End If

        'xlWorkSheet.Range("BV41").Resize(NumFish + 1, NumSteps + 2).Value = FishTMort

        '-------------------------- STOCK CATCH BY FISHERY ---
        Dim FishNum As Integer
        For Stk = 1 To PSStocks
            If TammChinookRunFlag = 1 Then
                ReDim UnMkCatch(NumFish - 3, NumSteps)
                ReDim UnMkTMort(NumFish - 3, NumSteps)
                ReDim MarkCatch(NumFish - 3, NumSteps)
                ReDim MarkTMort(NumFish - 3, NumSteps)
            Else
                ReDim UnMkCatch(NumFish - 2, NumSteps)
                ReDim UnMkTMort(NumFish - 2, NumSteps)
                ReDim MarkCatch(NumFish - 2, NumSteps)
                ReDim MarkTMort(NumFish - 2, NumSteps)
            End If
            '- Determine Stock Numbers and Sequence Values
            '- WhRvr Spring Yearling and Hoko Added (#33 & # 38)
            If Stk = 1 Then
                StkNum = 1
                StkVal = 1
            ElseIf Stk = 2 Or Stk = 3 Then
                StkNum = Stk
                StkVal = 2
            ElseIf Stk > 3 And Stk < 19 Then
                StkNum = Stk
                StkVal = Stk - 1
            ElseIf Stk = 19 Then
                StkNum = 33 '- WhRvr Spr Year
                StkVal = 18
            ElseIf Stk = 20 Then
                StkNum = 38 '- Hoko
                StkVal = 19
            End If
NooksackSpringReEntry2:
            For Fish = 1 To NumFish
                '- Combined Fisheries
                If TammChinookRunFlag = 1 Then
                    '- Old Style Format (Area 10/11 Sport Combined)
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 55
                            FishNum = Fish - 2
                        Case 56, 57
                            FishNum = 54
                        Case 58 To 73
                            FishNum = Fish - 3
                    End Select
                Else
                    '- Current Chinook TAMM Transfer
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 73
                            FishNum = Fish - 2
                    End Select
                End If
                For TStep = 2 To NumSteps
                    For Age = MinAge To MaxAge
                        '- AEQ Value NOT used for Total Mortality in Terminal Fisheries
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                            UnMkTMort(FishNum - 1, TStep - 1) += (LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + NonRetention(StkNum * 2 - 1, Age, Fish, TStep) + Shakers(StkNum * 2 - 1, Age, Fish, TStep) + DropOff(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFNonRetention(StkNum * 2 - 1, Age, Fish, TStep) + MSFShakers(StkNum * 2 - 1, Age, Fish, TStep) + MSFDropOff(StkNum * 2 - 1, Age, Fish, TStep))
                            MarkTMort(FishNum - 1, TStep - 1) += (LandedCatch(StkNum * 2, Age, Fish, TStep) + NonRetention(StkNum * 2, Age, Fish, TStep) + Shakers(StkNum * 2, Age, Fish, TStep) + DropOff(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep) + MSFNonRetention(StkNum * 2, Age, Fish, TStep) + MSFShakers(StkNum * 2, Age, Fish, TStep) + MSFDropOff(StkNum * 2, Age, Fish, TStep))
                            UnMkTMort(FishNum - 1, NumSteps) += (LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + NonRetention(StkNum * 2 - 1, Age, Fish, TStep) + Shakers(StkNum * 2 - 1, Age, Fish, TStep) + DropOff(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFNonRetention(StkNum * 2 - 1, Age, Fish, TStep) + MSFShakers(StkNum * 2 - 1, Age, Fish, TStep) + MSFDropOff(StkNum * 2 - 1, Age, Fish, TStep))
                            MarkTMort(FishNum - 1, NumSteps) += (LandedCatch(StkNum * 2, Age, Fish, TStep) + NonRetention(StkNum * 2, Age, Fish, TStep) + Shakers(StkNum * 2, Age, Fish, TStep) + DropOff(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep) + MSFNonRetention(StkNum * 2, Age, Fish, TStep) + MSFShakers(StkNum * 2, Age, Fish, TStep) + MSFDropOff(StkNum * 2, Age, Fish, TStep))
                        Else
                            UnMkTMort(FishNum - 1, TStep - 1) += ((LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + NonRetention(StkNum * 2 - 1, Age, Fish, TStep) + Shakers(StkNum * 2 - 1, Age, Fish, TStep) + DropOff(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFNonRetention(StkNum * 2 - 1, Age, Fish, TStep) + MSFShakers(StkNum * 2 - 1, Age, Fish, TStep) + MSFDropOff(StkNum * 2 - 1, Age, Fish, TStep)) * AEQ(StkNum * 2 - 1, Age, TStep))
                            MarkTMort(FishNum - 1, TStep - 1) += ((LandedCatch(StkNum * 2, Age, Fish, TStep) + NonRetention(StkNum * 2, Age, Fish, TStep) + Shakers(StkNum * 2, Age, Fish, TStep) + DropOff(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep) + MSFNonRetention(StkNum * 2, Age, Fish, TStep) + MSFShakers(StkNum * 2, Age, Fish, TStep) + MSFDropOff(StkNum * 2, Age, Fish, TStep)) * AEQ(StkNum * 2, Age, TStep))
                            UnMkTMort(FishNum - 1, NumSteps) += ((LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + NonRetention(StkNum * 2 - 1, Age, Fish, TStep) + Shakers(StkNum * 2 - 1, Age, Fish, TStep) + DropOff(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFNonRetention(StkNum * 2 - 1, Age, Fish, TStep) + MSFShakers(StkNum * 2 - 1, Age, Fish, TStep) + MSFDropOff(StkNum * 2 - 1, Age, Fish, TStep)) * AEQ(StkNum * 2 - 1, Age, TStep))
                            MarkTMort(FishNum - 1, NumSteps) += ((LandedCatch(StkNum * 2, Age, Fish, TStep) + NonRetention(StkNum * 2, Age, Fish, TStep) + Shakers(StkNum * 2, Age, Fish, TStep) + DropOff(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep) + MSFNonRetention(StkNum * 2, Age, Fish, TStep) + MSFShakers(StkNum * 2, Age, Fish, TStep) + MSFDropOff(StkNum * 2, Age, Fish, TStep)) * AEQ(StkNum * 2, Age, TStep))
                        End If
                        '- Landed Catch Arrays
                        UnMkCatch(FishNum - 1, TStep - 1) += (LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep))
                        MarkCatch(FishNum - 1, TStep - 1) += (LandedCatch(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep))
                        UnMkCatch(FishNum - 1, NumSteps) += (LandedCatch(StkNum * 2 - 1, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2 - 1, Age, Fish, TStep))
                        MarkCatch(FishNum - 1, NumSteps) += (LandedCatch(StkNum * 2, Age, Fish, TStep) + MSFLandedCatch(StkNum * 2, Age, Fish, TStep))
                    Next Age
                Next TStep
            Next Fish
            '- Nooksack Spring Chinook now has 2 stocks - Add together
            If Stk = 2 And StkNum = 2 Then

                If TammChinookRunFlag = 1 Then
                    RngVal1 = "AA" & (StkVal * 72 + 42).ToString
                    xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = UnMkTMort
                Else
                    RngVal1 = "AA" & (StkVal * 73 + 42).ToString
                    xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = UnMkTMort
                End If

                StkNum = 3
                Stk = 3
                StkVal = 2
                GoTo NooksackSpringReEntry2
            End If
            '- Convert Arrays into Long Integer Values for WorkSheet
            ' -.... NOTE: ...... AHB wants "real" numbers in worksheet, not integers!!!!! 2/9/2012 JFP
            'If TammChinookRunFlag = 1 Then
            '   For Fish = 0 To NumFish - 3
            '      For TStep = 0 To NumSteps
            '         UnMkCatch(Fish, TStep) = CLng(UnMkCatch(Fish, TStep))
            '         MarkCatch(Fish, TStep) = CLng(MarkCatch(Fish, TStep))
            '         UnMkTMort(Fish, TStep) = CLng(UnMkTMort(Fish, TStep))
            '         MarkTMort(Fish, TStep) = CLng(MarkTMort(Fish, TStep))
            '      Next
            '   Next
            'Else
            '   For Fish = 0 To NumFish - 2
            '      For TStep = 0 To NumSteps
            '         UnMkCatch(Fish, TStep) = CLng(UnMkCatch(Fish, TStep))
            '         MarkCatch(Fish, TStep) = CLng(MarkCatch(Fish, TStep))
            '         UnMkTMort(Fish, TStep) = CLng(UnMkTMort(Fish, TStep))
            '         MarkTMort(Fish, TStep) = CLng(MarkTMort(Fish, TStep))
            '      Next
            '   Next
            'End If
            '- Put Stock Catch and Stock TotalMort Arrays into WorkSheet
            If TammChinookRunFlag = 1 Then
                RngVal1 = "B" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = UnMkTMort
                RngVal1 = "H" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = MarkTMort
                RngVal1 = "P" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = UnMkCatch
                RngVal1 = "V" & (StkVal * 72 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 3, NumSteps + 1).Value = MarkCatch
            Else
                RngVal1 = "B" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = UnMkTMort
                RngVal1 = "H" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = MarkTMort
                RngVal1 = "P" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = UnMkCatch
                RngVal1 = "V" & (StkVal * 73 + 42).ToString
                xlWorkSheet.Range(RngVal1).Resize(NumFish - 2, NumSteps + 1).Value = MarkCatch
            End If
        Next
        xlWorkSheet.Range("N7").Value = NookSprETRS_NOR
       
        '- Save WorkBook and Close Application if Necessary
        'xlApp.Application.DisplayAlerts = False
        'xlWorkBook.Save()
        'If WorkBookWasNotOpen = True Then
        '   xlWorkBook.Close()
        'End If
        'If ExcelWasNotRunning = True Then
        '   xlApp.Application.Quit()
        '   xlApp.Quit()
        'Else
        '   xlApp.Visible = True
        '   xlApp.WindowState = Excel.XlWindowState.xlMinimized
        'End If
        xlApp.Visible = True
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Interactive = True
        'xlApp = Nothing

    End Sub

    Sub TammCohoProc()

        Dim TRS(55) As Double, TRSType(55) As Integer
        Dim TargetLocal, TRSLocalCatch As Double
        Dim TRSName(55) As String
        Dim NumTrsStk, NumTrsFish As Integer
        Dim TrsStk(55, 50), TrsFish(55, 30) As Integer
        Dim NumTRS, TS1, TS2, I, J, K As Integer
        'Dim SaveTermFlag(86, 1) As Integer
        'Dim SaveTermQuota(86, 1) As Double

        '- Get TaaEtrs Instructions from Database Table
        Dim CmdStr As String
        CmdStr = "SELECT * FROM TaaETRSList ORDER BY TaaNum"
        Dim TAAcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim TAADA As New System.Data.OleDb.OleDbDataAdapter
        TAADA.SelectCommand = TAAcm
        Dim TAAcb As New OleDb.OleDbCommandBuilder
        TAAcb = New OleDb.OleDbCommandBuilder(TAADA)
        If FramDataSet.Tables.Contains("TaaETRSList") Then
            FramDataSet.Tables("TaaETRSList").Clear()
        End If
        TAADA.Fill(FramDataSet, "TaaETRSList")

        Dim RecNum, ParseOld, ParseNew As Integer
        Dim OptionList As String
        NumTRS = FramDataSet.Tables("TaaETRSList").Rows.Count

        '- Check for TaaETRS Variables in DataBase Table
        If NumTRS = 0 Then
            MsgBox("Coho TAMM Procedure Requires 'TaaETRSList' in DataBase Table" & vbCrLf & "Please Read Text File (TaaETRSNum.Txt) and Re-Run", MsgBoxStyle.OkOnly)
            '- Done with TAMM WorkBook for this run .. Close and release object
            'xlApp.Application.DisplayAlerts = False
            'xlWorkBook.Save()
            'If WorkBookWasNotOpen = True Then
            '   xlWorkBook.Close()
            'End If
            'If ExcelWasNotRunning = True Then
            '   xlApp.Application.Quit()
            '   xlApp.Quit()
            'Else
            '   xlApp.Visible = True
            '   xlApp.WindowState = Excel.XlWindowState.xlMinimized
            'End If
            xlApp.Visible = True
            xlApp.Application.DisplayAlerts = True
            xlApp.Application.Interactive = True
            'xlApp = Nothing
            Exit Sub
        End If

        For RecNum = 0 To NumTRS - 1
            NumTrsStk = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(1)
            If NumTrsStk = 0 Then GoTo NextTaaETRS
            TrsStk(RecNum + 1, 0) = NumTrsStk
            OptionList = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(2)
            '- Parse TAA Stock Numbers from DataBase List
            ParseOld = 1
            For Stk = 1 To NumStk
                ParseNew = InStr(ParseOld, OptionList, ",")
                If ParseNew = 0 Then
                    '- Last Stock in List
                    TrsStk(RecNum + 1, Stk) = CInt(OptionList.Substring(ParseOld - 1, OptionList.Length - ParseOld + 1))
                    Exit For
                Else
                    TrsStk(RecNum + 1, Stk) = CInt(OptionList.Substring(ParseOld - 1, ParseNew - ParseOld))
                End If
                ParseOld = ParseNew + 1
            Next
            NumTrsFish = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(3)
            TrsFish(RecNum + 1, 0) = NumTrsFish
            OptionList = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(4)
            '- Parse TAA Fishery Numbers from DataBase List
            ParseOld = 1
            For Fish = 1 To NumFish
                ParseNew = InStr(ParseOld, OptionList, ",")
                If ParseNew = 0 Then
                    '- Last Stock in List
                    TrsFish(RecNum + 1, Fish) = CInt(OptionList.Substring(ParseOld - 1, OptionList.Length - ParseOld + 1))
                    Exit For
                Else
                    TrsFish(RecNum + 1, Fish) = CInt(OptionList.Substring(ParseOld - 1, ParseNew - ParseOld))
                End If
                ParseOld = ParseNew + 1
            Next
            TS1 = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(5)
            TS2 = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(6)
            TRSType(RecNum + 1) = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(7)
            TRSName(RecNum + 1) = FramDataSet.Tables("TaaETRSList").Rows(RecNum)(8)
NextTaaETRS:
        Next

        '- Run Coastal Iterations and Coho TAMM Iterations
        Dim myPoint As Point = FVS_RunModel.RunProgressLabel.Location
        '**************************************************************************************************************************************
        Itercount = 1
        If CoastalIterations = True Then
            KeepIter = False
            
            For K = 4 To 5
                ' overwrite coastal fisheries with TAMM quota

                If K = 4 Then
                    For Fish = 45 To 74
                        If Fish = 68 Then
                            Jim = 1
                        End If
                        Select Case Fish
                            Case 45
                                If BaseExploitationRate(161, 3, 45, 4) <> 0 Then
                                    If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                        KeepIter = True
                                    End If
                                    FisheryFlag(Fish, K) = 2
                                    FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                    FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                End If
                            Case 47, 48, 50, 52, 63, 68, 71, 74
                                If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                    KeepIter = True
                                End If
                                FisheryFlag(Fish, K) = 2
                                FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                        End Select
                    Next Fish
                Else
                    For Fish = 45 To 74
                        Select Case Fish
                            Case 45
                                If BaseExploitationRate(161, 3, 45, 4) <> 0 Then
                                    If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                        KeepIter = True
                                    End If
                                    FisheryFlag(Fish, K) = 2
                                    FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                    FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                Else
                                    If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                        KeepIter = True
                                    End If
                                    FisheryFlag(Fish, K) = 2
                                    FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K) + SaveCoastalQuota(Fish, 4)
                                    FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K) + SaveCoastalQuota(Fish, 4)
                                    FisheryQuota(Fish, K) = FisheryQuota(Fish, K).ToString("########0")
                                End If

                            Case 47, 48, 50, 52, 63, 68, 71, 74
                                If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                    KeepIter = True
                                End If
                                FisheryFlag(Fish, K) = 2
                                FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                                FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K).ToString("########0")
                            Case 46, 49, 51, 54, 55, 56, 65, 70, 73
                                If Abs(CInt(FisheryQuotaCompare(Fish, K)) - CInt(FisheryQuota(Fish, K))) > 1 Then
                                    KeepIter = True
                                End If
                                FisheryFlag(Fish, K) = 2
                                FisheryQuota(Fish, K) = SaveCoastalQuota(Fish, K) + SaveCoastalQuota(Fish, 4)
                                FisheryQuotaCompare(Fish, K) = SaveCoastalQuota(Fish, K) + SaveCoastalQuota(Fish, 4)
                                FisheryQuota(Fish, K) = FisheryQuota(Fish, K).ToString("########0")
                        End Select
                    Next Fish
                End If
            Next
            If KeepIter = True Then
                Itercount = Itercount + 1
            End If
        End If

        For TammIteration = 1 To 5
            '- Label Update
            FVS_RunModel.RunProgressLabel.Text = " TAMM Iteration = " & TammIteration.ToString & " "
            myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
            FVS_RunModel.RunProgressLabel.Location = myPoint
            FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
            FVS_RunModel.RunProgressLabel.Refresh()
            PrnLine = " TAMM Iteration " & TammIteration.ToString
            sw.WriteLine(PrnLine)
            '- Sum Terminal Runsizes

            If TammIteration = 2 Then
                Jim = 1
            End If

            For I = 1 To NumTRS
                TRS(I) = 0
                For J = 1 To TrsStk(I, 0)
                    TRS(I) = TRS(I) + Escape(TrsStk(I, J), 3, 5)
                Next J
                PrnLine = "TRS#=" & I.ToString & " Escape=" & Format("{0,7:G}", TRS(I).ToString("#####0"))
                For J = 1 To TrsFish(I, 0)
                    If TRSType(I) = 1 Then              '- TAA Calculations
                        For K = 1 To NumStk
                            TRS(I) = TRS(I) + LandedCatch(K, 3, TrsFish(I, J), 4) + LandedCatch(K, 3, TrsFish(I, J), 5) + MSFLandedCatch(K, 3, TrsFish(I, J), 4) + MSFLandedCatch(K, 3, TrsFish(I, J), 5)
                        Next K
                    Else                        '- ETRS Calculations
                        For K = 1 To TrsStk(I, 0)
                            TRS(I) = TRS(I) + LandedCatch(TrsStk(I, K), 3, TrsFish(I, J), 4) + LandedCatch(TrsStk(I, K), 3, TrsFish(I, J), 5) + MSFLandedCatch(TrsStk(I, K), 3, TrsFish(I, J), 4) + MSFLandedCatch(TrsStk(I, K), 3, TrsFish(I, J), 5)
                        Next K
                    End If
                Next J
                PrnLine &= " Total=" & Format("{0,7:G}", TRS(I).ToString("#####0"))
                sw.WriteLine(PrnLine)
            Next I


            '- Scale Terminal Fisheries using Terminal Runsizes
            For K = 4 To 5
                For I = 80 To 166
                    If I = 112 And K = 4 Then
                        Jim = 1
                    End If
                    If CohoTammFlag(K, I) = 3 Then '- Percent TAA
                        FisheryFlag(I, K) = 2
                        FisheryQuota(I, K) = TRS(CohoTammFish(K, I)) * CohoTammRate(K, I)

                        PrnLine = "TAAQuota-" & FisheryName(I) & I.ToString(" 000") & K.ToString(" 0") & String.Format("{0,10}", FisheryQuota(I, K).ToString("########0"))
                        sw.WriteLine(PrnLine)

                    ElseIf CohoTammFlag(K, I) = 4 Then '- Percent ETRS Local Stock
                        TRSLocalCatch = 0
                        For J = 1 To TrsStk(CohoTammFish(K, I), 0)
                            TRSLocalCatch = TRSLocalCatch + LandedCatch(TrsStk(CohoTammFish(K, I), J), 3, I, K)
                        Next J
                        TargetLocal = TRS(CohoTammFish(K, I)) * CohoTammRate(K, I)
                        FisheryFlag(I, K) = 2
                        If TRSLocalCatch = 0 Then
                            FisheryQuota(I, K) = TargetLocal
                        Else
                            FisheryQuota(I, K) = TargetLocal * (TotalLandedCatch(I, K) / TRSLocalCatch)
                        End If

                        PrnLine = "TAA-ETRS-" & FisheryName(I) & I.ToString(" 000") & K.ToString(" 0") & String.Format("{0,10}", FisheryQuota(I, K).ToString("########0")) & String.Format("{0,8}", TargetLocal.ToString(" #####0"))
                        sw.WriteLine(PrnLine)
                    End If
                Next I


            Next K


            '- ReSet Time 4 Pre-Terminal Cohort to Original Value for Next Iteration
            For Stk = 1 To NumStk
                Cohort(Stk, 3, PTerm, 4) = CohoTime4Cohort(Stk)
            Next Stk

            '- ReRun Catch and Escapement Calculations for Terminal Time Steps
            For TStep = 4 To NumSteps
                Call NatMort()
                Call CompCatch(PTerm)
                Call IncMort(PTerm)
                Call Mature()
                If TStep = 5 Then
                    Jim = 1
                End If
                Call CompCatch(Term)
                Call IncMort(Term)
                If TStep = 5 Then
                    Jim = 1
                End If
                Call CompEscape()
                '- Put Cohort Numbers into Next Time Step
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        If TStep < NumSteps Then
                            Cohort(Stk, Age, 0, TStep + 1) = Cohort(Stk, Age, 0, TStep)
                        End If
                    Next
                Next
            Next TStep

        Next TammIteration

        '- Old VB Method  ... No Longer Support Fishery Scaler Only Option from Run Menu
        'For Fish = 1 To NumFish
        '   For TStep = 1 To NumSteps
        '      If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 7 Then
        '         FisheryQuota(Fish, TStep) = CLng(TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish))
        '      ElseIf FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 8 Then
        '         If OptionReplaceQuota = True Then
        '            If FisheryFlag(Fish, TStep) = 2 Then
        '               FisheryFlag(Fish, TStep) = 1
        '            Else
        '               FisheryFlag(Fish, TStep) = 7
        '            End If
        '         End If
        '      End If
        '   Next
        'Next

        '- Restore Original FisheryQuota and FisheryFlag Values before Saving
        '- AHB 11/2/17 comment out this section to update Fishery Quota in order to get a match between scalar and quota
        'For Fish = 80 To 166
        '    For TStep = 4 To 5
        '        FisheryFlag(Fish, TStep) = SaveTermFlag(Fish - 80, TStep - 4)
        '        FisheryQuota(Fish, TStep) = SaveTermQuota(Fish - 80, TStep - 4)
        '    Next
        'Next

        'AHB 11/6/2017 compute quotas for scalar fisheries to create a match between scalars and quotas
        Dim TotalSum As Double
        For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps



                '- Retention Fishery Scaler
                If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Then
                    TotalSum = 0
                    For Stk As Integer = 1 To NumStk
                        For Age As Integer = MinAge To MaxAge
                            TotalSum += LandedCatch(Stk, Age, Fish, TStep)
                        Next
                    Next
                    If TotalSum > 0 Then
                        FisheryQuota(Fish, TStep) = CDbl(TotalSum / ModelStockProportion(Fish))
                    Else
                        FisheryQuota(Fish, TStep) = 0
                    End If
                End If
               

                If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Then
                    TotalSum = 0
                    For Stk As Integer = 1 To NumStk
                        For Age As Integer = MinAge To MaxAge
                            TotalSum += MSFLandedCatch(Stk, Age, Fish, TStep)
                        Next
                    Next
                    If TotalSum > 0 Then
                        MSFFisheryQuota(Fish, TStep) = CDbl(TotalSum / ModelStockProportion(Fish))
                    Else
                        MSFFisheryQuota(Fish, TStep) = 0
                    End If
                End If
                
            Next
        Next

        Call SaveDat()

        '- Label Update
        FVS_RunModel.RunProgressLabel.Text = " Transferring Data to TAMM Tables "
        myPoint.X = (FVS_RunModel.Width - FVS_RunModel.RunProgressLabel.Width) \ 2
        FVS_RunModel.RunProgressLabel.Location = myPoint
        FVS_RunModel.RunProgressLabel.TextAlign = ContentAlignment.MiddleCenter
        FVS_RunModel.RunProgressLabel.Refresh()

        If TammTransferSave = True Then Call CohoTAMMReports()

    End Sub

    Sub CohoTAMMReports()

        Dim ci As CultureInfo = CultureInfo.InvariantCulture

        '**************************************************************
        'automatically loads coho PRN-reports into coho TAMM

        '- Make Array like Old FishSumAll.PRN Report for Excel Transfer
        '- WorkBook still OPEN from ReadCohoTamm
        '- Need to Change WorkSheet for Data Transfer

        xlWorkSheet = xlWorkBook.Sheets("FishSumAllPRN")

        xlWorkSheet.Range("A2:J7").Clear()
        xlWorkSheet.Range("A2").Value = "Species: COHO"
        xlWorkSheet.Range("A3").Value = "Report: Fishery Summary"
        xlWorkSheet.Range("A7").Value = "Landed Catch by Fishery"
        xlWorkSheet.Range("C2").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A4").Value = "Title:" & RunIDTitleSelect
        xlWorkSheet.Range("A5").Value = RunIDNameSelect
        xlWorkSheet.Range("G2").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G3").Value = "RepDate:" & Now().ToString

        xlWorkSheet.Range("A211:J216").Clear()
        xlWorkSheet.Range("A211").Value = "Species: COHO"
        xlWorkSheet.Range("A212").Value = "Report: Fishery Summary"
        xlWorkSheet.Range("A216").Value = "Shaker Mortality by Fishery"
        xlWorkSheet.Range("C211").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A213").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G211").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G212").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G211").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G212").Value = "RepDate:" & Now().ToString

        xlWorkSheet.Range("A420:J425").Clear()
        xlWorkSheet.Range("A420").Value = "Species: COHO"
        xlWorkSheet.Range("A421").Value = "Report: Fishery Summary"
        xlWorkSheet.Range("A425").Value = "Non-Retention Mortality by Fishery"
        xlWorkSheet.Range("C420").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A422").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G420").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G421").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G420").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G421").Value = "RepDate:" & Now().ToString

        xlWorkSheet.Range("A629:J634").Clear()
        xlWorkSheet.Range("A629").Value = "Species: COHO"
        xlWorkSheet.Range("A630").Value = "Report: Fishery Summary"
        xlWorkSheet.Range("A634").Value = "Catch+CNR Mortality by Fishery"
        xlWorkSheet.Range("C629").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A631").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G629").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G630").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G629").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G630").Value = "RepDate:" & Now().ToString

        xlWorkSheet.Range("A838:J843").Clear()
        xlWorkSheet.Range("A838").Value = "Species: COHO"
        xlWorkSheet.Range("A839").Value = "Report: Fishery Summary"
        xlWorkSheet.Range("A843").Value = "Total Mortality by Fishery"
        xlWorkSheet.Range("C838").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A840").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G838").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G839").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G838").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G839").Value = "RepDate:" & Now().ToString

        '- Excel Transfer Variables 
        Dim FishSumLC(NumFish - 1, NumSteps) As Double
        Dim FishSumSH(NumFish - 1, NumSteps) As Double
        Dim FishSumNR(NumFish - 1, NumSteps) As Double
        Dim FishSumCC(NumFish - 1, NumSteps) As Double
        Dim FishSumTM(NumFish - 1, NumSteps) As Double

        Age = 3

        '- Landed Catch
        For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
                For Stk = 1 To NumStk
                    FishSumLC(Fish - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    FishSumLC(Fish - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                Next
            Next
        Next
        For Fish = 0 To NumFish - 1
            For TStep = 0 To NumSteps
                FishSumLC(Fish, TStep) = CLng(FishSumLC(Fish, TStep))
            Next
        Next
        'Transfer LandedCatch array to the worksheet starting at cell B12.
        xlWorkSheet.Range("B12").Resize(NumFish, NumSteps + 1).Value = FishSumLC

        '- Dropoff + Shakers
        For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
                For Stk = 1 To NumStk
                    FishSumSH(Fish - 1, TStep - 1) += (DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep))
                    FishSumSH(Fish - 1, NumSteps) += (DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep))
                Next
            Next
        Next
        For Fish = 0 To NumFish - 1
            For TStep = 0 To NumSteps
                FishSumSH(Fish, TStep) = CLng(FishSumSH(Fish, TStep))
            Next
        Next
        'Transfer Shaker+DropOff array to the worksheet starting at cell B221.
        xlWorkSheet.Range("B221").Resize(NumFish, NumSteps + 1).Value = FishSumSH

        '- NonRetention including MSF
        For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
                For Stk = 1 To NumStk
                    FishSumNR(Fish - 1, TStep - 1) += (NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))
                    FishSumNR(Fish - 1, NumSteps) += (NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))
                Next
            Next
        Next
        For Fish = 0 To NumFish - 1
            For TStep = 0 To NumSteps
                FishSumNR(Fish, TStep) = CLng(FishSumNR(Fish, TStep))
            Next
        Next
        'Transfer NonRetention array to the worksheet starting at cell B430.
        xlWorkSheet.Range("B430").Resize(NumFish, NumSteps + 1).Value = FishSumNR

        '- Catch + NonRetention (Old PFMC Standard .. No Longer Used but needed as place holder in Excel)
        For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
                For Stk = 1 To NumStk
                    FishSumCC(Fish - 1, TStep - 1) += (NonRetention(Stk, Age, Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                    FishSumCC(Fish - 1, NumSteps) += (NonRetention(Stk, Age, Fish, TStep) + LandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                Next
            Next
        Next
        For Fish = 0 To NumFish - 1
            For TStep = 0 To NumSteps
                FishSumCC(Fish, TStep) = CLng(FishSumCC(Fish, TStep))
            Next
        Next
        'Transfer NonRetention+LandedCatch array to the worksheet starting at cell B639.
        xlWorkSheet.Range("B639").Resize(NumFish, NumSteps + 1).Value = FishSumCC

        '- Total Mortality
        For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
                For Stk = 1 To NumStk
                    FishSumTM(Fish - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                    FishSumTM(Fish - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                Next
            Next
        Next
        For Fish = 0 To NumFish - 1
            For TStep = 0 To NumSteps
                FishSumTM(Fish, TStep) = CLng(FishSumTM(Fish, TStep))
            Next
        Next
        'Transfer TotalMortality array to the worksheet starting at cell B848.
        xlWorkSheet.Range("B848").Resize(NumFish, NumSteps + 1).Value = FishSumTM

        '===============================================================================
        xlWorkSheet = xlWorkBook.Sheets("Table2PRN")

        '- Get Instructions for Table2 WorkSheet from ReportDriver Table
        '- Read User Selected ReportDriver Data
        Dim ParseOld, ParseNew, RepStkNum, NumRepStks As Integer
        Dim CmdStr As String
        Dim Option1, Option2, Option3, Option4, Option5, Option6 As String
        Dim ReportNumber, RecNum As Integer
        Dim RngVal1, RngVal2, RngVal3 As String
        CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & "PSCTable2.Drv" & Chr(34) & " ORDER BY ReportNumber,Option5"
        Dim RDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim ReportDA As New System.Data.OleDb.OleDbDataAdapter
        ReportDA.SelectCommand = RDcm
        Dim RDcb As New OleDb.OleDbCommandBuilder
        RDcb = New OleDb.OleDbCommandBuilder(ReportDA)
        If FramDataSet.Tables.Contains("ReportDriver") Then
            FramDataSet.Tables("ReportDriver").Rows.Clear()
        End If
        ReportDA.Fill(FramDataSet, "ReportDriver")
        Dim NumRD As Integer
        NumRD = FramDataSet.Tables("ReportDriver").Rows.Count
        If NumRD = 0 Then
            MsgBox("ReportDriver Table Must have PSCTable2.Drv to do TAMM TRANSFER!!!!", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        NumRepGrps = 0
        Option1 = ""
        Option2 = ""
        Option3 = ""
        Option4 = ""
        Option5 = ""
        Option6 = ""

        '- Loop through Table Records for Actual Values
        For RecNum = 0 To NumRD - 1
            ReportNumber = FramDataSet.Tables("ReportDriver").Rows(RecNum)(2)
            If ReportNumber <> 3 Then
                MsgBox("Problem with PSCTable2.Drv - Wrong Report Number!", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(3)) Then
                Option1 = ""
            Else
                Option1 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(3)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(4)) Then
                Option2 = ""
            Else
                Option2 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(4)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(5)) Then
                Option3 = ""
            Else
                Option3 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(5)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(6)) Then
                Option4 = ""
            Else
                Option4 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(6)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(7)) Then
                Option5 = ""
            Else
                Option5 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(7)
                
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(8)) Then
                Option6 = ""
            Else
                Option6 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(8)
            End If
            'Dim ParseOld, ParseNew, RepStkNum, NumRepStks As Integer
            MortalityType = CInt(Option1)
            ParseOld = 1
            RepStkNum = 1
            NumRepStks = CInt(Option2)
            Dim RepStocks(NumRepStks) As Integer
            Dim RepGroupName As String
            If NumRepStks < 50 Then
                ParseOld = 1
                For Stk = 1 To NumStk
                    ParseNew = InStr(ParseOld, Option3, ",")
                    If ParseNew = 0 Then
                        RepStocks(Stk) = CInt(Option3.Substring(ParseOld - 1, Option3.Length - ParseOld + 1))
                        NumRepStks = Stk
                        Exit For
                    Else
                        RepStocks(Stk) = CInt(Option3.Substring(ParseOld - 1, ParseNew - ParseOld))
                    End If
                    ParseOld = ParseNew + 1
                Next
            Else
                For Stk = 1 To NumStk
                    If CInt(Option3.Substring(Stk - 1, 1)) = 1 Then
                        RepStocks(RepStkNum) = Stk
                        RepStkNum += 1
                    End If
                Next
            End If
            RepGroupName = Option4

            '- Total Mortality
            Dim TotMort(NumFish - 1, NumSteps)
            For RepStkNum = 1 To NumRepStks
                Stk = RepStocks(RepStkNum)
                If Stk > NumStk Then
                    MsgBox("The PSCTable2.DRV in your database has a STOCK ERROR!" & vbCrLf & "The DRV is from an old Base Period that had more stocks" & vbCrLf & _
                           "You must DELETE your current DRV and read the most recent file" & vbCrLf & "TAMM Transfer Aborted !!!", MsgBoxStyle.OkOnly)
                    Exit Sub
                End If
                For Fish = 1 To NumFish
                    For TStep = 1 To NumSteps
                        TotMort(Fish - 1, TStep - 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                        TotMort(Fish - 1, NumSteps) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                    Next
                Next
            Next
            For Fish = 0 To NumFish - 1
                For TStep = 0 To NumSteps
                    TotMort(Fish, TStep) = CLng(TotMort(Fish, TStep))
                Next
            Next
            'Transfer Stock TotalMortality array to the Table2 worksheet
            RngVal1 = "B" & (RecNum * 210 + 12).ToString
            RngVal2 = "C" & (RecNum * 210 + 13).ToString & ":" & "H" & (RecNum * 210 + 210).ToString
            RngVal3 = "B" & (RecNum * 210 + 12).ToString & ":" & "G" & (RecNum * 210 + 209).ToString
            'xlWorkSheet.Range(RngVal1).Resize(NumFish + 1, NumSteps + 2).Value = TotMort
            'xlWorkSheet.Range(RngVal2).Cut()
            'xlWorkSheet.Paste(Destination:=xlWorkSheet.Range(RngVal3))
            xlWorkSheet.Range(RngVal1).Resize(NumFish, NumSteps + 1).Value = TotMort

            '- Print Header Information for each Stock Catch Report
            RngVal1 = "A" & (RecNum * 210 + 2).ToString
            RngVal2 = "J" & (RecNum * 210 + 7).ToString
            RngVal3 = RngVal1 & ":" & RngVal2
            xlWorkSheet.Range(RngVal3).Clear()
            RngVal1 = "A" & (RecNum * 210 + 2).ToString
            xlWorkSheet.Range(RngVal1).Value = "Species: COHO"
            RngVal1 = "A" & (RecNum * 210 + 3).ToString
            xlWorkSheet.Range(RngVal1).Value = "Report: Stock Catch Summary"
            RngVal1 = "A" & (RecNum * 210 + 5).ToString
            xlWorkSheet.Range(RngVal1).Value = "Stock:" & RepGroupName
            RngVal1 = "A" & (RecNum * 210 + 7).ToString
            xlWorkSheet.Range(RngVal1).Value = "Total Mortality by Fishery"
            RngVal1 = "C" & (RecNum * 210 + 2).ToString
            xlWorkSheet.Range(RngVal1).Value = "Ver:" & FramVersion
            RngVal1 = "A" & (RecNum * 210 + 4).ToString
            xlWorkSheet.Range(RngVal1).Value = "Title:" & RunIDTitleSelect
            RngVal1 = "G" & (RecNum * 210 + 2).ToString
            xlWorkSheet.Range(RngVal1).Value = "Date:" & Now().Date
            RngVal1 = "G" & (RecNum * 210 + 3).ToString
            xlWorkSheet.Range(RngVal1).Value = "Time:" & Now().ToString("hh:ss tt", ci)

        Next

        '===============================================================================
        '- Stock Summary PRN Transfer
        xlWorkSheet = xlWorkBook.Sheets("StockSumPRN")
        xlWorkSheet.Range("A2:J6").Clear()
        xlWorkSheet.Range("A2").Value = "Species: COHO"
        xlWorkSheet.Range("A3").Value = "Report: Stock Summary"
        xlWorkSheet.Range("A7").Value = "Total Mortality by Stock for All Fisheries and Time Steps"
        xlWorkSheet.Range("C2").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A4").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G2").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G3").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G2").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G3").Value = "RepDate:" & Now().ToString

        Dim StockTotalMort(NumStk + 1, NumFish + 1), StkTotMort As Double
        Dim BegStk, EndStk, Page, PageStk As Integer
        Dim LastPage As Boolean

        '- Sum Total Mortality for All Stocks, Fisheries, and Time Steps ---
        BegStk = 1
        EndStk = 10
        LastPage = False
        For Fish = 1 To NumFish
            For Stk = 1 To NumStk
                For TStep = 1 To NumSteps
                    For Age = MinAge To MaxAge
                        StkTotMort = (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                        StockTotalMort(Stk, Fish) += StkTotMort
                        StockTotalMort(Stk, NumFish + 1) += StkTotMort
                        StockTotalMort(NumStk + 1, Fish) += StkTotMort
                    Next
                Next
            Next
        Next
        Dim TransArray(NumFish, 9)
        For Page = 1 To 100
            BegStk = Page * 10 - 9
            EndStk = Page * 10
            If EndStk >= NumStk Then EndStk = NumStk
            If BegStk > NumStk Then Exit For
            If BegStk <> EndStk - 9 Then
                ReDim TransArray(NumFish, EndStk - BegStk + 2)
                LastPage = True
            End If
            PageStk = 0
            For Stk = BegStk To EndStk
                For Fish = 1 To NumFish + 1
                    TransArray(Fish - 1, PageStk) = CLng(StockTotalMort(Stk, Fish))
                Next
                PageStk += 1
            Next
            If LastPage = True Then
                For Fish = 1 To NumFish
                    TransArray(Fish - 1, PageStk) = CLng(StockTotalMort(NumStk + 1, Fish))
                Next
            End If
            'Transfer Stock Total Mortality array to StockSumPRN worksheet
            RngVal1 = "B" & (Page * 205 - 195).ToString
            If BegStk <> EndStk - 9 Then
                xlWorkSheet.Range(RngVal1).Resize(NumFish + 1, EndStk - BegStk + 2).Value = TransArray
            Else
                xlWorkSheet.Range(RngVal1).Resize(NumFish + 1, 10).Value = TransArray
            End If
        Next

        '===============================================================================
        xlWorkSheet = xlWorkBook.Sheets("TRunsPRN")
        Dim SelStk, SelFish As Integer

        '- Get Instructions for Terminal Run Report WorkSheet from ReportDriver Table
        '- Read User Selected ReportDriver Data
        CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & "PSCTRuns.Drv" & Chr(34) & " ORDER BY ReportNumber,Option6"
        Dim TRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim ReportTR As New System.Data.OleDb.OleDbDataAdapter
        ReportTR.SelectCommand = TRcm
        Dim TRcb As New OleDb.OleDbCommandBuilder
        TRcb = New OleDb.OleDbCommandBuilder(ReportTR)
        If FramDataSet.Tables.Contains("ReportDriver") Then
            FramDataSet.Tables("ReportDriver").Rows.Clear()
        End If
        ReportTR.Fill(FramDataSet, "ReportDriver")
        'Dim NumRD As Integer
        NumRD = FramDataSet.Tables("ReportDriver").Rows.Count
        If NumRD = 0 Then
            MsgBox("ReportDriver Table Must have PSCTRuns.Drv to do TAMM TRANSFER!!!!", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Option1 = ""
        Option2 = ""
        Option3 = ""
        Option4 = ""
        Option5 = ""
        Option6 = ""

        '- Loop through Table Records for Actual Values
        Dim CohoTermRun(NumRD - 1, 2) As Object
        ReDim RepStks(NumRD - 1, NumStk)
        ReDim RepFish(NumRD - 1, NumFish)
        ReDim RepTStep(NumRD - 1, NumSteps)
        ReDim RepGrpType(NumRD - 1)
        ReDim RepGrpName(NumRD - 1)
        Dim NumGrpStks(NumRD - 1) As Integer
        ReDim NumRepFish(NumRD - 1)

        For RecNum = 0 To NumRD - 1
            ReportNumber = FramDataSet.Tables("ReportDriver").Rows(RecNum)(2)
            If ReportNumber <> 2 Then
                MsgBox("Problem with PSCTRuns.Drv - Wrong Report Number!", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(3)) Then
                Option1 = ""
            Else
                Option1 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(3)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(4)) Then
                Option2 = ""
            Else
                Option2 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(4)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(5)) Then
                Option3 = ""
            Else
                Option3 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(5)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(6)) Then
                Option4 = ""
            Else
                Option4 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(6)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(7)) Then
                Option5 = ""
            Else
                Option5 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(7)
            End If
            If IsDBNull(FramDataSet.Tables("ReportDriver").Rows(RecNum)(8)) Then
                Option6 = ""
            Else
                Option6 = FramDataSet.Tables("ReportDriver").Rows(RecNum)(8)
            End If
            '- Parse Option Strings from Report Driver Table Fields until All Groups are Read
            ParseOld = 1
            '- Stock Group
            For Stk = 1 To NumStk
                ParseNew = InStr(ParseOld, Option1, ",")
                If ParseNew = 0 Then
                    RepStks(RecNum, Stk) = CInt(Option1.Substring(ParseOld - 1, Option1.Length - ParseOld + 1))
                    NumGrpStks(RecNum) = Stk
                    Exit For
                Else
                    RepStks(RecNum, Stk) = CInt(Option1.Substring(ParseOld - 1, ParseNew - ParseOld))
                End If
                ParseOld = ParseNew + 1
            Next
            ParseOld = 1
            '- Fishery Group .. Can be Zero for Escapement Only
            If Option2 = "" Then
                NumRepFish(RecNum) = 0
                GoTo SkipFishGroup
            End If
            For Fish = 1 To NumFish
                ParseNew = InStr(ParseOld, Option2, ",")
                If ParseNew = 0 Then
                    RepFish(RecNum, Fish) = CInt(Option2.Substring(ParseOld - 1, Option2.Length - ParseOld + 1))
                    NumRepFish(RecNum) = Fish
                    Exit For
                Else
                    RepFish(RecNum, Fish) = CInt(Option2.Substring(ParseOld - 1, ParseNew - ParseOld))
                End If
                ParseOld = ParseNew + 1
            Next
SkipFishGroup:
            ParseOld = 1
            '- Terminal Time Steps
            For TStep = 1 To NumSteps
                ParseNew = InStr(ParseOld, Option3, ",")
                If ParseNew = 0 Then
                    RepTStep(RecNum, TStep) = CInt(Option3.Substring(ParseOld - 1, Option3.Length - ParseOld + 1))
                    Exit For
                Else
                    RepTStep(RecNum, TStep) = CInt(Option3.Substring(ParseOld - 1, ParseNew - ParseOld))
                End If
                ParseOld = ParseNew + 1
            Next
            RepGrpType(RecNum) = Option4
            RepGrpName(RecNum) = Option5

            '--- Normal Terminal Run Report Style
            '- First Get Escapement
            Age = 3
            For TStep = 1 To NumSteps
                For SelStk = 1 To NumGrpStks(RecNum)
                    Stk = RepStks(RecNum, SelStk)
                    If Stk > NumStk Then
                        MsgBox("The PSCTruns.DRV in your database has a STOCK ERROR!" & vbCrLf & "The DRV is from an old Base Period that had more stocks" & vbCrLf & _
                               "You must DELETE your current DRV and read the most recent file" & vbCrLf & "TAMM Transfer Aborted !!!", MsgBoxStyle.OkOnly)
                        Exit Sub
                    End If
                    CohoTermRun(RecNum, 0) += Escape(Stk, Age, TStep)
                    CohoTermRun(RecNum, 1) += Escape(Stk, Age, TStep)     
                    CohoTermRun(RecNum, 2) += Escape(Stk, Age, TStep)


                Next
            Next
            '- Next Get Catch 
            If NumRepFish(RecNum) = 0 Then GoTo SkipTermCatch2
            If RepGrpType(RecNum) = "TAA" Then
                '- TAA = All Stocks in Terminal Fishery 
                For Stk = 1 To NumStk
                    For TStep = RepTStep(RecNum, 1) To RepTStep(RecNum, 2)
                        For SelFish = 1 To NumRepFish(RecNum)
                            Fish = RepFish(RecNum, SelFish)
                            If Fish > NumFish Then
                                MsgBox("The PSCTruns.DRV in your database has a FISHERY ERROR!" & vbCrLf & "The DRV is from an old Base Period that had more fisheries" & vbCrLf & _
                                       "You must DELETE your current DRV and read the most recent file" & vbCrLf & "TAMM Transfer Aborted !!!", MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If
                            CohoTermRun(RecNum, 0) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                            CohoTermRun(RecNum, 2) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + _
                                                       MSFDropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        Next SelFish
                    Next TStep
                Next Stk
            Else
                '- ETRS = Catch of Local Stock Only in Terminal Fishery
                For SelStk = 1 To NumGrpStks(RecNum)
                    Stk = RepStks(RecNum, SelStk)
                    For TStep = RepTStep(RecNum, 1) To RepTStep(RecNum, 2)
                        For SelFish = 1 To NumRepFish(RecNum)
                            Fish = RepFish(RecNum, SelFish)
                            If Fish > NumFish Then
                                MsgBox("The PSCTruns.DRV in your database has a FISHERY ERROR!" & vbCrLf & "The DRV is from an old Base Period that had more fisheries" & vbCrLf & _
                                       "You must DELETE your current DRV and read the most recent file" & vbCrLf & "TAMM Transfer Aborted !!!", MsgBoxStyle.OkOnly)
                                Exit Sub
                            End If
                            CohoTermRun(RecNum, 0) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                            CohoTermRun(RecNum, 2) += (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + _
                                                       MSFDropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))

                        Next
                    Next
                Next
            End If
SkipTermCatch2:

        Next

        xlWorkSheet.Range("A2:J6").Clear()
        xlWorkSheet.Range("A2").Value = "Species: COHO"
        xlWorkSheet.Range("A3").Value = "Report: Terminal Run Report"
        xlWorkSheet.Range("C2").Value = "Ver:" & FramVersion
        xlWorkSheet.Range("A4").Value = "Title:" & RunIDTitleSelect
        'xlWorkSheet.Range("G2").Value = "Date:" & Now().Date
        'xlWorkSheet.Range("G3").Value = "Time:" & Now().ToString("hh:ss tt", ci)
        xlWorkSheet.Range("G2").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        xlWorkSheet.Range("G3").Value = "RepDate:" & Now().ToString

        For Stk = 0 To NumRD - 1
            '- Convert Terminal Run sizes to Long Integer
            CohoTermRun(Stk, 0) = CLng(CohoTermRun(Stk, 0))
            CohoTermRun(Stk, 1) = CLng(CohoTermRun(Stk, 1))
            'CohoTermRun(Stk, 2) = CLng(CohoTermRun(Stk, 2))
            'CohoTRS(Stk) = CLng(CohoTRS(Stk))
            '- Put TermGrpNames into TRuns WorkSheet
            RngVal1 = "A" & (Stk + 11).ToString
            xlWorkSheet.Range(RngVal1).Value = RepGrpName(Stk)
        Next
        'Transfer TermRun array to the TRunsPRN worksheet
        xlWorkSheet.Range("B11").Resize(NumRD, 3).Value = CohoTermRun
        'xlWorkSheet.Range("E11").Resize(NumRD).Value = CohoTRS

        ''- Done with TAMM WorkBook for this run .. Close and release object
        'xlApp.Application.DisplayAlerts = False

        ''xlWorkBook.SaveAs()

        'xlWorkBook.Save()

        'If WorkBookWasNotOpen = True Then
        '   xlWorkBook.Close()
        'End If
        'If ExcelWasNotRunning = True Then
        '   xlApp.Application.Quit()
        '   xlApp.Quit()
        'Else
        '   xlApp.Visible = True
        '   xlApp.WindowState = Excel.XlWindowState.xlMinimized
        'End If
        'xlApp.Application.DisplayAlerts = True
        'xlApp = Nothing

        '- test Excel without saving or closing
        xlApp.Visible = True
        xlApp.Application.DisplayAlerts = True
        xlApp.Application.Interactive = True

    End Sub

    Sub BYERReport()

        'Dim BY, Cost, BYAge

        Dim TerminalType As Integer
        Dim KTime, MeanSize, SizeStdDev As Double
        Dim LegalPopulation, SubLegalPopulation, EncounterRate As Double

        '- Open Text File to Save estimates of Cohort Sizes for Brood Years less than FRAM Ages
        '- These Cohort Size Estimates are calculated from Survival, Maturation, and ER rates

        File_Name = FVSdatabasepath & "\BY-Cohort-Compare-FramVS.txt"
        If Exists(File_Name) Then Delete(File_Name)

        Dim BYsw As StreamWriter = CreateText(File_Name)
        Dim BYsb As New StringBuilder

        PrnLine = "Command File =" + FVSdatabasepath + "\" & RunIDNameSelect.ToString & " RunDate:" & RunIDRunTimeDateSelect.ToString & " RepDate:" & Date.Today.ToString
        'xlWorkSheet.Range("G2").Value = "RunDate:" & RunIDRunTimeDateSelect.ToString
        'xlWorkSheet.Range("G3").Value = "RepDate:" & Now().ToString
        BYsw.WriteLine(PrnLine)
        BYsw.WriteLine(" ")

        '- First Dimension is Brood Year
        ReDim BYCohort(5, NumStk, MaxAge, 4, NumSteps - 1)
        ReDim BYEscape(5, NumStk, MaxAge, NumSteps - 1)
        ReDim BYLandedCatch(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYNonRetention(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYShakers(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYDropOff(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYMSFLandedCatch(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYMSFNonRetention(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYMSFShakers(5, NumStk, MaxAge, NumFish, NumSteps - 1)
        ReDim BYMSFDropOff(5, NumStk, MaxAge, NumFish, NumSteps - 1)

        '- The concept of this report is to calculate the Brood Year AEQ ER assuming this year's
        '- fishery harvest rate pattern.  The FRAM algorithms are applied to each Age class in
        '- forward and reverse to estimate a BYER for each age class.  Only Time Step 1 to 3 are used

        '- Fill BY arrays with FRAM values from this run (i.e. FRAM Age 2 for Brood Year Age 2, etc.)
        TerminalType = 0

        '- ReSet Cohort Sizes ... TAMM Calculations change array values

        If OptionChinookBYAEQ = 1 Then  '--- Called from Run Menu

            '- Put cohort sizes into BYCohort array
            For Stk = 1 To NumStk
                For Age = MinAge To MaxAge
                    For Cost = 0 To 4
                        For TStep = 1 To NumSteps - 1
                            BYCohort(Age, Stk, Age, Cost, TStep) = Cohort(Stk, Age, Cost, TStep)
                        Next TStep
                        '- Special Case for Age 2 Cohorts
                        If Age = 2 And Cost = 0 Then
                            If NumStk > 50 Then
                                '- SF Version ... check for Zero StockRecruitrs and ReSet to 0.1
                                If (Stk >= 7 And Stk <= 10) Or (Stk >= 13 And Stk <= 16) Then
                                    Jim = 1
                                Else
                                    If BYCohort(Age, Stk, Age, Cost, 1) = 0 Then
                                        BYCohort(Age, Stk, Age, Cost, 1) = BaseCohortSize(Stk, Age) * 0.1
                                    End If
                                End If
                            Else
                                '- Old Base Version ... Use Base Period Cohort Sizes because of Zero StockRecruits
                                If Stk = 4 Or Stk = 5 Or Stk = 7 Or Stk = 8 Then
                                    Jim = 1
                                Else
                                    BYCohort(Age, Stk, Age, Cost, 1) = BaseCohortSize(Stk, Age)
                                End If
                            End If
                        End If
                    Next Cost
                Next Age
            Next Stk
            '- Put Mortalities into BY Arrays
            For Stk = 1 To NumStk
                For Age = MinAge To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps - 1
                            BYLandedCatch(Age, Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
                            BYNonRetention(Age, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep)
                            BYShakers(Age, Stk, Age, Fish, TStep) = Shakers(Stk, Age, Fish, TStep)
                            BYDropOff(Age, Stk, Age, Fish, TStep) = DropOff(Stk, Age, Fish, TStep)
                            BYMSFLandedCatch(Age, Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep)
                            BYMSFNonRetention(Age, Stk, Age, Fish, TStep) = MSFNonRetention(Stk, Age, Fish, TStep)
                            BYMSFShakers(Age, Stk, Age, Fish, TStep) = MSFShakers(Stk, Age, Fish, TStep)
                            BYMSFDropOff(Age, Stk, Age, Fish, TStep) = MSFDropOff(Stk, Age, Fish, TStep)
                        Next TStep
                    Next Fish
                Next Age
            Next Stk
            '- Put Escapement into BYEscape array
            For Stk = 1 To NumStk
                For Age = MinAge To MaxAge
                    For TStep = 1 To NumSteps - 1
                        BYEscape(Age, Stk, Age, TStep) = Escape(Stk, Age, TStep)
                    Next TStep
                Next Age
            Next Stk

        ElseIf OptionChinookBYAEQ = 2 Then   '-- BYFlag = 2   Called from StkCatRep for PFMC Brood Year AEQ-ER
            'ahb for FRAMVS also called by Utilities/Update Coweeman Sheets
            '- Put cohort sizes into BYCohort array
            For TStep = 1 To NumSteps - 1
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        For Cost = 0 To 4 'Cost 4 = Starting Cohort, Cost 0 is Cohort remaining in ocean after maturation
                           
                            BYCohort(Age, Stk, Age, Cost, TStep) = Cohort(Stk, Age, Cost, TStep)
                            ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                            'If Cohort(Stk, 2, 4, 1) = 0 Then
                            '    If Cohort(Stk, 3, 4, 1) = 0 Then
                            '        BYCohort(2, Stk, 2, Cost, TStep) = Cohort(Stk, 4, Cost, TStep)
                            '    Else
                            '        BYCohort(2, Stk, 2, Cost, TStep) = Cohort(Stk, 3, Cost, TStep)
                            '    End If                      
                            'End If
                            'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                        Next Cost
                    Next Age
                Next Stk
            Next TStep
            '- Put Mortalities into BY Arrays
            For Stk = 1 To NumStk
                'If BYCohort(2, Stk, 2, 4, 1) = 0 Then
                '    BYCohort(2, Stk, 2, 0, 1) = BaseCohortSize(Stk, 2)
                'End If

                For Age = MinAge To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps - 1
                            BYLandedCatch(Age, Stk, Age, Fish, TStep) = LandedCatch(Stk, Age, Fish, TStep)
                            BYNonRetention(Age, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep)
                            BYShakers(Age, Stk, Age, Fish, TStep) = Shakers(Stk, Age, Fish, TStep)
                            BYDropOff(Age, Stk, Age, Fish, TStep) = DropOff(Stk, Age, Fish, TStep)
                            BYMSFLandedCatch(Age, Stk, Age, Fish, TStep) = MSFLandedCatch(Stk, Age, Fish, TStep)
                            BYMSFNonRetention(Age, Stk, Age, Fish, TStep) = MSFNonRetention(Stk, Age, Fish, TStep)
                            BYMSFShakers(Age, Stk, Age, Fish, TStep) = MSFShakers(Stk, Age, Fish, TStep)
                            BYMSFDropOff(Age, Stk, Age, Fish, TStep) = MSFDropOff(Stk, Age, Fish, TStep)
                        Next TStep
                    Next Fish
                Next Age
            Next Stk
            '- Put Escapement into BYEscape array
            For Stk = 1 To NumStk
                For Age = MinAge To MaxAge
                    For TStep = 1 To NumSteps - 1
                        BYEscape(Age, Stk, Age, TStep) = Escape(Stk, Age, TStep)
                    Next TStep
                Next Age
            Next Stk

        End If


        '-================================================================================
        '- Forward FRAM algorithms for Brood Year Ages greater than FRAM Ages
        '- i.e. Ages 3,4,5 for Age Class 2, Ages 4,5 for Age Class 3, etc.
        '- First use Age Ending Cohort size for Age+1 Starting Cohort Size

        PrnLine = "Cohort sizes for Forward FRAM- BY,Stk,Age-1,TS1-4 Co"
        BYsw.WriteLine(PrnLine)
        If OptionChinookBYAEQ = 1 Then
            For BY = 2 To MaxAge - 1
                For Stk = 1 To NumStk
                    For Age = 3 To MaxAge
                        If Age = BY + 1 Then
                            BYCohort(BY, Stk, Age, 0, 1) = Cohort(Stk, Age - 1, 0, 3)
                        End If
                    Next Age
                Next Stk
            Next BY
        ElseIf OptionChinookBYAEQ = 2 Then
            BY = 2
            For Age = 3 To 5
                For Stk = 1 To NumStk
                    BYCohort(BY, Stk, Age, 0, 1) = BYCohort(BY, Stk, Age - 1, 0, 3)
                Next Stk
            Next Age
        End If

        '-================================================================================
        '- Forward FRAM Algorithms
        TerminalType = 0

        'PrnLine = "Natural Mortality Calculations- BY,Stk,Age,TS,StartCo,WorkCo"
        'BYsw.WriteLine(PrnLine)
        For BY = 2 To 4
            '- Note: Special Case ... Age 2 run with Base Period Cohort sizes
            '- Normally this loop would start at Age 3
            For BYAge = MinAge To 5


                If OptionChinookBYAEQ = 2 And BYAge = 2 And BY = 2 Then GoTo NextBroodYearAge

                If BYAge <= BY And (BYAge <> 2 And BY <> 2) Then GoTo NextBroodYearAge
                For TStep = 1 To NumSteps - 1
                   
                    TerminalType = 0
                    '- Call NatMort
                    For Stk = 1 To NumStk
                        '- Subtract Natural Mortality
                        BYCohort(BY, Stk, BYAge, 4, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep)
                        BYCohort(BY, Stk, BYAge, PTerm, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep) * (1 - NaturalMortality(BYAge, TStep))
                        BYCohort(BY, Stk, BYAge, 3, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep)
                    Next Stk

                    '- Call CompCatch(PTerm)
TermFishEntry:
                    For Fish = 1 To NumFish
                        If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
                            For Stk = 1 To NumStk
                                Call CompLegProp(Stk, BYAge, Fish, TerminalType)
                                '- Retention Fishery Scaler & Quota
                                '********************************SIZE LIMIT FIX**************************************************
                                If SizeLimitFix = True And MinSizeLimit(Fish, TStep) < ChinookBaseSizeLimit(Fish, TStep) Then
                                    BYSizeLimitFixLanded(Fish, TerminalType)
                                Else
                                    '*********************************************************************************************
                                    If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                        If TStep = 3 And Fish = 54 Then
                                            TStep = 3
                                        End If
                                        BYLandedCatch(BY, Stk, BYAge, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, BYAge, Fish, TStep) * BYCohort(BY, Stk, BYAge, TerminalType, TStep) * FisheryScaler(Fish, TStep) * LegalProportion
                                    End If
                                    '- MSF Fishery Scaler & Quota
                                    If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                                        BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, BYAge, Fish, TStep) * BYCohort(BY, Stk, BYAge, TerminalType, TStep) * MSFFisheryScaler(Fish, TStep) * LegalProportion
                                        '--- Use Selective Incidental Rate on ALL fish encountered
                                        BYMSFDropOff(BY, Stk, BYAge, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep)
                                        '- All Stocks in Marked/UnMarked pairs
                                        If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                                            BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                                            BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                                        Else           '--- UnMarked (Wild) in Selective
                                            BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                                            BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                                        End If
                                    End If
                                    End If
                            Next Stk
                        End If
                    Next Fish

                    '- Call IncMort(PTerm)
                    For Fish = 1 To NumFish
                        ReDim PropSubPop(NumStk, MaxAge)
                        If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
                            '- Call CompShakers(Fish, TerminalType, EncRate, PropSubPop())
                            EncounterRate = 0
                            LegalPopulation = 0
                            SubLegalPopulation = 0
                            If TotalLandedCatch(Fish, TStep) > 0 Then
                                For Stk = 1 To NumStk
                                    '- Call CompLegProp(Stk, BYAge, Fish, TerminalType, SubLegalProportion, LegalProportion)
                                    KTime = (BYAge - 1) * 12 + MidTimeStep(TStep)
                                    MeanSize = VonBertL(Stk, TerminalType) * (1.0 - Exp(-VonBertK(Stk, TerminalType) * (KTime - VonBertT(Stk, TerminalType))))
                                    SizeStdDev = VonBertCV(Stk, BYAge, TerminalType) * MeanSize
                                    If (MinSizeLimit(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
                                        LegalProportion = 1
                                    End If
                                    If (MinSizeLimit(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
                                        LegalProportion = 0
                                    End If
                                    If ((MinSizeLimit(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (MinSizeLimit(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
                                        LegalProportion = (1 - NormlDistr(MinSizeLimit(Fish, TStep), MeanSize, SizeStdDev))
                                    End If
                                    SubLegalProportion = EncounterRateAdjustment(BYAge, Fish, TStep) * (1 - LegalProportion)
                                    LegalPopulation = LegalPopulation + BYCohort(BY, Stk, BYAge, TerminalType, TStep) * LegalProportion
                                    If NumStk < 50 And BYAge = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
                                        SubLegalPop = 0
                                    ElseIf NumStk > 50 And BYAge = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                                        SubLegalPop = 0
                                    Else
                                        SubLegalPop = BYCohort(BY, Stk, BYAge, TerminalType, TStep) * SubLegalProportion
                                    End If
                                    SubLegalPopulation = SubLegalPopulation + SubLegalPop

                                    If BYCohort(BY, Stk, BYAge, 3, TStep) > 0 Then
                                        '- Retention Fisheries
                                        '******************************************SIZE LIMI FIX******************************************
                                        If SizeLimitFix = True And MinSizeLimit(Fish, TStep) > ChinookBaseSizeLimit(Fish, TStep) Then
                                            BYSizeLimitFixShaker(Fish, TerminalType, EncounterRate)
                                        Else
                                            '*****************************************************************************************

                                            If Fish = 17 And Stk = 26 And TStep = 2 And BYAge = 4 Then
                                                TStep = 2
                                            End If

                                            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                                BYShakers(BY, Stk, BYAge, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                                            End If
                                            '- MSF Fisheries
                                            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                                                BYMSFShakers(BY, Stk, BYAge, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, BYAge, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                                            End If

                                            'If BY = 2 And Stk = 67 And TStep = 3 And BYShakers(BY, Stk, BYAge, Fish, TStep) <> 0 Then
                                            '   PrnLine = String.Format("SHK{0,2}", BY)
                                            '   PrnLine &= String.Format("{0,4}", Stk)
                                            '   PrnLine &= String.Format("{0,2}", BYAge)
                                            '   PrnLine &= String.Format("{0,4}", Fish)
                                            '   PrnLine &= String.Format("{0,2}", TStep)
                                            '   PrnLine &= String.Format("{0,11}", BYShakers(BY, Stk, BYAge, Fish, TStep).ToString("#####0.000"))
                                            '   PrnLine &= String.Format("{0,11}", FisheryScaler(Fish, TStep).ToString("####0.00"))
                                            '   PrnLine &= String.Format("{0,11}", SubLegalPop.ToString("#####0.00"))
                                            '   PrnLine &= String.Format("{0,11}", BYCohort(BY, Stk, BYAge, TerminalType, TStep).ToString("######0.0"))
                                            '   PrnLine &= String.Format("{0,11}", SubLegalProportion.ToString(" 0.000000"))
                                            '   BYsw.WriteLine(PrnLine)

                                            '   'PrnLine = String.Format("shk2221-{0,4}-{1,10}-{2,10}-{3,10}-{4,13}-{5,10}-{6,10}", Fish, BYShakers(BY, Stk, BYAge, Fish, TStep).ToString("#####0.00"), BYCohort(BY, Stk, BYAge, TerminalType, TStep).ToString("#####0.00"), SubLegalPopulation.ToString("#####0.00"), BaseSubLegalRate(Stk, BYAge, Fish, TStep).ToString("0.0000000000"), SubLegalProportion.ToString("#0.000000"), ShakerMortRate(Fish, TStep).ToString("#####0.00"))
                                            '   'BYsw.WriteLine(PrnLine)
                                            'End If

                                        End If
                                    End If
                                Next Stk


                                If LegalPopulation > 0 Then
                                    EncounterRate = SubLegalPopulation / LegalPopulation
                                Else
                                    EncounterRate = 0
                                End If
                            End If

                            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                'If FisheryFlag(Fish, TStep) < 6 Then
                                '- Call CompOthMort(Fish)
                                If TotalLandedCatch(Fish, TStep) > 0 Then
                                    For Stk = 1 To NumStk
                                        If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                            BYDropOff(BY, Stk, BYAge, Fish, TStep) = IncidentalRate(Fish, TStep) * BYLandedCatch(BY, Stk, BYAge, Fish, TStep)
                                        End If
                                    Next Stk
                                End If
                            End If
                            'If NonRetentionFlag(Fish, TStep) <> 0 Or (NonRetentionFlag(Fish, TStep) = 0 And (NonRetentionInput(Fish, TStep, 3) <> 0 Or NonRetentionInput(Fish, TStep, 4) <> 0)) Then
                            If NonRetentionFlag(Fish, TStep) <> 0 Then
                                '- Call CompCNR(Fish, TerminalType, EncRate, PropSubPop())
                                ReDim PropLegCatch(NumStk, MaxAge)
                                If TotalLandedCatch(Fish, TStep) = 0 And (NonRetentionFlag(Fish, TStep) = 1 Or NonRetentionFlag(Fish, TStep) = 2) Then
                                    GoTo NextBYFish
                                End If
                                Select Case NonRetentionFlag(Fish, TStep)
                                    Case 1                                   '...Computed CNR
                                        If FisheryScaler(Fish, TStep) < 1 Then
                                            For Stk = 1 To NumStk
                                                BYNonRetention(BY, Stk, BYAge, Fish, TStep) = BYLandedCatch(BY, Stk, BYAge, Fish, TStep) * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                                                BYNonRetention(BY, Stk, BYAge, Fish, TStep) = BYNonRetention(BY, Stk, BYAge, Fish, TStep) + TotalLandedCatch(Fish, TStep) * EncounterRate * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * PropSubPop(Stk, BYAge) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 3)
                                            Next Stk
                                        End If
                                    Case 2                  '...Ratio of CNR days to normal days
                                        For Stk = 1 To NumStk
                                            BYNonRetention(BY, Stk, BYAge, Fish, TStep) = BYLandedCatch(BY, Stk, BYAge, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                                            BYNonRetention(BY, Stk, BYAge, Fish, TStep) = BYNonRetention(BY, Stk, BYAge, Fish, TStep) + BYShakers(BY, Stk, BYAge, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * NonRetentionInput(Fish, TStep, 3)
                                        Next Stk
                                    Case 3                  '...External estimate of encounters
                                        '- Note: Can't do this method in "Forward" or "Reverse"
                                        '- Use FRAM Run estimates
                                        'AHB changed to Cohort from BYCohort to avoid divide by zero problem
                                        For Stk = 1 To NumStk
                                            If Cohort(Stk, BYAge, 3, TStep) > 0 Then
                                                BYNonRetention(BY, Stk, BYAge, Fish, TStep) = NonRetention(Stk, BYAge, Fish, TStep) * (BYCohort(BY, Stk, BYAge, 3, TStep) / Cohort(Stk, BYAge, 3, TStep))
                                                'If Fish = 22 And TStep = 3 And BYNonRetention(BY, Stk, BYAge, Fish, TStep) <> 0 Then
                                                'If BYNonRetention(BY, Stk, BYAge, Fish, TStep) <> 0 Then
                                                '   PrnLine = String.Format("CNR{0,2}", BY)
                                                '   PrnLine &= String.Format("{0,4}", Stk)
                                                '   PrnLine &= String.Format("{0,2}", BYAge)
                                                '   PrnLine &= String.Format("{0,4}", Fish)
                                                '   PrnLine &= String.Format("{0,2}", TStep)
                                                '   PrnLine &= String.Format("{0,11}", BYNonRetention(BY, Stk, BYAge, Fish, TStep).ToString("#####0.000"))
                                                '   PrnLine &= String.Format("{0,11}", NonRetention(Stk, BYAge, Fish, TStep).ToString("#####0.00"))
                                                '   PrnLine &= String.Format("{0,11}", BYCohort(BY, Stk, BYAge, 3, TStep).ToString("#####0.000"))
                                                '   PrnLine &= String.Format("{0,11}", Cohort(Stk, BYAge, 3, TStep).ToString("#####0.000"))
                                                '   BYsw.WriteLine(PrnLine)
                                                'End If
                                            Else
                                                BYNonRetention(BY, Stk, BYAge, Fish, TStep) = 0
                                            End If
                                        Next Stk
                                        'Call CompPropCatch(Fish, TerminalType)
                                        'For Stk = 1 To NumStk
                                        '   LegalProportion = PropLegCatch(Stk, BYAge)
                                        '   SubLegalProportion = PropSubPop(Stk, BYAge)
                                        '   '- PS Sport legal size rel mort rate set now to 50 of shaker release rate (10 vs 20)
                                        '   If Fish >= 36 And InStr(FisheryTitle$(Fish), "Sport") > 0 Then
                                        '      BYNonRetention(BY, Stk, BYAge, Fish, TStep) = LegalProportion * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * (ShakerMortRate(Fish, TStep) / 2)
                                        '   Else
                                        '      BYNonRetention(BY, Stk, BYAge, Fish, TStep) = LegalProportion * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep)
                                        '   End If
                                        '   If BYCohort(BY, Stk, BYAge, 3, TStep) > 0 Then
                                        '      BYNonRetention(BY, Stk, BYAge, Fish, TStep) = BYNonRetention(BY, Stk, BYAge, Fish, TStep) + SubLegalProportion * NonRetentionInput(Fish, TStep, 2) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep)
                                        '   Else
                                        '      BYNonRetention(BY, Stk, BYAge, Fish, TStep) = 0
                                        '   End If
                                        'Next Stk

                                        'If Fish = 22 And TStep = 3 Then
                                        '   For Stk = 1 To NumStk
                                        '      If BYNonRetention(BY, Stk, BYAge, Fish, TStep) <> 0 Then
                                        '         PrnLine = "BY-" & BY.ToString & "-" & Stk.ToString & "-" & Age.ToString & "-" & String.Format("{0,14}", BYNonRetention(BY, Stk, BYAge, Fish, TStep).ToString("#####0.0000000")) & "---" & StockTitle(Stk)
                                        '         BYsw.WriteLine(PrnLine)
                                        '      End If
                                        '   Next
                                        'End If

                                    Case 4    '--- Total Encounters Estimate (Legal + SubLegal)
                                        '---    Selective Fishery Sampling Estimates
                                        '- Note: Can't do this method in "Forward" or "Reverse"
                                        '- Use FRAM Run estimates
                                        For Stk = 1 To NumStk
                                            'AHB changed to Cohort from BYCohort to avoid divide by zero problem
                                            If Cohort(Stk, BYAge, 3, TStep) > 0 Then
                                                BYNonRetention(BY, Stk, BYAge, Fish, TStep) = NonRetention(Stk, BYAge, Fish, TStep) * (BYCohort(BY, Stk, BYAge, 3, TStep) / Cohort(Stk, BYAge, 3, TStep))
                                                'BYNonRetention(BY, Stk, BYAge, Fish, TStep) = NonRetention(Stk, BYAge, Fish, TStep)
                                            End If
                                        Next Stk
                                    Case Else
                                End Select
                            End If
                        End If
NextBYFish:
                    Next Fish
                    If TerminalType = 1 Then GoTo CalcEscape
                    '- Call Mature
                    For Stk = 1 To NumStk
                        For Fish = 1 To NumFish
                            If Stk = 24 And BY = 2 And BYAge = 3 And TStep = 3 Then
                                TStep = 3
                            End If
                            If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                                BYCohort(BY, Stk, BYAge, PTerm, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep) - BYLandedCatch(BY, Stk, BYAge, Fish, TStep) - BYShakers(BY, Stk, BYAge, Fish, TStep) - BYNonRetention(BY, Stk, BYAge, Fish, TStep) - BYDropOff(BY, Stk, BYAge, Fish, TStep) - BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) - BYMSFShakers(BY, Stk, BYAge, Fish, TStep) - BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) - BYMSFDropOff(BY, Stk, BYAge, Fish, TStep)
                            End If
                        Next Fish
                        BYCohort(BY, Stk, BYAge, 2, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep)
                        BYCohort(BY, Stk, BYAge, Term, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep)
                    Next Stk
                    For Stk = 1 To NumStk
                        BYCohort(BY, Stk, BYAge, Term, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep) * MaturationRate(Stk, BYAge, TStep)
                        BYCohort(BY, Stk, BYAge, PTerm, TStep) = BYCohort(BY, Stk, BYAge, PTerm, TStep) - BYCohort(BY, Stk, BYAge, Term, TStep)
                    Next Stk
                    '- Loop back for Terminal Fisheries on Mature Fish
                    If TerminalType = 0 Then
                        TerminalType = 1
                        GoTo TermFishEntry
                    End If
CalcEscape:
                    '- Call CompEscape
                    For Stk = 1 To NumStk
                        BYEscape(BY, Stk, BYAge, TStep) = BYCohort(BY, Stk, BYAge, TerminalType, TStep)
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = Term Then
                                BYEscape(BY, Stk, BYAge, TStep) = BYEscape(BY, Stk, BYAge, TStep) - BYLandedCatch(BY, Stk, BYAge, Fish, TStep) - BYShakers(BY, Stk, BYAge, Fish, TStep) - BYNonRetention(BY, Stk, BYAge, Fish, TStep) - BYDropOff(BY, Stk, BYAge, Fish, TStep) - BYMSFLandedCatch(BY, Stk, BYAge, Fish, TStep) - BYMSFShakers(BY, Stk, BYAge, Fish, TStep) - BYMSFNonRetention(BY, Stk, BYAge, Fish, TStep) - BYMSFDropOff(BY, Stk, BYAge, Fish, TStep)
                                If Escape(Stk, BYAge, TStep) < -1 Then
                                    AnyNegativeEscapement = 1
                                    'NegativeEsc(Stk, TStep) = 1
                                End If
                            End If
                        Next Fish
                    Next Stk
                    '- Call Savedat
                    '- Put Cohort Numbers into Next Time Step
                    For Stk = 1 To NumStk
                        '- BYAge Cohort to next BYAge or Time Step
                        If TStep < NumSteps - 1 Then
                            BYCohort(BY, Stk, BYAge, 0, TStep + 1) = BYCohort(BY, Stk, BYAge, 0, TStep)
                        Else
                            If BYAge <> 5 Then
                                BYCohort(BY, Stk, BYAge + 1, 0, 1) = BYCohort(BY, Stk, BYAge, 0, TStep)
                            End If
                        End If
                    Next Stk
                Next TStep
NextBroodYearAge:
            Next BYAge

            '- Special Case for Tulalip Bay Net
            If FisheryScaler(52, 3) = 0.99 Then
                If NumStk > 50 Then
                    '- SF Version
                    For Age = MinAge To 5
                        BYDropOff(BY, 19, Age, 52, 3) = (BYEscape(BY, 19, Age, 3) * (1 - FisheryScaler(51, 3))) * IncidentalRate(52, 3)
                        BYDropOff(BY, 20, Age, 52, 3) = (BYEscape(BY, 20, Age, 3) * (1 - FisheryScaler(51, 3))) * IncidentalRate(52, 3)
                        BYDropOff(BY, 19, Age, 51, 3) = (BYEscape(BY, 19, Age, 3) * FisheryScaler(51, 3)) * IncidentalRate(51, 3)
                        BYDropOff(BY, 20, Age, 51, 3) = (BYEscape(BY, 20, Age, 3) * FisheryScaler(51, 3)) * IncidentalRate(51, 3)
                        BYLandedCatch(BY, 19, Age, 52, 3) = (BYEscape(BY, 19, Age, 3) - BYDropOff(BY, 19, Age, 52, 3) - BYDropOff(BY, 19, Age, 51, 3)) * (1 - FisheryScaler(51, 3))
                        BYLandedCatch(BY, 20, Age, 52, 3) = (BYEscape(BY, 20, Age, 3) - BYDropOff(BY, 20, Age, 52, 3) - BYDropOff(BY, 20, Age, 51, 3)) * (1 - FisheryScaler(51, 3))
                        BYLandedCatch(BY, 19, Age, 51, 3) = (BYEscape(BY, 19, Age, 3) - BYDropOff(BY, 19, Age, 52, 3) - BYDropOff(BY, 19, Age, 51, 3)) * FisheryScaler(51, 3)
                        BYLandedCatch(BY, 20, Age, 51, 3) = (BYEscape(BY, 20, Age, 3) - BYDropOff(BY, 20, Age, 52, 3) - BYDropOff(BY, 20, Age, 51, 3)) * FisheryScaler(51, 3)
                        BYEscape(BY, 19, Age, 3) = 0
                        BYEscape(BY, 20, Age, 3) = 0
                    Next Age
                Else
                    '- Old Base w/o Marked/UnMarked
                    For Age = MinAge To 5
                        BYDropOff(BY, 10, Age, 52, 3) = (BYEscape(BY, 10, Age, 3) * (1 - FisheryScaler(51, 3))) * IncidentalRate(52, 3)
                        BYDropOff(BY, 10, Age, 51, 3) = (BYEscape(BY, 10, Age, 3) * FisheryScaler(51, 3)) * IncidentalRate(51, 3)
                        BYLandedCatch(BY, 10, Age, 52, 3) = (BYEscape(BY, 10, Age, 3) - BYDropOff(BY, 10, Age, 52, 3) - BYDropOff(BY, 10, Age, 51, 3)) * (1 - FisheryScaler(51, 3))
                        BYLandedCatch(BY, 10, Age, 51, 3) = (BYEscape(BY, 10, Age, 3) - BYDropOff(BY, 10, Age, 52, 3) - BYDropOff(BY, 10, Age, 51, 3)) * FisheryScaler(51, 3)
                        BYEscape(BY, 10, Age, 3) = 0
                    Next Age
                End If
            End If
            '- End Special Case Tulalip Bay Net
            If OptionChinookBYAEQ = 2 Then Exit For

        Next BY

        BYsw.WriteLine(" ")
        PrnLine = "Forward FRAM Cohort/Morts by,stk,age,ts,byco,ptmo,trmo"
        BYsw.WriteLine(PrnLine)

        Dim TempTotalMort(2) As Double
        For BY = 2 To 4
            For Stk = 1 To NumStk
                For Age = MinAge To 5
                    If Age < BY Then GoTo SkipFramAge
                    For TStep = 1 To NumSteps - 1
                        BYsb.AppendFormat("{0,3}", BY)
                        BYsb.AppendFormat("{0,3}", Stk)
                        BYsb.AppendFormat("{0,3}", Age)
                        BYsb.AppendFormat("{0,3}", TStep)
                        BYsb.AppendFormat("{0,11}", BYCohort(BY, Stk, Age, 4, TStep).ToString("#######0.0"))
                        ReDim TempTotalMort(2)
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                                TempTotalMort(1) = TempTotalMort(1) + (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                            Else
                                TempTotalMort(2) = TempTotalMort(2) + (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep))
                            End If
                        Next Fish
                        BYsb.AppendFormat("{0,11}", TempTotalMort(1).ToString("#######0.0"))
                        BYsb.AppendFormat("{0,11}", TempTotalMort(2).ToString("#######0.0"))
                        BYsb.AppendFormat("{0,11}", BYEscape(BY, Stk, Age, TStep).ToString("#######0.0"))
                        BYsb.Append(vbCrLf)
                    Next TStep
SkipFramAge:
                Next Age
            Next Stk
        Next BY

        If OptionChinookBYAEQ = 2 Then GoTo SkipReverse '--- Called from StkCatRep

        '
        '-============================================================================
        '- Reverse FRAM Algorithms for Brood Year Ages less than FRAM ages
        '
        '- Step 1 Calculate the Total Mortality ER using Cohort size as demoninator
        '-        for each Stk, Age, Time Step for the FRAM Ages

        Dim FramBYER(MaxAge, NumStk, MaxAge, NumSteps - 1) As Double

        '- Sum Mortalites
        For Stk = 1 To NumStk
            For Age = MinAge To MaxAge
                For TStep = 1 To NumSteps - 1
                    For Fish = 1 To NumFish
                        If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                            FramBYER(Age, Stk, Age, TStep) = FramBYER(Age, Stk, Age, TStep) + (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        Else
                            FramBYER(Age, Stk, Age, TStep) = FramBYER(Age, Stk, Age, TStep) + (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                        End If
                    Next Fish
                Next TStep
            Next Age
        Next Stk

        '- Print Mortalities
        BYsb.Append(vbCrLf)
        BYsb.Append("Mortality Summary - stk,age,tstep")
        BYsb.Append(vbCrLf)
        For Stk = 1 To NumStk
            For Age = MinAge To MaxAge
                For TStep = 1 To NumSteps - 1
                    If FramBYER(Age, Stk, Age, TStep) <> 0 Then
                        BYsb.AppendFormat("{0,3}", Stk)
                        BYsb.AppendFormat("{0,3}", Age)
                        BYsb.AppendFormat("{0,3}", TStep)
                        BYsb.AppendFormat("{0,11}", FramBYER(Age, Stk, Age, TStep).ToString("######0.0"))
                        BYsb.Append(vbCrLf)
                    End If
                Next TStep
            Next Age
        Next Stk

        '- Calculate Fishing Year ER using Cohort Size by Time Step
        BYsb.Append(vbCrLf)
        BYsb.Append("BYER Calculations - stk,age,tstep")
        BYsb.Append(vbCrLf)
        For Stk = 1 To NumStk
            For Age = MinAge To MaxAge
                For TStep = 1 To NumSteps - 1
                    If Cohort(Stk, Age, 3, TStep) = 0 Then
                        FramBYER(Age, Stk, Age, TStep) = 0
                    Else
                        FramBYER(Age, Stk, Age, TStep) = FramBYER(Age, Stk, Age, TStep) / Cohort(Stk, Age, 3, TStep)
                        If FramBYER(Age, Stk, Age, TStep) <> 0 Then
                            BYsb.AppendFormat("{0,3}", Stk)
                            BYsb.AppendFormat("{0,3}", Age)
                            BYsb.AppendFormat("{0,3}", TStep)
                            BYsb.AppendFormat("{0,11}", FramBYER(Age, Stk, Age, TStep).ToString("#0.000000"))
                            BYsb.Append(vbCrLf)
                        End If
                    End If
                Next TStep
            Next Age
        Next Stk

        '- Calculate BYCohort Sizes for Reverse FRAM
        BYsb.Append(vbCrLf)
        BYsb.Append("REVERSE BY Cohort Calculations - by,stk,age,tstep,new co,old co,byer,mr,sr")
        BYsb.Append(vbCrLf)
        For BY = 3 To MaxAge
            For Stk = 1 To NumStk
                For Age = MaxAge To 2 Step -1
                    If Age >= BY Then GoTo NextCohortAge
                    For TStep = (NumSteps - 1) To 1 Step -1

                        If TStep = (NumSteps - 1) Then
                            '- If Time Step 3, get Cohort From Time Step 1 Age + 1
                            If MaturationRate(Stk, Age, TStep) = 1 Then
                                BYCohort(BY, Stk, Age, PTerm, TStep) = (BYCohort(BY, Stk, Age + 1, PTerm, 1) / (1 - FramBYER(Age, Stk, Age, TStep))) / (1 - NaturalMortality(Age, TStep))
                            Else
                                BYCohort(BY, Stk, Age, PTerm, TStep) = ((BYCohort(BY, Stk, Age + 1, PTerm, 1) / (1 - MaturationRate(Stk, Age, TStep))) / (1 - FramBYER(Age, Stk, Age, TStep))) / (1 - NaturalMortality(Age, TStep))
                            End If
                        Else
                            '- Time Steps 1 or 2, get Cohort from Time Step + 1 Cohort, Same Age
                            If MaturationRate(Stk, Age, TStep) = 1 Then
                                BYCohort(BY, Stk, Age, PTerm, TStep) = (BYCohort(BY, Stk, Age, PTerm, TStep + 1) / (1 - FramBYER(Age, Stk, Age, TStep))) / (1 - NaturalMortality(Age, TStep))
                            Else
                                BYCohort(BY, Stk, Age, PTerm, TStep) = ((BYCohort(BY, Stk, Age, PTerm, TStep + 1) / (1 - MaturationRate(Stk, Age, TStep))) / (1 - FramBYER(Age, Stk, Age, TStep))) / (1 - NaturalMortality(Age, TStep))
                            End If
                        End If
                        BYsb.AppendFormat("{0,3}", BY)
                        BYsb.AppendFormat("{0,3}", Stk)
                        BYsb.AppendFormat("{0,3}", Age)
                        BYsb.AppendFormat("{0,3}", TStep)
                        BYsb.AppendFormat("{0,11}", BYCohort(BY, Stk, Age, PTerm, TStep).ToString("#######0.0"))
                        If TStep = (NumSteps - 1) Then
                            BYsb.AppendFormat("{0,11}", BYCohort(BY, Stk, Age + 1, PTerm, 1).ToString("#######0.0"))
                        Else
                            BYsb.AppendFormat("{0,11}", BYCohort(BY, Stk, Age, PTerm, TStep + 1).ToString("#######0.0"))
                        End If
                        BYsb.AppendFormat("{0,9}", FramBYER(Age, Stk, Age, TStep).ToString("0.000000"))
                        BYsb.AppendFormat("{0,7}", MaturationRate(Stk, Age, TStep).ToString("0.0000"))
                        BYsb.AppendFormat("{0,7}", (1 - NaturalMortality(Age, TStep)).ToString("0.0000"))
                        BYsb.Append(vbCrLf)
                    Next TStep
NextCohortAge:
                Next Age
            Next Stk
        Next BY

        BYsb.Append(vbCrLf)
        BYsb.Append("=========================================================================")
        BYsb.Append(vbCrLf)
        '- Forward FRAM Algorithms for REVERSE part of BY Matrix
        TerminalType = 0
        ' AHB this code has issues was missing portions of MSF shaker calculations (I added it back in), code also uses (BY, stk, BY) LandedCatch arrays 
        ' should probably be BY, Stk, Age
        BYsb.Append("Natural Mortality Calculations- BY,Stk,Age,TS,StartCo,WorkCo")
        BYsb.Append(vbCrLf)
        For BY = 2 To 5
            For Age = MinAge To 5
                If Age >= BY Then GoTo NextBroodYearAge2
                For TStep = 1 To NumSteps - 1
                    TerminalType = 0
                    '- Call NatMort
                    For Stk = 1 To NumStk
                        '- Subtract Natural Mortality
                        BYsb.AppendFormat("{0,3}", BY)
                        BYsb.AppendFormat("{0,3}", Stk)
                        BYsb.AppendFormat("{0,3}", Age)
                        BYsb.AppendFormat("{0,3}", TStep)
                        BYsb.AppendFormat("{0,9}", BYCohort(BY, Stk, Age, PTerm, TStep).ToString("#######0"))
                        BYsb.AppendFormat("{0,9}", (BYCohort(BY, Stk, Age, PTerm, TStep) * (1 - NaturalMortality(Age, TStep))).ToString("#######0"))
                        BYsb.Append(vbCrLf)
                        BYCohort(BY, Stk, Age, 4, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep)
                        BYCohort(BY, Stk, Age, PTerm, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep) * (1 - NaturalMortality(Age, TStep))
                        BYCohort(BY, Stk, Age, 3, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep)
                    Next Stk

                    '- Call CompCatch(PTerm)
TermFishEntry2:
                    For Fish = 1 To NumFish
                        If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
                            For Stk = 1 To NumStk
                                'Call CompLegProp(Stk, Age, Fish, TerminalType, SubLegalProportion, LegalProportion)
                                Call CompLegProp(Stk, Age, Fish, TerminalType)
                                '- Retention Fishery Scaler & Quota
                                '********************************SIZE LIMIT FIX**************************************************
                                'If SizeLimitFix = True And MinSizeLimit(Fish, TStep) < ChinookBaseSizeLimit(Fish, TStep) Then
                                '    BYSizeLimitFixLanded(Fish, TerminalType)
                                'Else
                                '*********************************************************************************************
                                If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then

                                    BYLandedCatch(BY, Stk, BY, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, BY, Fish, TStep) * BYCohort(BY, Stk, BY, TerminalType, TStep) * FisheryScaler(Fish, TStep) * LegalProportion
                                End If
                                '- MSF Fishery Scaler & Quota
                                If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                                    BYMSFLandedCatch(BY, Stk, BY, Fish, TStep) = StockFishRateScalers(Stk, Fish, TStep) * BaseExploitationRate(Stk, BY, Fish, TStep) * BYCohort(BY, Stk, BY, TerminalType, TStep) * MSFFisheryScaler(Fish, TStep) * LegalProportion
                                    '--- Use Selective Incidental Rate on ALL fish encountered
                                    BYMSFDropOff(BY, Stk, Age, Fish, TStep) = MarkSelectiveIncRate(Fish, TStep) * BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                                    '- All Stocks in Marked/UnMarked pairs
                                    If (Stk Mod 2) = 0 Then '--- Marked Fish in Selective
                                        BYMSFNonRetention(BY, Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) * MarkSelectiveMarkMisID(Fish, TStep) * MarkSelectiveMortRate(Fish, TStep)
                                        BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveMarkMisID(Fish, TStep))
                                    Else           '--- UnMarked (Wild) in Selective
                                        BYMSFNonRetention(BY, Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) * (1.0 - MarkSelectiveUnMarkMisID(Fish, TStep)) * MarkSelectiveMortRate(Fish, TStep)
                                        BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) * MarkSelectiveUnMarkMisID(Fish, TStep)
                                    End If
                                End If
                                'End If
                            Next Stk
                        End If
                    Next Fish

                    '- Call IncMort(PTerm)
                    For Fish = 1 To NumFish
                        ReDim PropSubPop(NumStk, MaxAge)
                        If TerminalFisheryFlag(Fish, TStep) = TerminalType Then
                            '- Call CompShakers(Fish, TerminalType, EncRate, PropSubPop())
                            EncounterRate = 0
                            LegalPopulation = 0
                            SubLegalPopulation = 0
                            If TotalLandedCatch(Fish, TStep) > 0 Then
                                For Stk = 1 To NumStk
                                    '- Call CompLegProp(Stk, age, Fish, TerminalType, SubLegalProportion, LegalProportion)
                                    KTime = (Age - 1) * 12 + MidTimeStep(TStep)
                                    MeanSize = VonBertL(Stk, TerminalType) * (1.0 - Exp(-VonBertK(Stk, TerminalType) * (KTime - VonBertT(Stk, TerminalType))))
                                    SizeStdDev = VonBertCV(Stk, Age, TerminalType) * MeanSize
                                    If (MinSizeLimit(Fish, TStep) < MeanSize - 3 * SizeStdDev) Then
                                        LegalProportion = 1
                                    End If
                                    If (MinSizeLimit(Fish, TStep) > MeanSize + 3 * SizeStdDev) Then
                                        LegalProportion = 0
                                    End If
                                    If ((MinSizeLimit(Fish, TStep) >= MeanSize - 3 * SizeStdDev) And (MinSizeLimit(Fish, TStep) <= MeanSize + 3 * SizeStdDev)) Then
                                        LegalProportion = (1 - NormlDistr(MinSizeLimit(Fish, TStep), MeanSize, SizeStdDev))
                                    End If
                                    'If Age <= MaxAgeEncAdtstep Then
                                    SubLegalProportion = EncounterRateAdjustment(Age, Fish, TStep) * (1 - LegalProportion)
                                    'Else
                                    'SubLegalProportion = (1 - LegalProportion)
                                    'End If
                                    LegalPopulation = LegalPopulation + BYCohort(BY, Stk, Age, TerminalType, TStep) * LegalProportion
                                    If NumStk < 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
                                        SubLegalPop = 0
                                    ElseIf NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
                                        SubLegalPop = 0
                                    Else
                                        SubLegalPop = BYCohort(BY, Stk, Age, TerminalType, TStep) * SubLegalProportion
                                    End If
                                    SubLegalPopulation = SubLegalPopulation + SubLegalPop
                                    If BYCohort(BY, Stk, Age, 3, TStep) > 0 Then

                                        If BYCohort(BY, Stk, Age, 3, TStep) > 0 Then
                                            '- Retention Fisheries
                                            '******************************************SIZE LIMI FIX******************************************
                                            'If SizeLimitFix = True And MinSizeLimit(Fish, TStep) > ChinookBaseSizeLimit(Fish, TStep) Then
                                            '    BYSizeLimitFixShaker(Fish, TerminalType, EncounterRate)
                                            'Else
                                            '*****************************************************************************************
                                            If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                                BYShakers(BY, Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                                            End If
                                            '- MSF Fisheries
                                            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 28 Then
                                                BYMSFShakers(BY, Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * SubLegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
                                            End If
                                        End If
                                    End If
                                    'End If
                                Next Stk

                                If LegalPopulation > 0 Then
                                    EncounterRate = SubLegalPopulation / LegalPopulation
                                Else
                                    EncounterRate = 0
                                End If

                            End If
                            If FisheryFlag(Fish, TStep) < 6 Then
                                '- Call CompOthMort(Fish)
                                If TotalLandedCatch(Fish, TStep) > 0 Then
                                    For Stk = 1 To NumStk
                                        If FisheryFlag(Fish, TStep) = 1 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 2 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                                            BYDropOff(BY, Stk, Age, Fish, TStep) = IncidentalRate(Fish, TStep) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                        End If
                                    Next Stk
                                End If
                            End If
                            If NonRetentionFlag(Fish, TStep) <> 0 Or (NonRetentionFlag(Fish, TStep) = 0 And (NonRetentionInput(Fish, TStep, 3) <> 0 Or NonRetentionInput(Fish, TStep, 4) <> 0)) Then
                                '- Call CompCNR(Fish, TerminalType, EncRate, PropSubPop())
                                ReDim PropLegCatch(NumStk, MaxAge)
                                If TotalLandedCatch(Fish, TStep) = 0 And (NonRetentionFlag(Fish, TStep) = 0 Or NonRetentionFlag(Fish, TStep) = 1) Then
                                    GoTo NextBYFish2
                                End If
                                Select Case NonRetentionFlag(Fish, TStep)
                                    Case 1                                   '...Computed CNR
                                        If FisheryScaler(Fish, TStep) < 1 Then
                                            For Stk = 1 To NumStk
                                                BYNonRetention(BY, Stk, Age, Fish, TStep) = BYLandedCatch(BY, Stk, Age, Fish, TStep) * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                                                BYNonRetention(BY, Stk, Age, Fish, TStep) = BYNonRetention(BY, Stk, Age, Fish, TStep) + TotalLandedCatch(Fish, TStep) * EncounterRate * ((1 - FisheryScaler(Fish, TStep)) / FisheryScaler(Fish, TStep)) * PropSubPop(Stk, Age) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 3)
                                            Next Stk
                                        End If
                                    Case 2                  '...Ratio of CNR days to normal days
                                        For Stk = 1 To NumStk
                                            BYNonRetention(BY, Stk, Age, Fish, TStep) = BYLandedCatch(BY, Stk, Age, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * ShakerMortRate(Fish, TStep) * NonRetentionInput(Fish, TStep, 4)
                                            BYNonRetention(BY, Stk, Age, Fish, TStep) = BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) * (NonRetentionInput(Fish, TStep, 1) / NonRetentionInput(Fish, TStep, 2)) * NonRetentionInput(Fish, TStep, 3)
                                        Next Stk
                                    Case 3                  '...External estimate of encounters
                                        '- Note: Can't do this method in "Forward" or "Reverse"
                                        '- Use FRAM Run estimates
                                        'AHB changed to Cohort from BYCohort to avoid divide by zero problem
                                        For Stk = 1 To NumStk
                                            If Cohort(Stk, Age, 3, TStep) > 0 Then
                                                BYNonRetention(BY, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep) * (BYCohort(BY, Stk, Age, 3, TStep) / Cohort(Stk, Age, 3, TStep))
                                                'BYNonRetention(BY, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep)
                                            End If
                                        Next Stk
                                        'Call CompPropCatch(Fish, TerminalType)
                                        'For Stk = 1 To NumStk
                                        '   LegalProportion = PropLegCatch(Stk, Age)
                                        '   SubLegalProportion = PropSubPop(Stk, Age)
                                        '   '- PS Sport legal size rel mort rate set now to 50 of shaker release rate (10 vs 20)
                                        '   If Fish >= 36 And InStr(FisheryTitle$(Fish), "Sport") > 0 Then
                                        '      BYNonRetention(BY, Stk, Age, Fish, TStep) = LegalProportion * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * (ShakerMortRate(Fish, TStep) / 2)
                                        '   Else
                                        '      BYNonRetention(BY, Stk, Age, Fish, TStep) = LegalProportion * NonRetentionInput(Fish, TStep, 1) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep)
                                        '   End If
                                        '   If BYCohort(BY, Stk, Age, 3, TStep) > 0 Then
                                        '      BYNonRetention(BY, Stk, Age, Fish, TStep) = BYNonRetention(BY, Stk, Age, Fish, TStep) + SubLegalProportion * NonRetentionInput(Fish, TStep, 2) * ModelStockProportion(Fish) * ShakerMortRate(Fish, TStep)
                                        '   Else
                                        '      BYNonRetention(BY, Stk, Age, Fish, TStep) = 0
                                        '   End If
                                        'Next Stk

                                        'If Fish = 22 And TStep = 3 Then
                                        '   For Stk = 1 To NumStk
                                        '      If BYNonRetention(BY, Stk, Age, Fish, TStep) <> 0 Then
                                        '         PrnLine = "BY-" & BY.ToString & "-" & Stk.ToString & "-" & Age.ToString & "-" & String.Format("{0,14}", BYNonRetention(BY, Stk, Age, Fish, TStep).ToString("#####0.0000000")) & "---" & StockTitle(Stk)
                                        '         BYsw.WriteLine(PrnLine)
                                        '      End If
                                        '   Next
                                        'End If

                                    Case 4    '--- Total Encounters Estimate (Legal + SubLegal)
                                        '---    Selective Fishery Sampling Estimates
                                        '- Note: Can't do this method in "Forward" or "Reverse"
                                        '- Use FRAM Run estimates
                                        '- AHB changed to Cohort from BYCohort to avoid divide by zero problem
                                        For Stk = 1 To NumStk
                                            If Cohort(Stk, Age, 3, TStep) > 0 Then
                                                BYNonRetention(BY, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep) * (BYCohort(BY, Stk, Age, 3, TStep) / Cohort(Stk, Age, 3, TStep))
                                                'BYNonRetention(BY, Stk, Age, Fish, TStep) = NonRetention(Stk, Age, Fish, TStep)
                                            End If
                                        Next Stk
                                    Case Else
                                End Select
                            End If
                        End If
NextBYFish2:
                    Next Fish
                    If TerminalType = 1 Then GoTo CalcEscape2
                    '- Call Mature
                    For Stk = 1 To NumStk
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                                BYCohort(BY, Stk, Age, PTerm, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep) - BYLandedCatch(BY, Stk, Age, Fish, TStep) - BYShakers(BY, Stk, Age, Fish, TStep) - BYNonRetention(BY, Stk, Age, Fish, TStep) - BYDropOff(BY, Stk, Age, Fish, TStep) - BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) - BYMSFShakers(BY, Stk, Age, Fish, TStep) - BYMSFNonRetention(BY, Stk, Age, Fish, TStep) - BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                            End If
                        Next Fish
                        BYCohort(BY, Stk, Age, 2, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep)
                        BYCohort(BY, Stk, Age, Term, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep)
                    Next Stk
                    For Stk = 1 To NumStk
                        BYCohort(BY, Stk, Age, Term, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep) * MaturationRate(Stk, Age, TStep)
                        BYCohort(BY, Stk, Age, PTerm, TStep) = BYCohort(BY, Stk, Age, PTerm, TStep) - BYCohort(BY, Stk, Age, Term, TStep)
                    Next Stk
                    '- Loop back for Terminal Fisheries on Mature Fish
                    If TerminalType = 0 Then
                        TerminalType = 1
                        GoTo TermFishEntry2
                    End If
CalcEscape2:
                    '- Call CompEscape
                    For Stk = 1 To NumStk
                        BYEscape(BY, Stk, Age, TStep) = BYCohort(BY, Stk, Age, TerminalType, TStep)
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = Term Then
                                BYEscape(BY, Stk, Age, TStep) = BYEscape(BY, Stk, Age, TStep) - BYLandedCatch(BY, Stk, Age, Fish, TStep) - BYShakers(BY, Stk, Age, Fish, TStep) - BYNonRetention(BY, Stk, Age, Fish, TStep) - BYDropOff(BY, Stk, Age, Fish, TStep) - BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) - BYMSFShakers(BY, Stk, Age, Fish, TStep) - BYMSFNonRetention(BY, Stk, Age, Fish, TStep) - BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                                If Escape(Stk, Age, TStep) < -1 Then
                                    AnyNegativeEscapement = 1
                                    'NegativeEsc(Stk, TStep) = 1
                                End If
                            End If
                        Next Fish
                    Next Stk
                    '- Call Savedat
                    '- Put Cohort Numbers into Next Time Step
                    For Stk = 1 To NumStk
                        '- Age Cohort to next Age or Time Step
                        If TStep < NumSteps - 1 Then
                            BYCohort(BY, Stk, Age, 0, TStep + 1) = BYCohort(BY, Stk, Age, 0, TStep)
                        Else
                            If Age <> 5 Then
                                BYCohort(BY, Stk, Age + 1, 0, 1) = BYCohort(BY, Stk, Age, 0, TStep)
                            End If
                        End If
                    Next Stk
                Next TStep
NextBroodYearAge2:
            Next Age
        Next BY

        BYsb.Append(vbCrLf)
        BYsb.Append("===Reverse FRAM Cohort/Morts by,stk,age,ts,byco,ptmo,trmo")
        BYsb.Append(vbCrLf)
        ReDim TempTotalMort(2)
        For BY = 2 To 5
            For Stk = 1 To NumStk
                For Age = MinAge To 5
                    If Age > BY Then GoTo SkipFramAge2
                    For TStep = 1 To NumSteps - 1
                        BYsb.AppendFormat("{0,3}", BY)
                        BYsb.AppendFormat("{0,3}", Stk)
                        BYsb.AppendFormat("{0,3}", Age)
                        BYsb.AppendFormat("{0,3}", TStep)
                        BYsb.AppendFormat("{0,11}", BYCohort(BY, Stk, Age, 4, TStep).ToString("#######0.0"))
                        ReDim TempTotalMort(2)
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                                TempTotalMort(1) = TempTotalMort(1) + (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                            Else
                                TempTotalMort(2) = TempTotalMort(2) + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                            End If
                        Next Fish
                        BYsb.AppendFormat("{0,11}", TempTotalMort(1).ToString("#######0.0"))
                        BYsb.AppendFormat("{0,11}", TempTotalMort(2).ToString("#######0.0"))
                        BYsb.AppendFormat("{0,11}", BYEscape(BY, Stk, Age, TStep).ToString("#######0.0"))
                        BYsb.Append(vbCrLf)
                    Next TStep
SkipFramAge2:
                Next Age
            Next Stk
        Next BY

SkipReverse:
        Dim testfile As String
      'testfile = "C:\data\FRAM\SizeLimits\Testfile.txt"
      'FileOpen(77, testfile, OpenMode.Output)

        BY = 2
        Fish = 54
        TStep = 3
      'Print(77, "Tstep" & "," & "Stk" & "," & "Fish" & "," & "Age" & "," & "Landed" & "," _
      '      & "Shakers" & "," & "DO" & "," & "MSFLanded" & "," & "MSFNR" & "," & "MSFShaker" _
      '      & "," & "MSFDO" & "," & "NR" & "," & "RunID" & vbCrLf)

        'For TStep = 1 To 3
        For Stk = 1 To NumStk
            'For Fish = 1 To NumFish
            For Age = 2 To MaxAge
            'Print(77, TStep & "," & Stk & "," & Fish & "," & Age & "," & BYLandedCatch(BY, Stk, Age, Fish, TStep) & "," _
            '    & BYShakers(BY, Stk, Age, Fish, TStep) & "," & BYDropOff(BY, Stk, Age, Fish, TStep) & "," & BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) & "," _
            '    & BYMSFNonRetention(BY, Stk, Age, Fish, TStep) & "," & BYMSFShakers(BY, Stk, Age, Fish, TStep) _
            '    & "," & BYMSFDropOff(BY, Stk, Age, Fish, TStep) & "," & BYNonRetention(BY, Stk, Age, Fish, TStep) & "," & RunIDNameSelect & vbCrLf)
            Next
            'Next
        Next
        'Next

        FileClose(77)

        '- Calculate the BY AEQ-ER Values for All Stocks and Brood Years
        Dim BYAEQCatch As Double
        Dim BYAEQEscape As Double
        Dim TimeStepCat(3) As Double
        For Stk = 1 To NumStk
            BYsb.Append(vbCrLf)
            BYsb.Append(StockName(Stk).ToString)
            BYsb.Append(vbCrLf)
            '- Fishing Year AEQ-ER
            BYAEQCatch = 0
            BYAEQEscape = 0
            For Age = MinAge To 5
                For TStep = 1 To NumSteps - 1
                    For Fish = 1 To NumFish
                        If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                            BYAEQCatch = BYAEQCatch + ((LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
                        Else
                            BYAEQCatch = BYAEQCatch + ((LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)))
                        End If
                    Next Fish
                    BYAEQEscape = BYAEQEscape + Escape(Stk, Age, TStep)
                Next TStep
            Next Age
            BYsb.Append("  FishYr  ")
            If BYAEQEscape = 0 Then
                BYsb.Append("0.0")
                BYsb.Append(vbCrLf)
            Else
                BYsb.AppendFormat("{0,13}", (BYAEQCatch / (BYAEQCatch + BYAEQEscape)).ToString("0.000000"))
                BYsb.AppendFormat("{0,13}", BYAEQCatch.ToString("######0.0"))
                BYsb.AppendFormat("{0,13}", BYAEQEscape.ToString("######0.0"))
                BYsb.Append(vbCrLf)
            End If
            '- Brood Year AEQ ER
            For BY = 2 To 5
                BYsb.AppendFormat("  BY-Age{0,2:N}", BY.ToString("00"))
                BYAEQCatch = 0
                BYAEQEscape = 0
                ReDim TimeStepCat(3)
                For Age = MinAge To 5
                    For TStep = 1 To NumSteps - 1
                        For Fish = 1 To NumFish
                            If TerminalFisheryFlag(Fish, TStep) = Term Then
                                BYAEQCatch = BYAEQCatch + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                                TimeStepCat(TStep) = TimeStepCat(TStep) + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                            Else
                                BYAEQCatch = BYAEQCatch + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
                                TimeStepCat(TStep) = TimeStepCat(TStep) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
                            End If
                        Next Fish
                        BYAEQEscape = BYAEQEscape + BYEscape(BY, Stk, Age, TStep)
                    Next TStep
                Next Age
                If BYAEQEscape = 0 Then
                    BYsb.Append(vbCrLf)
                Else
                    If OptionChinookBYAEQ = 2 Then
                        For TStep = 1 To 3
                            BYsb.AppendFormat("{0,11}  ", TimeStepCat(TStep).ToString("######0.0"))
                        Next TStep
                    End If
                    BYsb.AppendFormat("{0,13}", (BYAEQCatch / (BYAEQCatch + BYAEQEscape)).ToString("0.000000"))
                    BYsb.AppendFormat("{0,13}", BYAEQCatch.ToString("######0.0"))
                    BYsb.AppendFormat("{0,13}", BYAEQEscape.ToString("######0.0"))
                    BYsb.Append(vbCrLf)
                End If
                If OptionChinookBYAEQ = 2 Then Exit For
            Next BY
        Next Stk

        If OptionChinookBYAEQ = 1 Then
            If NumStk > 50 Then
                Call BYCHINSFTran()
            Else
                Call BYCHINTran()
            End If
        End If

        'If OptionChinookBYAEQ = 2 Then '- test oregon tule
        '   BY = 2
        '   For Stk = 37 To 38
        '      BYsb.Append(vbCrLf)
        '      BYsb.Append(StockName(Stk).ToString)
        '      BYsb.Append(vbCrLf)
        '      For Fish = 1 To NumFish
        '         BYsb.AppendFormat("{0,25}", (Left(FisheryName$(Fish), 25)).ToString)
        '         ReDim TimeStepCat(3)
        '         For TStep = 1 To NumSteps - 1
        '            For Age = MinAge To 5
        '               If TerminalFisheryFlag(Fish, TStep) = Term Then
        '                  TimeStepCat(TStep) = TimeStepCat(TStep) + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYLegalShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)
        '               Else
        '                  TimeStepCat(TStep) = TimeStepCat(TStep) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYLegalShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
        '               End If
        '            Next Age
        '            BYsb.AppendFormat("{0,14}  ", TimeStepCat(TStep).ToString("#######0"))
        '         Next TStep
        '         BYsb.Append(vbCrLf)
        '      Next Fish
        '      BYsb.Append(vbCrLf)
        '   Next Stk
        'End If

        BYsw.WriteLine(BYsb)
        'BYsw.Flush()
        BYsw.Close()
        BYsb = Nothing

    End Sub

    Sub BYCHINSFTran()
        Exit Sub
    End Sub
    Sub BYCHINSFTranOld()

        '--- TAMM Transfer File for Selective Fishery Style CHINOOK ---

        Dim BY
        Dim PSStocks, AllStocks, SomeStocks As Integer
        Dim USPS0u, UWACCu, DSPS0u, SPSYRu, NONSSu, Sps1011u, SSETACu As Double
        Dim USPS0m, UWACCm, DSPS0m, SPSYRm, NONSSm, Sps1011m, SSETACm As Double
        Dim USSETRSu, DSSETRSu, SUMETRSu As Double
        Dim USSETRSm, DSSETRSm, SUMETRSm As Double
        Dim StkNum, StkVal, I, J As Integer
        Dim TotalChinEsc(6, 23) As Double
        Dim SptSave(14) As Double
        Dim TermChinAbun(14) As Double

        BY = 2
        If NumStk = 38 Or NumStk = 76 Then
            PSStocks = 20
            AllStocks = 38
            SomeStocks = 23
        ElseIf NumStk = 33 Or NumStk = 66 Then
            PSStocks = 19
            AllStocks = 33
            SomeStocks = 22
        Else
            PSStocks = 20
            AllStocks = NumStk / 2
            SomeStocks = 23
        End If

        '------------------------ Terminal Area Escapements ---
        For TStep = 1 To NumSteps - 1
            For Stk = 1 To PSStocks
                If Stk > 2 And Stk < 19 Then
                    StkNum = Stk - 1
                    StkVal = Stk
                Else
                    If Stk = 19 Then
                        StkVal = 33 '- WhRvr Spr Year
                        StkNum = 22
                    ElseIf Stk = 20 Then
                        StkVal = 38 '- Hoko
                        StkNum = 23
                    Else
                        StkVal = Stk
                        StkNum = Stk
                    End If
                End If
                For Age = MinAge To MaxAge
                    If Age = 2 Then
                        TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                        TotalChinEsc(4, StkNum) = TotalChinEsc(4, StkNum) + BYEscape(BY, StkVal * 2, Age, TStep)
                    Else
                        TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                        TotalChinEsc(5, StkNum) = TotalChinEsc(5, StkNum) + BYEscape(BY, StkVal * 2, Age, TStep)
                    End If
                    If StkNum = 13 Then  '--- sps yearling split
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + (SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(4, 20) = TotalChinEsc(4, 20) + (SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(4, 21) = TotalChinEsc(4, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2 - 1, Age, TStep)
                            TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + (SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                            TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, StkVal * 2, Age, TStep)
                        End If
                    End If
                    If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + BYEscape(BY, Stk * 2 - 1, Age, TStep)
                            TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + BYEscape(BY, Stk * 2, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + BYEscape(BY, Stk * 2 - 1, Age, TStep)
                            TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + BYEscape(BY, Stk * 2, Age, TStep)
                        End If
                    End If
                    If StkNum = 12 Then               '--- Deep SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + BYEscape(BY, Stk * 2 - 1, Age, TStep)
                            TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + BYEscape(BY, Stk * 2, Age, TStep)
                        Else
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + BYEscape(BY, Stk * 2 - 1, Age, TStep)
                            TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + BYEscape(BY, Stk * 2, Age, TStep)
                        End If
                    End If
                Next Age
            Next Stk
        Next TStep

        'ADD IN FRESHWATER NET and Sport CATCH TO GET EXTREME TERMINAL RUN

        For Fish = 72 To 73
            For Stk = 1 To PSStocks
                If Stk > 2 And Stk < 19 Then
                    StkNum = Stk - 1
                    StkVal = Stk
                Else
                    If Stk = 19 Then
                        StkVal = 33 '- WhRvr Spr Year
                        StkNum = 22
                    ElseIf Stk = 20 Then
                        StkVal = 38 '- Hoko
                        StkNum = 23
                    Else
                        StkVal = Stk
                        StkNum = Stk
                    End If
                End If
                For TStep = 1 To NumSteps - 1
                    For Age = MinAge To MaxAge
                        If Age = 2 Then
                            TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(4, StkNum) = TotalChinEsc(4, StkNum) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                        Else
                            TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(5, StkNum) = TotalChinEsc(5, StkNum) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                        End If
                        If StkNum = 13 Then  '---- sps yearling split
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(4, 20) = TotalChinEsc(4, 20) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(4, 21) = TotalChinEsc(4, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                                TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            End If
                        End If
                        If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(4, 18) = TotalChinEsc(4, 18) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            End If
                        End If
                        If StkNum = 12 Then               '--- Deep SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(4, 19) = TotalChinEsc(4, 19) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + BYLandedCatch(BY, StkVal * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + BYLandedCatch(BY, StkVal * 2, Age, Fish, TStep)
                            End If
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- New Section for TAA for 7B, 8, 8A, 10, and 12 plus TRS
        For Stk = 1 To SomeStocks
            TotalChinEsc(3, Stk) = TotalChinEsc(2, Stk) '--- Start with Age 3 ETRS local stock
            TotalChinEsc(6, Stk) = TotalChinEsc(5, Stk) '--- Start with Age 3 ETRS local stock
        Next Stk

        '-------------------------------- NkSam TAA
        TermChinAbun(1) = TotalChinEsc(3, 1)
        TermChinAbun(8) = TotalChinEsc(6, 1)
        TStep = 3 '- Only Time 3 by definition
        For Fish = 39 To 40  '---- B'Ham Bay Net 7B
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk = 1 Then
                        TotalChinEsc(3, 1) = TotalChinEsc(3, 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TotalChinEsc(6, 1) = TotalChinEsc(6, 1) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        '--- TRS
                    End If
                    If Stk = 2 Or Stk = 3 Then
                        TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TotalChinEsc(6, 2) = TotalChinEsc(6, 2) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        '--- TRS
                    End If
                    TermChinAbun(1) = TermChinAbun(1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                    TermChinAbun(8) = TermChinAbun(8) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                    '---------- TAA
                Next Age
            Next Stk
        Next Fish

        TStep = 2
        '--- B'Ham Bay Net 7B Nooksack Spring Chinook time step 2
        For Stk = 2 To 3
            For Fish = 39 To 40
                For Age = MinAge To MaxAge   '---- All Ages in ETRS marine catches
                    TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                    TotalChinEsc(6, 2) = TotalChinEsc(6, 2) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                Next Age
            Next Fish
        Next Stk
        '-------------------------------- Skagit TAA
        TermChinAbun(2) = TotalChinEsc(3, 3) + TotalChinEsc(3, 4)
        TermChinAbun(9) = TotalChinEsc(6, 3) + TotalChinEsc(6, 4)
        For Fish = 46 To 47  '--- Skagit Bay Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For TStep = 2 To 3          '---- only Step 2 and 3 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk = 4 Or Stk = 5 Or Stk = 6 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            '--- TRS Falls and Springs
                        End If
                        TermChinAbun(2) = TermChinAbun(2) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TermChinAbun(9) = TermChinAbun(9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        '---------- TAA
                    Next Age
                Next TStep
            Next Stk
        Next Fish
        '------------------------------------ Still/Snohomish 8A TAA
        TermChinAbun(3) = TotalChinEsc(3, 6) + TotalChinEsc(3, 7) + TotalChinEsc(3, 8) + TotalChinEsc(3, 9)
        TermChinAbun(10) = TotalChinEsc(6, 6) + TotalChinEsc(6, 7) + TotalChinEsc(6, 8) + TotalChinEsc(6, 9)
        TStep = 3 '- by definition
        For Fish = 49 To 50   '---- Area 8A Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk >= 7 And Stk <= 9 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        '--- Tulalip uses ETRS ... Don't add 8A Catch for Stock #10  3/24/99
                    End If
                    TermChinAbun(3) = TermChinAbun(3) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                    TermChinAbun(10) = TermChinAbun(10) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                Next Age
            Next Stk
        Next Fish
        '--- Tulalip Bay Net

        For Stk = 1 To AllStocks '- NumStk / 2
            For TStep = 3 To 3 '---- only Step 3 and 4 by Definition
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk >= 7 And Stk <= 9 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                    End If
                    TermChinAbun(3) = TermChinAbun(3) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                    TermChinAbun(10) = TermChinAbun(10) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                    If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                        TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(4, 9) = TotalChinEsc(4, 9) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                        TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                        TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                    Else
                        TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, 48, TStep)
                        TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                        TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + BYLandedCatch(BY, Stk * 2, Age, 48, TStep)
                    End If
                Next Age
            Next TStep
        Next Stk

        For Fish = 51 To 52
            For Stk = 1 To AllStocks '- NumStk / 2
                For TStep = 3 To 3 '---- only Step 3 and 4 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk >= 7 And Stk <= 9 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        End If
                        TermChinAbun(3) = TermChinAbun(3) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TermChinAbun(10) = TermChinAbun(10) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                            TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(4, 9) = TotalChinEsc(4, 9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        Else
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                            TotalChinEsc(5, 9) = TotalChinEsc(5, 9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            TotalChinEsc(6, 9) = TotalChinEsc(6, 9) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish
        '-------------------------------------------- South Sound TAA
        TermChinAbun(4) = TotalChinEsc(3, 10) + TotalChinEsc(3, 11) + TotalChinEsc(3, 12) + TotalChinEsc(3, 13) + TotalChinEsc(3, 22)
        TermChinAbun(11) = TotalChinEsc(6, 10) + TotalChinEsc(6, 11) + TotalChinEsc(6, 12) + TotalChinEsc(6, 13) + TotalChinEsc(6, 22)
        Sps1011u = 0.0
        USPS0u = 0.0
        UWACCu = 0.0
        DSPS0u = 0.0
        SPSYRu = 0.0
        NONSSu = 0.0
        Sps1011m = 0.0
        USPS0m = 0.0
        UWACCm = 0.0
        DSPS0m = 0.0
        SPSYRm = 0.0
        NONSSm = 0.0
        TStep = 3 '- by definition
        For Fish = 58 To 71
            If Fish > 63 And Fish < 68 Then GoTo NotFish
            For Stk = 1 To AllStocks '- NumStk / 2
                '      For TStep = 1 To NumSteps
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS catches
                    TermChinAbun(4) = TermChinAbun(4) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                    TermChinAbun(11) = TermChinAbun(11) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                    '---------- TAA
                    '-- 10/11 and 13 catches TRS for FRAM SS Stocks
                    '   both Falls and 13A Springs
                    If ((Stk > 10 And Stk < 16) Or Stk = 33) And (Fish <= 59 Or Fish = 68 Or Fish = 69) Then
                        Select Case Stk
                            Case 11
                                TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 10) = TotalChinEsc(6, 10) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 12
                                TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 11) = TotalChinEsc(6, 11) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 13
                                TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 12) = TotalChinEsc(6, 12) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 14
                                TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 13) = TotalChinEsc(6, 13) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 15  '--- WhRvr Spring Fing
                                TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 14) = TotalChinEsc(6, 14) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 33  '--- WhRvr Spring Year
                                TotalChinEsc(3, 22) = TotalChinEsc(3, 22) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                TotalChinEsc(6, 22) = TotalChinEsc(6, 22) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                                '               Case 38  '--- Hoko
                                '                  TotalChinEsc(3, 23) = TotalChinEsc(3, 23) + BYLandedCatch(BY,Stk * 2 - 1, age, Fish, TStep)
                                '                  TotalChinEsc(6, 23) = TotalChinEsc(6, 23) + BYLandedCatch(BY,Stk * 2, age, Fish, TStep)
                        End Select
                    End If
                    If Fish <= 59 Then       '--- 10/11 Net Catches for Split TAA
                        Sps1011u = Sps1011u + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        Sps1011m = Sps1011m + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                    End If
                    '------ 10A,10E,13A,SPS Net ETRS Catches
                    If (Fish >= 60 And Fish <= 63) Or (Fish >= 68 And Fish <= 71) Then
                        Select Case Stk
                            Case 11
                                USPS0u = USPS0u + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                USPS0m = USPS0m + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 12
                                UWACCu = UWACCu + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                UWACCm = UWACCm + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 13
                                DSPS0u = DSPS0u + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                DSPS0m = DSPS0m + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                            Case 14
                                SPSYRu = SPSYRu + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                SPSYRm = SPSYRm + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                                '                  Case 33
                                '                     SPSYRu = SPSYRu + BYLandedCatch(BY,Stk * 2 - 1, age, Fish, TStep)
                                '                     SPSYRm = SPSYRm + BYLandedCatch(BY,Stk * 2, age, Fish, TStep)
                            Case Else
                                NONSSu = NONSSu + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                                NONSSm = NONSSm + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        End Select
                    End If
                Next Age
                '      NEXT TStep
            Next Stk
NotFish:
        Next Fish

        SSETACu = USPS0u + UWACCu + DSPS0u + SPSYRu
        If SSETACu <> 0.0 Then
            TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0u + (NONSSu * (USPS0u / SSETACu))
            TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACCu + (NONSSu * (UWACCu / SSETACu))
            TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0u + (NONSSu * (DSPS0u / SSETACu))
            TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYRu + (NONSSu * (SPSYRu / SSETACu))
            TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0u + (NONSSu * (USPS0u / SSETACu))
            TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACCu + (NONSSu * (UWACCu / SSETACu))
            TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0u + (NONSSu * (DSPS0u / SSETACu))
            TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYRu + (NONSSu * (SPSYRu / SSETACu))

            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0u + (NONSSu * (USPS0u / SSETACu)) + _
               UWACCu + (NONSSu * (UWACCu / SSETACu)) + _
               (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0u + (NONSSu * (DSPS0u / SSETACu)) + _
               (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0u + (NONSSu * (USPS0u / SSETACu)) + _
               UWACCu + (NONSSu * (UWACCu / SSETACu)) + _
               (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0u + (NONSSu * (DSPS0u / SSETACu)) + _
               (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * SpsYrSpl
            TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + (SPSYRu + (NONSSu * (SPSYRu / SSETACu))) * (1.0 - SpsYrSpl)
        End If

        SSETACm = USPS0m + UWACCm + DSPS0m + SPSYRm
        If SSETACm <> 0.0 Then
            TotalChinEsc(5, 10) = TotalChinEsc(5, 10) + USPS0m + (NONSSm * (USPS0m / SSETACm))
            TotalChinEsc(5, 11) = TotalChinEsc(5, 11) + UWACCm + (NONSSm * (UWACCm / SSETACm))
            TotalChinEsc(5, 12) = TotalChinEsc(5, 12) + DSPS0m + (NONSSm * (DSPS0m / SSETACm))
            TotalChinEsc(5, 13) = TotalChinEsc(5, 13) + SPSYRm + (NONSSm * (SPSYRm / SSETACm))
            TotalChinEsc(6, 10) = TotalChinEsc(6, 10) + USPS0m + (NONSSm * (USPS0m / SSETACm))
            TotalChinEsc(6, 11) = TotalChinEsc(6, 11) + UWACCm + (NONSSm * (UWACCm / SSETACm))
            TotalChinEsc(6, 12) = TotalChinEsc(6, 12) + DSPS0m + (NONSSm * (DSPS0m / SSETACm))
            TotalChinEsc(6, 13) = TotalChinEsc(6, 13) + SPSYRm + (NONSSm * (SPSYRm / SSETACm))

            TotalChinEsc(5, 18) = TotalChinEsc(5, 18) + USPS0m + (NONSSm * (USPS0m / SSETACm)) + _
               UWACCm + (NONSSm * (UWACCm / SSETACm)) + _
               (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            TotalChinEsc(5, 19) = TotalChinEsc(5, 19) + DSPS0m + (NONSSm * (DSPS0m / SSETACm)) + _
               (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            TotalChinEsc(5, 20) = TotalChinEsc(5, 20) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            TotalChinEsc(5, 21) = TotalChinEsc(5, 21) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            TotalChinEsc(6, 18) = TotalChinEsc(6, 18) + USPS0m + (NONSSm * (USPS0m / SSETACm)) + _
               UWACCm + (NONSSm * (UWACCm / SSETACm)) + _
               (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            TotalChinEsc(6, 19) = TotalChinEsc(6, 19) + DSPS0m + (NONSSm * (DSPS0m / SSETACm)) + _
               (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
            TotalChinEsc(6, 20) = TotalChinEsc(6, 20) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * SpsYrSpl
            TotalChinEsc(6, 21) = TotalChinEsc(6, 21) + (SPSYRm + (NONSSm * (SPSYRm / SSETACm))) * (1.0 - SpsYrSpl)
        End If

        '------ Area 10/11 Net Catch Split between Upper and Deep South Sound TAA
        USSETRSu = TotalChinEsc(3, 18)
        DSSETRSu = TotalChinEsc(3, 19)
        SUMETRSu = USSETRSu + DSSETRSu
        If SUMETRSu <> 0.0 Then
            TermChinAbun(6) = USSETRSu + ((USSETRSu / SUMETRSu) * Sps1011u)
            TermChinAbun(7) = DSSETRSu + ((DSSETRSu / SUMETRSu) * Sps1011u)
        End If
        USSETRSm = TotalChinEsc(6, 18)
        DSSETRSm = TotalChinEsc(6, 19)
        SUMETRSm = USSETRSm + DSSETRSm
        If SUMETRSm <> 0.0 Then
            TermChinAbun(13) = USSETRSm + ((USSETRSm / SUMETRSm) * Sps1011m)
            TermChinAbun(14) = DSSETRSm + ((DSSETRSm / SUMETRSm) * Sps1011m)
        End If

        '- WhRvrSpr not used in 13A
        'Stk = 15
        'TStep = 2
        '--- SPS Yearling 13A Time Step 2 ETRS catches
        'For Fish = 70 To 71
        '   For age = Minage to Maxage   '---- All Ages in ETRS marine catches
        '      TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + BYLandedCatch(BY,Stk, age, Fish, TStep)
        '   Next age
        'Next Fish

        '-------------------------------------------- Hood Canal TAA
        TermChinAbun(5) = TotalChinEsc(3, 15) + TotalChinEsc(3, 16)
        TermChinAbun(12) = TotalChinEsc(6, 15) + TotalChinEsc(6, 16)
        TStep = 3 '- by definition
        For Fish = 65 To 66  '--- HC Net
            For Stk = 1 To AllStocks '- NumStk / 2
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS
                    If Stk >= 16 And Stk <= 17 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                        TotalChinEsc(6, Stk - 1) = TotalChinEsc(6, Stk - 1) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                        '--- TRS
                    End If
                    TermChinAbun(5) = TermChinAbun(5) + BYLandedCatch(BY, Stk * 2 - 1, Age, Fish, TStep)
                    TermChinAbun(12) = TermChinAbun(12) + BYLandedCatch(BY, Stk * 2, Age, Fish, TStep)
                    '---------- TAA
                Next Age
            Next Stk
        Next Fish

        '--- Subtract TAMM FW Sport and Add Marine Sport Savings to TAA's
        '--- Split proportion by Unmarked + Marked
        For I = 1 To 14
            SptSave(I) = 0
        Next I
        SptSave(1) = ((TNkFWSpt! - TNkMSA!) * (TermChinAbun(1) / (TermChinAbun(1) + TermChinAbun(8))))
        SptSave(2) = ((TSkFWSpt! - TSkMSA!) * (TermChinAbun(2) / (TermChinAbun(2) + TermChinAbun(9))))
        SptSave(3) = ((TSnFWSpt! - TSnMSA!) * (TermChinAbun(3) / (TermChinAbun(3) + TermChinAbun(10))))
        SptSave(4) = TermChinAbun(4) / (TermChinAbun(4) + TermChinAbun(11))
        SptSave(5) = (THCFWSpt! * (TermChinAbun(5) / (TermChinAbun(5) + TermChinAbun(12))))
        SptSave(6) = TermChinAbun(6) / (TermChinAbun(6) + TermChinAbun(13))
        SptSave(7) = TermChinAbun(7) / (TermChinAbun(7) + TermChinAbun(14))
        SptSave(8) = ((TNkFWSpt! - TNkMSA!) * (TermChinAbun(8) / (TermChinAbun(1) + TermChinAbun(8))))
        SptSave(9) = ((TSkFWSpt! - TSkMSA!) * (TermChinAbun(9) / (TermChinAbun(2) + TermChinAbun(9))))
        SptSave(10) = ((TSnFWSpt! - TSnMSA!) * (TermChinAbun(10) / (TermChinAbun(3) + TermChinAbun(10))))
        SptSave(11) = TermChinAbun(11) / (TermChinAbun(4) + TermChinAbun(11))
        SptSave(12) = (THCFWSpt! * (TermChinAbun(12) / (TermChinAbun(5) + TermChinAbun(12))))
        SptSave(13) = TermChinAbun(13) / (TermChinAbun(6) + TermChinAbun(13))
        SptSave(14) = TermChinAbun(14) / (TermChinAbun(7) + TermChinAbun(14))
        For I = 1 To 14
            TermChinAbun(I) = TermChinAbun(I) - SptSave(I)
        Next I

        '------- Print Version Number and Command File Number ---
        'Print #13, VersNumb$ & " -BY=2"
        'Print #13, Mid(CMDFile$, 1, 4)
        'Print #13,
        'Print #13,

        '----- Print Terminal and Extreme Terminal Run Sizes ---
        For Stk = 1 To 17
            'Print #13, Format(Str(CLng(TotalChinEsc(3, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, Stk))), " @@@@@@@");
            'Print #13, Format(Str(CLng(TotalChinEsc(6, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, Stk))), " @@@@@@@")
            Select Case Stk
                Case 1
                    'Print #13, Format(Str(CLng(TermChinAbun(1))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(8))), "                 @@@@@@@")
                Case 4
                    'Print #13, Format(Str(CLng(TermChinAbun(2))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(9))), "                 @@@@@@@")
                Case 9
                    'Print #13, Format(Str(CLng(TermChinAbun(3))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(10))), "                 @@@@@@@")
                Case 13
                    '--- Upper South Sound Yr.
                    'Print #13, Format(Str(CLng(TotalChinEsc(3, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 20))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TotalChinEsc(6, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 20))), " @@@@@@@")
                    '--- Deep South Sound Yr.
                    'Print #13, Format(Str(CLng(TotalChinEsc(3, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 21))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TotalChinEsc(6, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 21))), " @@@@@@@")
                    '--- Upper South Sound Agg.
                    'Print #13, Format(Str(CLng(TermChinAbun(6))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 18))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 18))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(13))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 18))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 18))), " @@@@@@@")
                    '--- Deep South Sound Agg.
                    'Print #13, Format(Str(CLng(TermChinAbun(7))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 19))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 19))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(14))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 19))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 19))), " @@@@@@@")
                    '--- Total TAA
                    'Print #13, Format(Str(CLng(TermChinAbun(4))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(11))), "                 @@@@@@@")
                Case 16
                    'Print #13, Format(Str(CLng(TermChinAbun(5))), " @@@@@@@");
                    'Print #13, Format(Str(CLng(TermChinAbun(12))), "                 @@@@@@@")
                Case Else
            End Select
        Next Stk
        '- Add White River Springs to Bottom of List (Sum Both Components)
        'Print #13, Format(Str(CLng(TotalChinEsc(3, 22))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 22))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 22))), " @@@@@@@");
        'Print #13, Format(Str(CLng(TotalChinEsc(6, 22))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 22))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 22))), " @@@@@@@")

        '- Add Hoko to Bottom of List (Sum Both Components)
        'Print #13, Format(Str(CLng(TotalChinEsc(3, 23))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 23))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 23))), " @@@@@@@");
        'Print #13, Format(Str(CLng(TotalChinEsc(6, 23))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(4, 23))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(5, 23))), " @@@@@@@")

        '---------------------------- GET TOTAL FISHERY CATCH DATA ---
        Dim MarkLCat(NumSteps) As Double '- Marked Landed Catch
        Dim MarkTMrt(NumSteps) As Double '- Marked Total Mortality
        Dim AddFish(4, NumSteps + 1) As Double
        Dim TotalCatch, TotalMort, MarkTotalCatch, MarkTotalMort As Double
        Dim LndCatch, LndMort As Double
        Dim TStepVal As Integer
        For I = 1 To 4
            For J = 1 To 5
                AddFish(I, J) = 0.0
            Next J
        Next I
        For Fish = 1 To NumFish
            TotalCatch = 0.0
            TotalMort = 0.0
            MarkTotalCatch = 0.0
            MarkTotalMort = 0.0
            ReDim MarkLCat(NumSteps)
            ReDim MarkTMrt(NumSteps)
            For TStep = 1 To NumSteps
                LndCatch = 0.0
                LndMort = 0.0
                TStepVal = 0
                Select Case TStep
                    Case 1
                        GoTo SkipTime1Print
                    Case 2, 3
                        TStepVal = TStep
                    Case 4
                        TStepVal = 1
                End Select
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        If Fish = 13 Or Fish = 14 Or (TammChinookRunFlag = 1 And Fish = 56) Then
                            '            If Fish = 13 Or Fish = 14 Then
                            If (Stk Mod 2) <> 0 Then '- UnMarked Catch & Mortality
                                AddFish(1, TStepVal) = AddFish(1, TStepVal) + ((BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)) / ModelStockProportion(Fish))
                                AddFish(2, TStepVal) = AddFish(2, TStepVal) + (BYLandedCatch(BY, Stk, Age, Fish, TStepVal) / ModelStockProportion(Fish))
                                AddFish(1, NumSteps + 1) = AddFish(1, NumSteps + 1) + ((BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)) / ModelStockProportion(Fish))
                                AddFish(2, NumSteps + 1) = AddFish(2, NumSteps + 1) + (BYLandedCatch(BY, Stk, Age, Fish, TStepVal) / ModelStockProportion(Fish))
                            Else
                                AddFish(3, TStepVal) = AddFish(3, TStepVal) + ((BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)) / ModelStockProportion(Fish))
                                AddFish(4, TStepVal) = AddFish(4, TStepVal) + (BYLandedCatch(BY, Stk, Age, Fish, TStepVal) / ModelStockProportion(Fish))
                                AddFish(3, NumSteps + 1) = AddFish(3, NumSteps + 1) + ((BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)) / ModelStockProportion(Fish))
                                AddFish(4, NumSteps + 1) = AddFish(4, NumSteps + 1) + (BYLandedCatch(BY, Stk, Age, Fish, TStepVal) / ModelStockProportion(Fish))
                            End If
                        ElseIf (Stk Mod 2) <> 0 Then '- UnMarked Catch & Mortality
                            LndMort = LndMort + BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)
                            TotalMort = TotalMort + BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)
                            LndCatch = LndCatch + BYLandedCatch(BY, Stk, Age, Fish, TStepVal)
                            TotalCatch = TotalCatch + BYLandedCatch(BY, Stk, Age, Fish, TStepVal)
                        Else                      '- Marked Catch & Mortality
                            MarkTMrt(TStepVal) = MarkTMrt(TStepVal) + BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)
                            MarkTotalMort = MarkTotalMort + BYLandedCatch(BY, Stk, Age, Fish, TStepVal) + BYNonRetention(BY, Stk, Age, Fish, TStepVal) + BYShakers(BY, Stk, Age, Fish, TStepVal) + BYDropOff(BY, Stk, Age, Fish, TStepVal)
                            MarkLCat(TStepVal) = MarkLCat(TStepVal) + BYLandedCatch(BY, Stk, Age, Fish, TStepVal)
                            MarkTotalCatch = MarkTotalCatch + BYLandedCatch(BY, Stk, Age, Fish, TStepVal)
                        End If
                    Next Age
                Next Stk
SkipTime1Print:
                '- Print UnMarked Catch & Mortality
                If Fish = 13 Or Fish = 14 Or (TammChinookRunFlag = 1 And Fish = 56) Then
                    GoTo NextTStep
                ElseIf Fish = 15 Or (TammChinookRunFlag = 1 And Fish = 57) Then '- Combined Fisheries
                    'Print #15, Format(Str(CLng((LndMort / ModelStockProportion(Fish) + AddFish(1, TStep)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((LndCatch / ModelStockProportion(Fish) + AddFish(2, TStep)))), " @@@@@@@");
                Else '- Normal Fishery Print UnMarked Fish
                    'Print #15, Format(Str(CLng((LndMort / ModelStockProportion(Fish)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((LndCatch / ModelStockProportion(Fish)))), " @@@@@@@");
                End If
NextTStep:
            Next TStep
            If Fish = 13 Or Fish = 14 Or (TammChinookRunFlag = 1 And Fish = 56) Then
                GoTo NextFish
            ElseIf Fish = 15 Or (TammChinookRunFlag = 1 And Fish = 57) Then '- Combined Fisheries
                'Print #15, Format(Str(CLng((TotalMort / ModelStockProportion(Fish) + AddFish(1, NumSteps + 1)))), " @@@@@@@");
                'Print #14, Format(Str(CLng((TotalCatch / ModelStockProportion(Fish) + AddFish(2, NumSteps + 1)))), " @@@@@@@");
            Else
                'Print #15, Format(Str(CLng((TotalMort / ModelStockProportion(Fish)))), " @@@@@@@");
                'Print #14, Format(Str(CLng((TotalCatch / ModelStockProportion(Fish)))), " @@@@@@@");
            End If
            '- Print Marked Catch and Mortality at End-of-Line
            For TStep = 1 To NumSteps
                Select Case TStep
                    Case 1
                        TStepVal = 4
                    Case 2, 3
                        TStepVal = TStep
                    Case 4
                        TStepVal = 1
                End Select
                If Fish = 15 Or (TammChinookRunFlag = 1 And Fish = 57) Then '- Combined Fisheries
                    'Print #15, Format(Str(CLng((MarkTMrt(TStepVal) / ModelStockProportion(Fish) + AddFish(3, TStepVal)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((MarkLCat(TStepVal) / ModelStockProportion(Fish) + AddFish(4, TStepVal)))), " @@@@@@@");
                Else
                    'Print #15, Format(Str(CLng((MarkTMrt(TStepVal) / ModelStockProportion(Fish)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((MarkLCat(TStepVal) / ModelStockProportion(Fish)))), " @@@@@@@");
                End If
            Next TStep
            If Fish = 15 Or (TammChinookRunFlag = 1 And Fish = 57) Then '- Combined Fisheries
                'Print #15, Format(Str(CLng((MarkTotalMort / ModelStockProportion(Fish) + AddFish(3, NumSteps + 1)))), " @@@@@@@")
                'Print #14, Format(Str(CLng((MarkTotalCatch / ModelStockProportion(Fish) + AddFish(4, NumSteps + 1)))), " @@@@@@@")
            Else
                'Print #15, Format(Str(CLng((MarkTotalMort / ModelStockProportion(Fish)))), " @@@@@@@")
                'Print #14, Format(Str(CLng((MarkTotalCatch / ModelStockProportion(Fish)))), " @@@@@@@")
            End If
            For I = 1 To 4
                For J = 1 To 5
                    AddFish(I, J) = 0.0
                Next J
            Next I
NextFish:
        Next Fish
        'Print #15,
        'Print #14,

        '-------------------------- STOCK CATCH BY FISHERY ---
        Dim FishNum, EndFish As Integer
        Dim PageVals(NumSteps + 1, NumFish - 2, 2) As Double
        Dim PageLCat(NumSteps + 1, NumFish - 2, 2) As Double

        '---- ReInitialize Page Matrix -----
        For Stk = 1 To PSStocks
            '- NF and SF Nooksack Spring Combined
            If Stk <> 3 Then
                'Print #15,
                'Print #14,
                For Fish = 1 To NumFish - 2
                    For TStep = 1 To NumSteps + 1
                        For I = 1 To 2 '- UnMarked and Marked Loop
                            PageVals(TStep, Fish, I) = 0
                            PageLCat(TStep, Fish, I) = 0
                        Next I
                    Next TStep
                Next Fish
            End If
            '- WhRvr Spring Yearling Added #33
            If Stk = 19 Then
                StkNum = 33
            ElseIf Stk = 20 Then
                StkNum = 38
            Else
                StkNum = Stk
            End If
            For Fish = 1 To NumFish
                '- Combined Fisheries
                If TammChinookRunFlag = 1 Then
                    '- Old Style Format (Area 10/11 Net Combined)
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 55
                            FishNum = Fish - 2
                        Case 56, 57
                            FishNum = 54
                        Case 58 To 73
                            FishNum = Fish - 3
                    End Select
                Else
                    '- Current Chinook TAMM Transfer
                    Select Case Fish
                        Case 1 To 12
                            FishNum = Fish
                        Case 13, 14, 15
                            FishNum = 13
                        Case 16 To 73
                            FishNum = Fish - 2
                    End Select
                End If
                For TStep = 1 To NumSteps - 1
                    For Age = MinAge To MaxAge
                        '- AEQ Value NOT used for Terminal Fisheries
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                            PageVals(TStep, FishNum, 1) = PageVals(TStep, FishNum, 1) + ((BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYShakers(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2 - 1, Age, Fish, TStep)))
                            PageVals(5, FishNum, 1) = PageVals(5, FishNum, 1) + ((BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYShakers(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2 - 1, Age, Fish, TStep)))
                            PageVals(TStep, FishNum, 2) = PageVals(TStep, FishNum, 2) + ((BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2, Age, Fish, TStep) + BYShakers(BY, StkNum * 2, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2, Age, Fish, TStep)))
                            PageVals(5, FishNum, 2) = PageVals(5, FishNum, 2) + ((BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2, Age, Fish, TStep) + BYShakers(BY, StkNum * 2, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2, Age, Fish, TStep)))
                        Else
                            PageVals(TStep, FishNum, 1) = PageVals(TStep, FishNum, 1) + ((BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYShakers(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2 - 1, Age, Fish, TStep)) * AEQ(StkNum * 2 - 1, Age, TStep))
                            PageVals(5, FishNum, 1) = PageVals(5, FishNum, 1) + ((BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYShakers(BY, StkNum * 2 - 1, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2 - 1, Age, Fish, TStep)) * AEQ(StkNum * 2 - 1, Age, TStep))
                            PageVals(TStep, FishNum, 2) = PageVals(TStep, FishNum, 2) + ((BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2, Age, Fish, TStep) + BYShakers(BY, StkNum * 2, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2, Age, Fish, TStep)) * AEQ(StkNum * 2, Age, TStep))
                            PageVals(5, FishNum, 2) = PageVals(5, FishNum, 2) + ((BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep) + BYNonRetention(BY, StkNum * 2, Age, Fish, TStep) + BYShakers(BY, StkNum * 2, Age, Fish, TStep) + BYDropOff(BY, StkNum * 2, Age, Fish, TStep)) * AEQ(StkNum * 2, Age, TStep))
                        End If
                        PageLCat(TStep, FishNum, 1) = PageLCat(TStep, FishNum, 1) + BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep)
                        PageLCat(5, FishNum, 1) = PageLCat(5, FishNum, 1) + BYLandedCatch(BY, StkNum * 2 - 1, Age, Fish, TStep)
                        PageLCat(TStep, FishNum, 2) = PageLCat(TStep, FishNum, 2) + BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep)
                        PageLCat(5, FishNum, 2) = PageLCat(5, FishNum, 2) + BYLandedCatch(BY, StkNum * 2, Age, Fish, TStep)
                    Next Age
                Next TStep
            Next Fish
            '--- Print Page Matrix ----
            If TammChinookRunFlag = 1 Then
                EndFish = 70
            Else
                EndFish = 71
            End If
            If Stk <> 2 Then
                For Fish = 1 To EndFish
                    For TStep = 1 To NumSteps + 1
                        Select Case TStep
                            Case 1
                                'Print #15, Format(Str(CLng(PageVals(4, Fish, 1))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(4, Fish, 1))), " @@@@@@@");
                            Case 2, 3, 5
                                'Print #15, Format(Str(CLng(PageVals(TStep, Fish, 1))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(TStep, Fish, 1))), " @@@@@@@");
                            Case 4
                                'Print #15, Format(Str(CLng(PageVals(1, Fish, 1))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(1, Fish, 1))), " @@@@@@@");
                        End Select
                    Next TStep
                    For TStep = 1 To NumSteps + 1
                        Select Case TStep
                            Case 1
                                'Print #15, Format(Str(CLng(PageVals(4, Fish, 2))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(4, Fish, 2))), " @@@@@@@");
                            Case 2, 3, 5
                                'Print #15, Format(Str(CLng(PageVals(TStep, Fish, 2))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(TStep, Fish, 2))), " @@@@@@@");
                            Case 4
                                'Print #15, Format(Str(CLng(PageVals(1, Fish, 2))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(1, Fish, 2))), " @@@@@@@");
                        End Select
                    Next TStep

                    'Print #15, Format(Str(CLng(PageVals(NumSteps + 1, Fish, 1) + PageVals(NumSteps + 1, Fish, 2))), " = @@@@@@@");
                    'Print #14, Format(Str(CLng(PageLCat(NumSteps + 1, Fish, 1) + PageLCat(NumSteps + 1, Fish, 2))), " = @@@@@@@");

                    'Print #15,
                    'Print #14,
                Next Fish
                'Print #15,
                'Print #14,
            End If
        Next Stk
        'Close #15
        'Close #14
        'Close #13
        'Close #2

    End Sub

    Sub BYCHINTran()
        Exit Sub
    End Sub

    Sub BYCHINTranOld()
        '------------------ TAMM Transfer File for CHINOOK ---
        Dim BY
        Dim StkNum, I, J As Integer
        Dim USPS0, UWACC, DSPS0, SPSYR, NONSS, Sps1011, SSETAC As Double
        Dim USSETRS, DSSETRS, SUMETRS As Double
        Dim TotalChinEsc(3, 22) As Double
        Dim TermChinAbun(7)

        BY = 2

        '      TammXfr$ = CMDDirect$ + "\" + "BKTT" + Left$(UCase$(CMDFile$), 4) + ".TAM"
        'Open TammXfr$ For Output As #15
        '      TammXfr$ = CMDDirect$ + "\" + "BKTL" + Left$(UCase$(CMDFile$), 4) + ".TAM"
        'Open TammXfr$ For Output As #14
        '      TammXfr$ = CMDDirect$ + "\" + "BKTX" + Left$(UCase$(CMDFile$), 4) + ".TAM"
        'Open TammXfr$ For Output As #13

        '------------------------ Terminal Area Escapements ---

        For TStep = 1 To NumSteps - 1
            For Stk = 1 To 18
                If Stk > 2 Then
                    StkNum = Stk - 1
                Else
                    StkNum = Stk
                End If
                For Age = MinAge To MaxAge
                    If Age = 2 Then
                        TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + BYEscape(BY, Stk, Age, TStep)
                    Else
                        TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + BYEscape(BY, Stk, Age, TStep)
                    End If
                    If StkNum = 13 Then  '--- sps yearling split
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * BYEscape(BY, Stk, Age, TStep)
                        End If
                    End If
                    If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + BYEscape(BY, Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + BYEscape(BY, Stk, Age, TStep)
                        End If
                    End If
                    If StkNum = 12 Then               '--- Deep SPS
                        If Age = 2 Then
                            TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + BYEscape(BY, Stk, Age, TStep)
                        Else
                            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + BYEscape(BY, Stk, Age, TStep)
                        End If
                    End If
                Next Age
            Next Stk
        Next TStep

        '
        'Print #13, "TotCkEsc Values 1-19 After Esc"
        'For I = 1 To 19
        '   For J = 1 To 3
        '      'Print #13, Format(CStr(CLng(TotCkEsc(J, I))), " @@@@@@");
        '   Next J
        '   'Print #13,
        'Next I
        '
        'ADD IN FRESHWATER NET and Sport CATCH TO GET EXTREME TERMINAL RUN

        'NumPerStep& = NumFish * NumStk * (Maxage - 1)
        For Fish = 72 To 73
            For Stk = 1 To 18
                If Stk > 2 Then
                    StkNum = Stk - 1
                Else
                    StkNum = Stk
                End If
                For TStep = 1 To NumSteps - 1
                    For Age = MinAge To MaxAge
                        If Age = 2 Then
                            TotalChinEsc(1, StkNum) = TotalChinEsc(1, StkNum) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        Else
                            TotalChinEsc(2, StkNum) = TotalChinEsc(2, StkNum) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        End If
                        If StkNum = 13 Then  '---- sps yearling split
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + (SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(1, 20) = TotalChinEsc(1, 20) + (SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(1, 21) = TotalChinEsc(1, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + (SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (1.0 - SpsYrSpl) * BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            End If
                        End If
                        If StkNum = 10 Or StkNum = 11 Then  '--- Upper SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 18) = TotalChinEsc(1, 18) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            End If
                        End If
                        If StkNum = 12 Then               '--- Deep SPS
                            If Age = 2 Then
                                TotalChinEsc(1, 19) = TotalChinEsc(1, 19) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Else
                                TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            End If
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish

        '--- New Section for TAA for 7B, 8, 8A, 10, and 12 plus TRS
        For Stk = 1 To 21
            TotalChinEsc(3, Stk) = TotalChinEsc(2, Stk) '--- Start with Age 3 ETRS local stock
        Next Stk

        '-------------------------------- NkSam TAA
        TermChinAbun(1) = TotalChinEsc(3, 1)
        TStep = 3 '- Only Time 3 by definition
        For Fish = 39 To 40  '---- B'Ham Bay Net 7B
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk = 1 Then
                        TotalChinEsc(3, 1) = TotalChinEsc(3, 1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        '--- TRS
                    End If
                    If Stk = 2 Or Stk = 3 Then
                        TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        '--- TRS
                    End If
                    TermChinAbun(1) = TermChinAbun(1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    '---------- TAA
                Next Age
            Next Stk
        Next Fish

        TStep = 2
        '--- B'Ham Bay Net 7B Nooksack Spring Chinook time step 2
        For Stk = 2 To 3
            For Fish = 39 To 40
                For Age = MinAge To MaxAge   '---- All Ages in ETRS marine catches
                    TotalChinEsc(3, 2) = TotalChinEsc(3, 2) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                Next Age
            Next Fish
        Next Stk
        '-------------------------------- Skagit TAA
        TermChinAbun(2) = TotalChinEsc(3, 3) + TotalChinEsc(3, 4)
        For Fish = 46 To 47  '--- Skagit Bay Net
            For Stk = 1 To 32
                For TStep = 2 To 3          '---- only Step 2 and 3 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk = 4 Or Stk = 5 Or Stk = 6 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            '--- TRS Falls and Springs
                        End If
                        TermChinAbun(2) = TermChinAbun(2) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        '---------- TAA
                    Next Age
                Next TStep
            Next Stk
        Next Fish
        '------------------------------------ Still/Snohomish 8A TAA
        TermChinAbun(3) = TotalChinEsc(3, 6) + TotalChinEsc(3, 7) + TotalChinEsc(3, 8) + TotalChinEsc(3, 9)
        TStep = 3 '- by definition
        For Fish = 49 To 50   '---- Area 8A Net
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                    If Stk >= 7 And Stk <= 9 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        '--- Tulalip uses ETRS ... Don't add 8A Catch for Stock #10  3/24/99
                    End If
                    TermChinAbun(3) = TermChinAbun(3) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                Next Age
            Next Stk
        Next Fish
        For Fish = 51 To 52     '--- Tulalip Bay Net
            For Stk = 1 To 32
                For TStep = 3 To 3 '---- only Step 3 and 4 by Definition
                    For Age = MinAge To MaxAge   '---- All ages in TAA or TRS catches
                        If Stk >= 7 And Stk <= 9 Then
                            TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        End If
                        TermChinAbun(3) = TermChinAbun(3) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        If Age = 2 Then '--- Tulalip ETRS Includes 8D Catches ... Oddity
                            TotalChinEsc(1, 9) = TotalChinEsc(1, 9) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        Else
                            TotalChinEsc(2, 9) = TotalChinEsc(2, 9) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            TotalChinEsc(3, 9) = TotalChinEsc(3, 9) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        End If
                    Next Age
                Next TStep
            Next Stk
        Next Fish
        '-------------------------------------------- South Sound TAA
        TermChinAbun(4) = TotalChinEsc(3, 10) + TotalChinEsc(3, 11) + TotalChinEsc(3, 12) + TotalChinEsc(3, 13)
        TStep = 3 '- by definition
        Sps1011 = 0.0
        USPS0 = 0.0
        UWACC = 0.0
        DSPS0 = 0.0
        SPSYR = 0.0
        NONSS = 0.0
        For Fish = 58 To 71
            If Fish > 63 And Fish < 68 Then GoTo NotFish
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS catches
                    TermChinAbun(4) = TermChinAbun(4) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    '---------- TAA
                    '-- 10/11 and 13 catches TRS for FRAM SS Stocks
                    '   both Falls and 13A Springs
                    If ((Stk > 10 And Stk < 16) Or Stk = 33) And (Fish <= 59 Or Fish = 68 Or Fish = 69) Then
                        Select Case Stk
                            Case 11
                                TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 12
                                TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 13
                                TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 14
                                TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 15  '--- WhRvrSpr Fing
                                TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                                '               Case 33  '--- WhRvrSpr Fing
                                '                  TotalChinEsc(3, 22) = TotalChinEsc(3, 22) + BYLandedCatch(BY,Stk, age, Fish, TStep)
                        End Select
                    End If
                    If Fish <= 59 Then       '--- 10/11 Net Catches for Split TAA
                        Sps1011 = Sps1011 + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    End If
                    '------ 10A,10E,13A,SPS Net ETRS Catches
                    If (Fish >= 60 And Fish <= 63) Or (Fish >= 65 And Fish <= 68) Then
                        Select Case Stk
                            Case 11
                                USPS0 = USPS0 + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 12
                                UWACC = UWACC + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 13
                                DSPS0 = DSPS0 + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case 14
                                SPSYR = SPSYR + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                            Case Else
                                NONSS = NONSS + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        End Select
                    End If
                    '            If Stk = 15 And Fish >= 70 Then   '--- Spring Yearling
                    '               TotalChinEsc(2, 14) = TotalChinEsc(2, 14) + BYLandedCatch(BY,Stk, age, Fish, TStep) '--- ETRS
                    '               TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + BYLandedCatch(BY,Stk, age, Fish, TStep) '--- ETRS
                    '            End If
                Next Age
                '      NEXT TStep
            Next Stk
NotFish:
        Next Fish

        SSETAC = USPS0 + UWACC + DSPS0 + SPSYR
        If SSETAC <> 0.0 Then
            TotalChinEsc(2, 10) = TotalChinEsc(2, 10) + USPS0 + (NONSS * (USPS0 / SSETAC))
            TotalChinEsc(2, 11) = TotalChinEsc(2, 11) + UWACC + (NONSS * (UWACC / SSETAC))
            TotalChinEsc(2, 12) = TotalChinEsc(2, 12) + DSPS0 + (NONSS * (DSPS0 / SSETAC))
            TotalChinEsc(2, 13) = TotalChinEsc(2, 13) + SPSYR + (NONSS * (SPSYR / SSETAC))
            TotalChinEsc(3, 10) = TotalChinEsc(3, 10) + USPS0 + (NONSS * (USPS0 / SSETAC))
            TotalChinEsc(3, 11) = TotalChinEsc(3, 11) + UWACC + (NONSS * (UWACC / SSETAC))
            TotalChinEsc(3, 12) = TotalChinEsc(3, 12) + DSPS0 + (NONSS * (DSPS0 / SSETAC))
            TotalChinEsc(3, 13) = TotalChinEsc(3, 13) + SPSYR + (NONSS * (SPSYR / SSETAC))

            TotalChinEsc(2, 18) = TotalChinEsc(2, 18) + USPS0 + (NONSS * (USPS0 / SSETAC)) + _
               UWACC + (NONSS * (UWACC / SSETAC)) + _
               (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            TotalChinEsc(2, 19) = TotalChinEsc(2, 19) + DSPS0 + (NONSS * (DSPS0 / SSETAC)) + _
               (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            TotalChinEsc(2, 20) = TotalChinEsc(2, 20) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            TotalChinEsc(2, 21) = TotalChinEsc(2, 21) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 18) = TotalChinEsc(3, 18) + USPS0 + (NONSS * (USPS0 / SSETAC)) + _
               UWACC + (NONSS * (UWACC / SSETAC)) + _
               (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            TotalChinEsc(3, 19) = TotalChinEsc(3, 19) + DSPS0 + (NONSS * (DSPS0 / SSETAC)) + _
               (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
            TotalChinEsc(3, 20) = TotalChinEsc(3, 20) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * SpsYrSpl
            TotalChinEsc(3, 21) = TotalChinEsc(3, 21) + (SPSYR + (NONSS * (SPSYR / SSETAC))) * (1.0 - SpsYrSpl)
        End If

        '------ Area 10/11 Net LandedCatch Split between Upper and Deep South Sound TAA
        USSETRS = TotalChinEsc(3, 18)
        DSSETRS = TotalChinEsc(3, 19)
        SUMETRS = USSETRS + DSSETRS
        If SUMETRS <> 0.0 Then
            TermChinAbun(6) = USSETRS + ((USSETRS / SUMETRS) * Sps1011)
            TermChinAbun(7) = DSSETRS + ((DSSETRS / SUMETRS) * Sps1011)
        End If

        '----------- 13A Not Used for WhRvrSpr
        'Stk = 15
        'TStep = 2
        ''--- SPS Yearling 13A Time Step 2 ETRS catches
        'For Fish = 70 To 71
        '   For age = Minage to Maxage   '---- All Ages in ETRS marine catches
        '      TotalChinEsc(3, 14) = TotalChinEsc(3, 14) + BYLandedCatch(BY,Stk, age, Fish, TStep)
        '   Next age
        'Next Fish

        '-------------------------------------------- Hood Canal TAA
        TermChinAbun(5) = TotalChinEsc(3, 15) + TotalChinEsc(3, 16)
        TStep = 3 '- by definition
        For Fish = 65 To 66  '--- HC Net
            For Stk = 1 To 32
                For Age = MinAge To MaxAge   '---- All ages in TAA and TRS
                    If Stk >= 16 And Stk <= 17 Then
                        TotalChinEsc(3, Stk - 1) = TotalChinEsc(3, Stk - 1) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        '--- TRS
                    End If
                    TermChinAbun(5) = TermChinAbun(5) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    '---------- TAA
                Next Age
            Next Stk
        Next Fish

        '--- Subtract TAMM FW Sport and Add Marine Sport Savings to TAA's
        TermChinAbun(1) = TermChinAbun(1) - TNkFWSpt! + TNkMSA!
        TermChinAbun(2) = TermChinAbun(2) - TSkFWSpt! + TSkMSA!
        TermChinAbun(3) = TermChinAbun(3) - TSnFWSpt! + TSnMSA!
        'TermChinAbun(4) = TermChinAbun(4)
        TermChinAbun(5) = TermChinAbun(5) - THCFWSpt!
        'TermChinAbun(6) = TermChinAbun(6)
        'TermChinAbun(7) = TermChinAbun(7)

        '
        'Print #13, "TotCkEsc Values 1-19 After Catch"
        'For I = 1 To 19
        '   For J = 1 To 3
        '      'Print #13, Format(CStr(CLng(TotCkEsc(J, I))), " @@@@@@");
        '   Next J
        '   'Print #13,
        'Next I
        '
        '------- Print Version Number and Command File Number ---
        'Print #13, VersNumb$ & " -BY=2"
        'Print #13, Mid(CMDFile$, 1, 4)
        'Print #13,
        'Print #13,

        '----- Print Terminal and Extreme Terminal Run Sizes ---
        For Stk = 1 To 17
            'Print #13, Format(Str(CLng(TotalChinEsc(3, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, Stk))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, Stk))), " @@@@@@@")
            Select Case Stk
                Case 1
                    'Print #13, Format(Str(CLng(TermChinAbun(1))), " @@@@@@@")
                Case 4
                    'Print #13, Format(Str(CLng(TermChinAbun(2))), " @@@@@@@")
                Case 9
                    'Print #13, Format(Str(CLng(TermChinAbun(3))), " @@@@@@@")
                Case 13
                    '--- Upper South Sound Yr.
                    'Print #13, Format(Str(CLng(TotalChinEsc(3, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 20))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 20))), " @@@@@@@")
                    '--- Deep South Sound Yr.
                    'Print #13, Format(Str(CLng(TotalChinEsc(3, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 21))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 21))), " @@@@@@@")
                    '--- Upper South Sound Agg.
                    'Print #13, Format(Str(CLng(TermChinAbun(6))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 18))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 18))), " @@@@@@@")
                    '--- Deep South Sound Agg.
                    'Print #13, Format(Str(CLng(TermChinAbun(7))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 19))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 19))), " @@@@@@@")
                    '--- Total TAA
                    'Print #13, Format(Str(CLng(TermChinAbun(4))), " @@@@@@@")
                Case 16
                    'Print #13, Format(Str(CLng(TermChinAbun(5))), " @@@@@@@")
                Case Else
            End Select
        Next Stk
        '- Add White River Spring to End
        'Print #13, Format(Str(CLng(TotalChinEsc(3, 14))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(1, 14))), " @@@@@@@") + Format(Str(CLng(TotalChinEsc(2, 14))), " @@@@@@@")

        '----------------------- GET TOTAL FISHERY MORTALITY DATA ---
        Dim AddFish(2, NumSteps + 1) As Double
        Dim TotLine, TotLCat, LndCatch, LndMort As Double
        For I = 1 To 2
            For J = 1 To 5
                AddFish(I, J) = 0.0
            Next J
        Next I
        '-========================================================
        Dim FishStep
        For Fish = 1 To NumFish
            TotLine = 0.0
            TotLCat = 0.0
            For FishStep = 1 To NumSteps
                LndCatch = 0.0
                LndMort = 0.0
                '      If TStep = 4 Then GoTo SkipTime1PrintBY2
                Select Case FishStep
                    Case 1
                        TStep = 4
                        GoTo SkipTime1PrintBY2
                    Case 2, 3
                        TStep = FishStep
                    Case 4
                        TStep = 1
                End Select
                For Stk = 1 To NumStk
                    For Age = MinAge To MaxAge
                        If Fish = 13 Or Fish = 14 Or (TammChinookRunFlag = 1 And Fish = 56) Then
                            AddFish(1, TStep) = AddFish(1, TStep) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            AddFish(2, TStep) = AddFish(2, TStep) + (BYLandedCatch(BY, Stk, Age, Fish, TStep) / ModelStockProportion(Fish))
                            AddFish(1, NumSteps + 1) = AddFish(1, NumSteps + 1) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)) / ModelStockProportion(Fish))
                            AddFish(2, NumSteps + 1) = AddFish(2, NumSteps + 1) + (BYLandedCatch(BY, Stk, Age, Fish, TStep) / ModelStockProportion(Fish))
                        Else                      '- Marked Catch & Mortality
                            LndMort = LndMort + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)
                            LndCatch = LndCatch + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        End If
                        TotLine = TotLine + BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)
                        TotLCat = TotLCat + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    Next Age
                Next Stk
SkipTime1PrintBY2:
                '- Print UnMarked Catch & Mortality
                If Fish = 13 Or Fish = 14 Or Fish = 56 Then
                    GoTo NextTStepBY2
                ElseIf Fish = 15 Or Fish = 57 Then '- Combined Fisheries
                    'Print #15, Format(Str(CLng((LndMort / ModelStockProportion(Fish) + AddFish(1, TStep)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((LndCatch / ModelStockProportion(Fish) + AddFish(2, TStep)))), " @@@@@@@");
                Else '- Normal Fishery Print UnMarked Fish
                    'Print #15, Format(Str(CLng((LndMort / ModelStockProportion(Fish)))), " @@@@@@@");
                    'Print #14, Format(Str(CLng((LndCatch / ModelStockProportion(Fish)))), " @@@@@@@");
                End If
NextTStepBY2:
            Next FishStep
            If Not (Fish = 13 Or Fish = 14 Or Fish = 56) Then
                'Print #15, Format(Str(CLng((TotLine / ModelStockProportion(Fish)))), " @@@@@@@")
                'Print #14, Format(Str(CLng((TotLCat / ModelStockProportion(Fish)))), " @@@@@@@")
            End If
            For I = 1 To 2
                For J = 1 To 5
                    AddFish(I, J) = 0.0
                Next J
            Next I
NextFishBY2:
        Next Fish
        '-======================================================== end total fish mort
        'For Fish = 1 To NumFish
        '   TotLine = 0!
        '   TotLCat = 0!
        '   For TStep = 1 To NumSteps
        '      If TStep = 1 Then GoTo SkipStep1
        '      If Fish = 13 Or Fish = 14 Or Fish = 56 Then
        '         AddFish(1, TStep) = AddFish(1, TStep) + ((TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalLegalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish))
        '         AddFish(2, TStep) = AddFish(2, TStep) + (TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish))
        '         AddFish(1, NumSteps + 1) = AddFish(1, NumSteps + 1) + ((TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalLegalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish))
        '         AddFish(2, NumSteps + 1) = AddFish(2, NumSteps + 1) + (TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish))
        '      Else
        '         If Fish = 15 Or Fish = 57 Then
        '            'Print #15, Format(Str(CLng(((TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalLegalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish)) + AddFish(1, TStep))), " @@@@@@@");
        '            'Print #14, Format(Str(CLng((TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish)) + AddFish(2, TStep))), " @@@@@@@");
        '         Else
        '            'Print #15, Format(Str(CLng((TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalLegalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish))), " @@@@@@@");
        '            'Print #14, Format(Str(CLng((TotalLandedCatch(Fish, TStep)) / ModelStockProportion(Fish))), " @@@@@@@");
        '         End If
        '         TotLine = TotLine + (TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalLegalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish)
        '         TotLCat = TotLCat + (TotalLandedCatch(Fish, TStep) / ModelStockProportion(Fish))
        '      End If
        'SkipStep1:
        '      If (TStep = 1 And Not (Fish = 13 Or Fish = 14 Or Fish = 56)) Then
        '         'Print #15, "       0";
        '         'Print #14, "       0";
        '      End If
        '   Next TStep
        '   If Fish = 15 Or Fish = 57 Then
        '      'Print #15, Format(Str(CLng(TotLine + AddFish(1, NumSteps + 1))), " @@@@@@@")
        '      'Print #14, Format(Str(CLng(TotLCat + AddFish(2, NumSteps + 1))), " @@@@@@@")
        '      For I = 1 To 2
        '         For J = 1 To 5
        '            AddFish(I, J) = 0!
        '         Next J
        '      Next I
        '   Else
        '      If Not (Fish = 13 Or Fish = 14 Or Fish = 56) Then
        '         'Print #15, Format(Str(CLng(TotLine)), " @@@@@@@")
        '         'Print #14, Format(Str(CLng(TotLCat)), " @@@@@@@")
        '      End If
        '   End If
        'Next Fish
        'Print #15,
        'Print #14,

        '-------------------------- STOCK CATCH BY FISHERY ---
        Dim FishNum As Integer
        Dim PageVals(5, 71) As Double
        Dim PageLCat(5, 71) As Double

        '---- ReInitialize Page Matrix -----
        For Stk = 1 To 18
            If Stk <> 3 Then
                'Print #15,
                'Print #14,
                For Fish = 1 To 71
                    For TStep = 1 To NumSteps + 1
                        PageVals(TStep, Fish) = 0
                        PageLCat(TStep, Fish) = 0
                    Next TStep
                Next Fish
            End If
            For Fish = 1 To NumFish
                Select Case Fish
                    Case 1 To 12
                        FishNum = Fish
                    Case 13, 14, 15
                        FishNum = 13
                    Case 16 To 55
                        FishNum = Fish - 2
                    Case 56, 57
                        FishNum = 54
                    Case 58 To 73
                        FishNum = Fish - 3
                End Select
                For TStep = 1 To NumSteps - 1
                    For Age = MinAge To MaxAge
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                            PageVals(TStep, FishNum) = PageVals(TStep, FishNum) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)))
                            PageVals(5, FishNum) = PageVals(5, FishNum) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)))
                        Else
                            PageVals(TStep, FishNum) = PageVals(TStep, FishNum) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
                            PageVals(5, FishNum) = PageVals(5, FishNum) + ((BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep))
                        End If
                        PageLCat(TStep, FishNum) = PageLCat(TStep, FishNum) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        PageLCat(5, FishNum) = PageLCat(5, FishNum) + BYLandedCatch(BY, Stk, Age, Fish, TStep)
                    Next Age
                Next TStep
            Next Fish
            '--- Print Page Matrix ----
            If Stk <> 2 Then
                For Fish = 1 To 70
                    For TStep = 1 To NumSteps + 1
                        Select Case TStep
                            Case 1
                                'Print #15, Format(Str(CLng(PageVals(4, Fish))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(4, Fish))), " @@@@@@@");
                            Case 2, 3, 5
                                'Print #15, Format(Str(CLng(PageVals(TStep, Fish))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(TStep, Fish))), " @@@@@@@");
                            Case 4
                                'Print #15, Format(Str(CLng(PageVals(1, Fish))), " @@@@@@@");
                                'Print #14, Format(Str(CLng(PageLCat(1, Fish))), " @@@@@@@");
                        End Select
                    Next TStep
                    'Print #15,
                    'Print #14,
                Next Fish
                'Print #15,
                'Print #14,
            End If
        Next Stk
        'Close #15
        'Close #14
        'Close #13
        'Close #2

   End Sub

   ''Pete Added Dec 2013 -- Commented out to Deactivate Feb 2013
   'Sub CompExternalChinookShakers(ByVal TerminalType, ByVal Fish)

   '   Dim SublegalStkFraction(NumStk, MaxAge, NumFish, NumSteps), SublegalTotal As Double 'Stock-age assignment variables if no differing size limit in combo
   '   Dim SublegalStkFractionNS(NumStk, MaxAge, NumFish, NumSteps), SublegalNSTotal As Double 'Stock-age assignment variables for Nonselective in combo with unequal limits
   '   Dim SublegalStkFractionMSF(NumStk, MaxAge, NumFish, NumSteps), SublegalMSFTotal As Double 'Stock-age assignment variables for Mark Selective in combo with unequal limits
   '   Dim ShakersNSOnly, ShakersMSFOnly
   '   Dim NSSublegalPopulation, NSLegalPopulation, NSSublegalPop, NSLegalPop
   '   Dim MSFSublegalPopulation, MSFLegalPopulation, MSFSublegalPop, MSFLegalPop

   '   'Zero these out to avoid any trouble with the law
   '   ShakersNSOnly = 0
   '   ShakersMSFOnly = 0
   '   NSLegalPopulation = 0
   '   NSSublegalPopulation = 0
   '   MSFLegalPopulation = 0
   '   MSFSublegalPopulation = 0

   '   'Step 1:    Compute total shaker encounters for fishery based on inputs loaded by the btnLimitChange_Click()
   '   '           subroutine in FVS_SizeLimitEdit.vb, if appropriate 

   '   Debug.Print(FisheryName(Fish) & "  TS " & TStep & "  " & TotalShakers(Fish, TStep) & " Before")

   '   If AltFlag(Fish, TStep) = 1 Or AltFlag(Fish, TStep) = 2 Then
   '      If ShakerFlagNS(Fish, TStep) = 2 Then
   '         NSShakerExtTotal(Fish, TStep) = LSRatioNS(Fish, TStep) * NSEncountersTotal(Fish, TStep)
   '      ElseIf ShakerFlagNS(Fish, TStep) = 3 Then
   '         NSShakerExtTotal(Fish, TStep) = ExtShakerNS(Fish, TStep) * ModelStockProportion(Fish)
   '      End If
   '   ElseIf AltFlag(Fish, TStep) = 7 Or AltFlag(Fish, TStep) = 8 Then
   '      If ShakerFlagMSF(Fish, TStep) = 2 Then
   '         MSFShakerExtTotal(Fish, TStep) = LSRatioMSF(Fish, TStep) * MSFEncountersTotal(Fish, TStep)
   '      ElseIf ShakerFlagMSF(Fish, TStep) = 3 Then
   '         MSFShakerExtTotal(Fish, TStep) = ExtShakerMSF(Fish, TStep) * ModelStockProportion(Fish)
   '      End If
   '   ElseIf AltFlag(Fish, TStep) > 8 Then
   '      If ShakerFlagMSF(Fish, TStep) = 2 Then
   '         MSFShakerExtTotal(Fish, TStep) = LSRatioMSF(Fish, TStep) * MSFEncountersTotal(Fish, TStep)
   '      ElseIf ShakerFlagMSF(Fish, TStep) = 3 Then
   '         MSFShakerExtTotal(Fish, TStep) = ExtShakerMSF(Fish, TStep) * ModelStockProportion(Fish)
   '      End If
   '      If ShakerFlagNS(Fish, TStep) = 2 Then
   '         NSShakerExtTotal(Fish, TStep) = LSRatioNS(Fish, TStep) * NSEncountersTotal(Fish, TStep)
   '      ElseIf ShakerFlagNS(Fish, TStep) = 3 Then
   '         NSShakerExtTotal(Fish, TStep) = ExtShakerNS(Fish, TStep) * ModelStockProportion(Fish)
   '      End If
   '   End If


   '   ' Step 2:   Get the proportion of shaker mortalities expected for each stock and age based on the existing shaker algorithm,
   '   '           parse out external estimates accordingly, and replace the internally computed values.


   '   If AltFlag(Fish, TStep) < 8 Or AltLimitMSF(Fish, TStep) = AltLimitNS(Fish, TStep) Then
   '      'Use the same shaker stock-age comp for NS & MSF fisheries (and combo fisheries, if size limits are equivalent)

   '      'Compute stock-age proportions for the fishery
   '      For Stk = 1 To NumStk
   '         For Age = MinAge To MaxAge
   '            SublegalStkFraction(Stk, Age, Fish, TStep) = (Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) / TotalShakers(Fish, TStep)
   '         Next Age
   '      Next Stk

   '      'Wipe out internally computed shakers and replace with external value apportioned to the stock-age group in question
   '      For Stk = 1 To NumStk
   '         For Age = MinAge To MaxAge
   '            If AltFlag(Fish, TStep) = 1 Or AltFlag(Fish, TStep) = 2 Or AltFlag(Fish, TStep) = 17 Or AltFlag(Fish, TStep) = 18 Or AltFlag(Fish, TStep) = 27 Or AltFlag(Fish, TStep) = 28 Then
   '               If ShakerFlagNS(Fish, TStep) = 2 Or ShakerFlagNS(Fish, TStep) = 3 Then
   '                  TotalShakers(Fish, TStep) -= Shakers(Stk, Age, Fish, TStep)
   '                  Shakers(Stk, Age, Fish, TStep) = NSShakerExtTotal(Fish, TStep) * SublegalStkFraction(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep)
   '                  TotalShakers(Fish, TStep) += Shakers(Stk, Age, Fish, TStep)
   '                  'totNS = Shakers(Stk, Age, Fish, TStep)
   '               End If
   '            End If
   '            If AltFlag(Fish, TStep) = 7 Or AltFlag(Fish, TStep) = 8 Or AltFlag(Fish, TStep) = 17 Or AltFlag(Fish, TStep) = 18 Or AltFlag(Fish, TStep) = 27 Or AltFlag(Fish, TStep) = 28 Then
   '               If ShakerFlagMSF(Fish, TStep) = 2 Or ShakerFlagMSF(Fish, TStep) = 3 Then
   '                  TotalShakers(Fish, TStep) -= MSFShakers(Stk, Age, Fish, TStep)
   '                  MSFShakers(Stk, Age, Fish, TStep) = MSFShakerExtTotal(Fish, TStep) * SublegalStkFraction(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep)
   '                  TotalShakers(Fish, TStep) += MSFShakers(Stk, Age, Fish, TStep)
   '                  'Debug.Print("Stock, " & StockName(Stk) & ",Age " & Age & ",Fish " & Fish & ",TS " & TStep & ",MSFShakers, " & MSFShakers(Stk, Age, Fish, TStep) & ",NSShakers, " & totNS)
   '               End If
   '            End If
   '         Next
   '      Next

   '   Else  '  i.e., this is what happens if it's a combo fishery with a different size limit during MSF and NS periods
   '      '     It's slightly clunky and replicates CompShakers() with two limits, however given the rarity of combo fisheries
   '      '     with different MSF and NS limits, it shouldn't be invoked very often...

   '      For Stk = 1 To NumStk
   '         For Age = MinAge To MaxAge
   '            Call CompLegProp(Stk, Age, Fish, TerminalType)


   '            NSLegalPopulation = NSLegalPopulation + Cohort(Stk, Age, TerminalType, TStep) * NSLegalProp
   '            NSLegalPop = Cohort(Stk, Age, TerminalType, TStep) * NSLegalProp
   '            '- Zero Time 1 Yearling Shakers ...
   '            '- Fish not Recruited Yet - Temp Fix 1/3/2000 JFP
   '            If NumStk < 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 5 Or Stk = 6 Or Stk = 8 Or Stk = 14 Or Stk = 17 Or Stk = 25) Then
   '               '- Regular Chinook FRAM
   '               NSSublegalPop = 0
   '               MSFSublegalPop = 0
   '            ElseIf NumStk > 50 And Age = 2 And (TStep = 1 Or TStep = 4) And (Stk = 9 Or Stk = 10 Or Stk = 11 Or Stk = 12 Or Stk = 15 Or Stk = 16 Or Stk = 27 Or Stk = 28 Or Stk = 33 Or Stk = 34 Or Stk = 49 Or Stk = 50) Then
   '               '- Selective Fishery Version
   '               NSSublegalPop = 0
   '               MSFSublegalPop = 0
   '            Else
   '               NSSublegalPop = Cohort(Stk, Age, TerminalType, TStep) * NSSublegalProp
   '               MSFSublegalPop = Cohort(Stk, Age, TerminalType, TStep) * MSFSublegalProp
   '            End If
   '            NSSublegalPopulation = NSSublegalPopulation + NSSublegalPop
   '            MSFSublegalPopulation = MSFSublegalPopulation + MSFSublegalPop

   '            '- Retention Fishery Shaker Calculation
   '            Shakers(Stk, Age, Fish, TStep) = FisheryScaler(Fish, TStep) * NSSublegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
   '            TotalShakers(Fish, TStep) = TotalShakers(Fish, TStep) + Shakers(Stk, Age, Fish, TStep)
   '            ShakersNSOnly = ShakersNSOnly + Shakers(Stk, Age, Fish, TStep)

   '            '- MSF Shaker Calculation
   '            MSFShakers(Stk, Age, Fish, TStep) = MSFFisheryScaler(Fish, TStep) * MSFSublegalPop * BaseSubLegalRate(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep) * StockFishRateScalers(Stk, Fish, TStep)
   '            TotalShakers(Fish, TStep) = TotalShakers(Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
   '            ShakersMSFOnly = ShakersMSFOnly + MSFShakers(Stk, Age, Fish, TStep)

   '         Next Age
   '      Next Stk



   '      For Stk = 1 To NumStk
   '         For Age = MinAge To MaxAge
   '            SublegalStkFractionNS(Stk, Age, Fish, TStep) = (Shakers(Stk, Age, Fish, TStep)) / ShakersNSOnly 'use separate stock-age fractions given different size limits
   '            SublegalStkFractionMSF(Stk, Age, Fish, TStep) = (MSFShakers(Stk, Age, Fish, TStep)) / ShakersMSFOnly 'use separate stock-age fractions given different size limits

   '            If ShakerFlagNS(Fish, TStep) = 2 Or ShakerFlagNS(Fish, TStep) = 3 Then
   '               TotalShakers(Fish, TStep) -= Shakers(Stk, Age, Fish, TStep)
   '               Shakers(Stk, Age, Fish, TStep) = NSShakerExtTotal(Fish, TStep) * SublegalStkFractionNS(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep)
   '               TotalShakers(Fish, TStep) += Shakers(Stk, Age, Fish, TStep)
   '            End If

   '            If ShakerFlagMSF(Fish, TStep) = 2 Or ShakerFlagMSF(Fish, TStep) = 3 Then
   '               TotalShakers(Fish, TStep) -= MSFShakers(Stk, Age, Fish, TStep)
   '               MSFShakers(Stk, Age, Fish, TStep) = MSFShakerExtTotal(Fish, TStep) * SublegalStkFractionMSF(Stk, Age, Fish, TStep) * ShakerMortRate(Fish, TStep)
   '               TotalShakers(Fish, TStep) += MSFShakers(Stk, Age, Fish, TStep)
   '            End If
   '         Next
   '      Next

   '   End If

   '   Debug.Print(FisheryName(Fish) & "  TS " & TStep & "  " & TotalShakers(Fish, TStep) & " After")

   'End Sub


End Module

