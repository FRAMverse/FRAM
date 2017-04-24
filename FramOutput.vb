Imports System
Imports System.IO
Imports System.Text
Imports System.IO.File
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Module FramOutput

   Public rw As StreamWriter
   Public rr As StreamReader
   Public PrnLine As String

   Sub RunReportDriver()

      'Dim Result As Integer
      Dim TerminalRunReportSorted As Boolean

      '- OPEN Report Output File
      'If Exists(ReportFileName) Then
      '   Result = MsgBox("Report FileName Already EXISTS" & vbCrLf & "OK to Overwrite ???", MsgBoxStyle.YesNo)
      '   If Result = vbYes Then
      '      Delete(ReportFileName)
      '   Else
      '      Exit Sub
      '   End If
      'End If
      rw = CreateText(ReportFileName)

      '- Read User Selected ReportDriver Data
      Dim CmdStr As String
      Dim Option1, Option2, Option3, Option4, Option5, Option6 As String
      Dim ReportNumber, RecNum As Integer
      TerminalRunReportSorted = False
      CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName.ToString & Chr(34) & " ORDER BY DriverName,ReportNumber,Option1,Option2,Option3,Option4,Option5,Option6"
SortTermRun:
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
         If ReportNumber = 2 And TerminalRunReportSorted = False Then
            CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName.ToString & Chr(34) & " ORDER BY ReportNumber, Option6"
            TerminalRunReportSorted = True
            GoTo SortTermRun
         End If
         If ReportNumber > 2 And TerminalRunReportSelected = True Then
            '- Finish Terminal Run Report after all Options Selected 
            Option6 = "FINISH"
            NumRepGrps += 1
            Call TerminalRunReport(Option1, Option2, Option3, Option4, Option5, Option6)
            TerminalRunReportSelected = False
         End If
         '- Stock Catch Report needs to be sorted per user selection order
         If ReportNumber = 3 And TerminalRunReportSorted = False Then
            CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName.ToString & Chr(34) & " ORDER BY ReportNumber, Option5"
            TerminalRunReportSorted = True
            GoTo SortTermRun
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

         If ReportNumber = 2 And RecNum = NumRD - 1 Then
            '- Finish Terminal Run Report when run out of driver records
            If InStr(Option4, "Brood") > 0 Then
               TermRunBYAEQ = True
            Else
               TermRunBYAEQ = False
            End If
            If InStr(Option4, "Brood") > 0 Then
               TermRunBYAEQ = True
            Else
               TermRunBYAEQ = False
            End If
            Option6 = "FINISH"
            NumRepGrps += 1
            If NumRD = 1 Then
               ReDim RepStks(300, NumStk)
               ReDim NumRepStks(300)
               ReDim RepFish(300, NumFish)
               ReDim NumRepFish(300)
               ReDim RepTStep(300, 2)
               ReDim RepGrpName(300)
               ReDim RepGrpType(300)
            End If
            Call TerminalRunReport(Option1, Option2, Option3, Option4, Option5, Option6)
            TerminalRunReportSelected = False
            GoTo ReportDone
         End If

         Select Case ReportNumber
            Case 1
               Call MortalitySummaryReport("Fishery", Option1, Option2, Option3, Option4, Option5, Option6)
            Case 2
               NumRepGrps += 1
               If InStr(Option4, "Brood") > 0 Then
                  TermRunBYAEQ = True
               Else
                  TermRunBYAEQ = False
               End If
               Option6 = "GROUP"
               '- On First Call ReDim Arrays for Max of 300 Groups
               If TerminalRunReportSelected = False Then
                  ReDim RepStks(300, NumStk)
                  ReDim NumRepStks(300)
                  ReDim RepFish(300, NumFish)
                  ReDim NumRepFish(300)
                  ReDim RepTStep(300, 2)
                  ReDim RepGrpName(300)
                  ReDim RepGrpType(300)
                  Call TerminalRunReport(Option1, Option2, Option3, Option4, Option5, Option6)
                  TerminalRunReportSelected = True
               Else
                  Call TerminalRunReport(Option1, Option2, Option3, Option4, Option5, Option6)
               End If
            Case 3
               Call MortalitySummaryReport("Stock", Option1, Option2, Option3, Option4, Option5, Option6)
            Case 5
               Call MortAgeReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 6
               Call FisheryScalerReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 7
               Call StockSummaryReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 8
               Call PopulationStatisticsReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 9
               Call ERDistributionReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 14
               Call SelectiveFisheryReport(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 15
               Call PSCCohoER(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 16
               If SpeciesName = "COHO" Then
                  Call CohoStockER(Option1, Option2, Option3, Option4, Option5, Option6)
               ElseIf SpeciesName = "CHINOOK" Then
                  Call ChinookStockER(Option1, Option2, Option3, Option4, Option5, Option6)
               End If
            Case 17
               Call FisheryStockComposition(Option1, Option2, Option3, Option4, Option5, Option6)
            Case 18
               Call StockImpactsPer1000(Option1, Option2, Option3, Option4, Option5, Option6)
         End Select
      Next

ReportDone:

      ReportDA = Nothing
      rw.Close()

   End Sub

   Sub ReadOldDriverFile()

      '- Text File Reader
      Dim DRVReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(OldDriverFileName)
      DRVReader.TextFieldType = FileIO.FieldType.Delimited
      '- DRV File Format has Blanks for some Field Separators
      DRVReader.SetDelimiters(",", " ")

      Dim CurrentRow As String()
      Dim RepSpecies As String
      Dim CurrentField, ReportField As String
      Dim LineNum, NumComms, LoopLen, FieldNum As Integer
      Dim Comments, ReportOption, TermName As String
      Dim Option1, Option2, Option3, Option4, Option5, Option6 As String
      Dim ReportNumber, NumReps, RepNum, Rep, Mortalitytype As Integer
      Dim NumRepStk, NumRepFish, NumRepGrps As Integer
      Dim RepStks(1), RepFish(1), RepTime1, RepTime2, RepType As Integer
      Dim NumStkCatReps As Integer
      Dim BroodYearStyle As Boolean

      BroodYearStyle = False
      LineNum = 1
      Comments = ""
      ReportDriverName = ""
      ReportDriverName = My.Computer.FileSystem.GetFileInfo(OldDriverFileName).Name

      '- Check if Driver Name already in use & Setup Data Adapter Commands
      Dim CmdStr As String
      CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName & Chr(34)
      Dim DrvDA As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim DriverDA As New System.Data.OleDb.OleDbDataAdapter
      DriverDA.SelectCommand = DrvDA

      CmdStr = "DELETE * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName & Chr(34)
      Dim DrvDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      DriverDA.DeleteCommand = DrvDcm

      Dim DRVcb As New OleDb.OleDbCommandBuilder
      DRVcb = New OleDb.OleDbCommandBuilder(DriverDA)

      '- Check if Report Driver Name is already in use
      If FramDataSet.Tables.Contains("ReportDriver") Then
         FramDataSet.Tables("ReportDriver").Rows.Clear()
      End If
      DriverDA.Fill(FramDataSet, "ReportDriver")
      Dim NumRD As Integer
      NumRD = FramDataSet.Tables("ReportDriver").Rows.Count
      Dim Result
      If NumRD <> 0 Then
         Result = MsgBox("Report Driver Name already in use" & vbCrLf & "Do You Want to OverWrite???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            FramDB.Open()
            DriverDA.DeleteCommand.ExecuteScalar()
            DriverDA = Nothing
         Else
            DriverDA = Nothing
            Exit Sub
         End If
      End If

      '========================================
      '- Open ReportDriver Table in DataBase to Store Parameters
      Dim DrvTrans As OleDb.OleDbTransaction
      Dim DFC As New OleDbCommand
      If FramDB.State <> ConnectionState.Open Then
         FramDB.Open()
      End If
      DrvTrans = FramDB.BeginTransaction
      DFC.Connection = FramDB
      DFC.Transaction = DrvTrans

      '========================================

      '- Read DRV Header Information ...
      '   Species, ReportDriverName, Comments, BasePeriod, and Stock Recruits
      RepSpecies = ""
      NumStkCatReps = 0
      While Not DRVReader.EndOfData
         'Try
         CurrentRow = DRVReader.ReadFields()
         LoopLen = CurrentRow.Length
         FieldNum = 0
         '- Fields are Line Number Specific
         Select Case LineNum
            Case 1
               '- SpeciesName
               RepSpecies = Trim(CurrentRow(0))
               If RepSpecies = "COHO" Then
                  MsgBox("The Old COHO DRIVER Files and Base Periods" & vbCrLf & "are not supported", MsgBoxStyle.OkOnly)
                  Exit Sub
               End If
               If RepSpecies = "NEWCOHO" Then
                  '- NEWCOHO was used in Previous FRAM Versions
                  RepSpecies = "COHO"
               ElseIf RepSpecies = "CHINOOK" Then
                  Jim = 1
               Else
                  MsgBox("Species Name not recognized for this Command File!!!", MsgBoxStyle.OkOnly)
                  Exit Sub
               End If
            Case 2 To 6
               '- NumStk, NumFish, NumSteps, MaxAge, EncRt, Lines
            Case 7
               For Each CurrentField In CurrentRow
                  If CurrentField <> "" Then
                     NumComms = CInt(CurrentField)
                     Exit For
                  End If
               Next
            Case 8 To (NumComms + 7)
               For Each CurrentField In CurrentRow
                  Comments &= CurrentField
                  If RepSpecies = "CHINOOK" Then
                     If InStr(Comments, "Brood", CompareMethod.Text) > 0 Then
                        BroodYearStyle = True
                     End If
                  End If
               Next
            Case Is > (NumComms + 7)
               '- Report Description Lines of DRIVER File
               FieldNum = 1
               CurrentField = ""
               For Each CurrentField In CurrentRow
                  If CurrentField = "" Then GoTo NextLineField
                  If FieldNum = 1 Then
                     ReportNumber = CInt(CurrentField)
                     FieldNum += 1
                  ElseIf FieldNum = 2 Then
                     ReportOption = CurrentField
                     If ReportOption = "N" Then GoTo ReportNotSelected
                     FieldNum += 1
                     '- Reports Selected with NO Options (no more fields)
                     '  StockSummary, PopStat, SelectiveFishery, PSC Coho ER
                     If ReportNumber = 7 Or ReportNumber = 8 Or ReportNumber = 14 Or ReportNumber = 15 Then
                        DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber) " & _
                           "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                           Chr(34) & RepSpecies & Chr(34) & "," & _
                           ReportNumber.ToString & ")"
                        DFC.ExecuteNonQuery()
                     End If
                  Else
                     '- Report Options (if Selected)
                     Select Case ReportNumber

                        Case 1
                           '- Fishery Summary Report
                           Mortalitytype = CInt(CurrentField)
                           Option1 = Mortalitytype.ToString
                           '- Other Fields not needed .. Write Table Record
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & "," & _
                              Chr(34) & Option1 & Chr(34) & ")"
                           DFC.ExecuteNonQuery()
                           GoTo ReportNotSelected

                        Case 2
                           '- Terminal Run Report
                           NumReps = CInt(CurrentField)
                           For Rep = 1 To NumReps
                              FieldNum = 1
                              CurrentRow = DRVReader.ReadFields()
                              For Each ReportField In CurrentRow
                                 If ReportField = "" Then GoTo NextReportField
                                 Select Case FieldNum
                                    Case 1
                                       RepNum = ReportField
                                       If RepNum <> ReportNumber Then
                                          MsgBox("ERROR in TermRunRep of DRV File = " & OldDriverFileName & vbCrLf & "Existing DRV File Read", MsgBoxStyle.OkOnly)
                                          Exit Sub
                                       End If
                                    Case 2
                                       NumRepStk = CInt(ReportField)
                                       ReDim RepStks(NumRepStk)
                                    Case 3 To (NumRepStk + 2)
                                       RepStks(FieldNum - 2) = CInt(ReportField)
                                    Case (NumRepStk + 3)
                                       NumRepFish = CInt(ReportField)
                                       ReDim RepFish(NumRepFish)
                                    Case (NumRepStk + 4) To (NumRepStk + 3 + NumRepFish)
                                       RepFish(FieldNum - NumRepStk - 3) = CInt(ReportField)
                                    Case (NumRepStk + 4 + NumRepFish)
                                       RepTime1 = CInt(ReportField)
                                    Case (NumRepStk + 5 + NumRepFish)
                                       RepTime2 = CInt(ReportField)
                                    Case (NumRepStk + 6 + NumRepFish)
                                       RepType = CInt(Mid(ReportField, 1, 2))
                                    Case (NumRepStk + 7 + NumRepFish)
                                       TermName = ReportField
                                       '=========================================
                                       Option1 = ""
                                       For Stk As Integer = 1 To NumRepStk
                                          Option1 &= RepStks(Stk).ToString
                                          If Stk <> NumRepStk Then Option1 &= ","
                                       Next
                                       Option2 = ""
                                       For Fish As Integer = 1 To NumRepFish
                                          Option2 &= RepFish(Fish).ToString
                                          If Fish <> NumRepFish Then Option2 &= ","
                                       Next
                                       Option3 = RepTime1.ToString & "," & RepTime2.ToString
                                       If SpeciesName = "CHINOOK" And BroodYearStyle = True Then
                                          If RepType = 0 Then
                                             Option4 = "ETRS Brood Year"
                                          Else
                                             Option4 = "TAA Brood Year"
                                          End If
                                       Else
                                          If RepType = 0 Then
                                             Option4 = "ETRS"
                                          Else
                                             Option4 = "TAA"
                                          End If
                                       End If
                                       Option5 = TermName
                                       If Rep < 10 Then
                                          Option6 = "00" & Rep.ToString
                                       ElseIf Rep > 9 And Rep < 100 Then
                                          Option6 = "0" & Rep.ToString
                                       Else
                                          Option6 = Rep.ToString
                                       End If
                                       DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1,Option2,Option3,Option4,Option5,Option6) " & _
                                          "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                                          Chr(34) & RepSpecies & Chr(34) & "," & _
                                          ReportNumber.ToString & "," & _
                                          Chr(34) & Option1 & Chr(34) & "," & _
                                          Chr(34) & Option2 & Chr(34) & "," & _
                                          Chr(34) & Option3 & Chr(34) & "," & _
                                          Chr(34) & Option4 & Chr(34) & "," & _
                                          Chr(34) & Option5 & Chr(34) & "," & _
                                          Chr(34) & Option6 & Chr(34) & ")"
                                       DFC.ExecuteNonQuery()
                                       '==========================================
                                 End Select
                                 FieldNum += 1
NextReportField:
                              Next
                           Next

                        Case 3
                           '- Stock Catch Report
                           NumStkCatReps += 1
                           If FieldNum = 3 Then
                              NumRepGrps = CInt(CurrentField)
                              FieldNum += 1
                              GoTo NextLineField
                           ElseIf FieldNum = 4 Then
                              Mortalitytype = CInt(CurrentField)
                           End If
                           '- First Get Line with Fishery Selection .. No Longer Needed
                           CurrentRow = DRVReader.ReadFields()
                           'ReDim RepStks(NumRepGrps)
                           For Rep = 1 To NumRepGrps
                              FieldNum = 1
                              TermName = ""
                              '- Next Get Line with Stock Selections
                              CurrentRow = DRVReader.ReadFields()
                              For Each ReportField In CurrentRow
                                 If ReportField = "" Then GoTo NextStockCatchField
                                 Select Case FieldNum
                                    Case 1
                                       RepNum = ReportField
                                       If RepNum <> ReportNumber Then
                                          MsgBox("ERROR in Stock Catch Rep of DRV File = " & OldDriverFileName & vbCrLf & "Existing DRV File Read", MsgBoxStyle.OkOnly)
                                          Exit Sub
                                       End If
                                    Case 2
                                       NumRepStk = CInt(ReportField)
                                       If NumRepStk < 50 Then
                                          ReDim RepStks(NumRepStk)
                                       Else
                                          ReDim RepStks(NumStk)
                                       End If
                                    Case 3 To (NumRepStk + 2)
                                       '- Option 2 Field too large if too many Stocks are Selected
                                       If NumRepStk < 50 Then
                                          RepStks(FieldNum - 2) = CInt(ReportField)
                                       Else
                                          RepStks(CInt(ReportField)) = 1
                                       End If
                                    Case NumRepStk + 3
                                       TermName = ReportField
                                 End Select
                                 FieldNum += 1
NextStockCatchField:
                              Next
                              '- Put Each StkCatch Rep into Table
                              Option1 = Mortalitytype.ToString
                              Option2 = NumRepStk.ToString
                              Option3 = ""
                              If NumRepStk < 50 Then
                                 For Stk = 1 To NumRepStk
                                    Option3 &= RepStks(Stk).ToString
                                    If Stk <> NumRepStk Then Option3 &= ","
                                 Next
                              Else
                                 For Stk = 1 To NumStk
                                    Option3 &= RepStks(Stk).ToString("0")
                                 Next
                              End If
                              Option4 = TermName
                              If Rep < 10 Then
                                 Option5 = "0" & Rep.ToString
                              Else
                                 Option5 = Rep.ToString
                              End If
                              '- Other Fields not needed .. Write Table Record
                              DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1,Option2,Option3,Option4,Option5) " & _
                                 "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                                 Chr(34) & RepSpecies & Chr(34) & "," & _
                                 ReportNumber.ToString & "," & _
                                 Chr(34) & Option1 & Chr(34) & "," & _
                                 Chr(34) & Option2 & Chr(34) & "," & _
                                 Chr(34) & Option3 & Chr(34) & "," & _
                                 Chr(34) & Option4 & Chr(34) & "," & _
                                 Chr(34) & Option5 & Chr(34) & ")"
                              DFC.ExecuteNonQuery()
                           Next

                        Case 4
                           '- ER Comparison Report

                        Case 5
                           '- StockMortAge Report
                           If FieldNum = 3 Then
                              NumRepFish = CInt(CurrentField)
                              ReDim RepFish(NumRepFish)
                              FieldNum += 1
                              GoTo NextLineField
                           End If
                           If FieldNum = 4 Then
                              NumRepStk = CInt(CurrentField)
                              ReDim RepStks(NumRepStk)
                              FieldNum += 1
                              GoTo NextLineField
                           End If
                           If FieldNum = 5 Then
                              Mortalitytype = CInt(CurrentField)
                           End If
                           '- Read Fishery Field with Zero and Ones
                           CurrentRow = DRVReader.ReadFields()
                           FieldNum = 1
                           Option2 = ""
                           For Each ReportField In CurrentRow
                              If ReportField = "" Then GoTo NextFishAgeField
                              Select Case FieldNum
                                 Case 1
                                    Rep = ReportField
                                    If Rep <> 5 Then
                                       'problem!
                                    End If
                                    FieldNum += 1
                                 Case 2
                                    'Rep = 1
                                    'For Fish = 1 To NumFish
                                    '   If ReportField.Substring(Fish - 1, 1) = "1" Then
                                    '      RepFish(Rep) = Fish
                                    '      Rep += 1
                                    '   End If
                                    'Next
                                    Option2 = ReportField
                              End Select
NextFishAgeField:
                           Next

                           '- Read Stock Field with Zero and Ones
                           FieldNum = 1
                           CurrentRow = DRVReader.ReadFields()
                           Option3 = ""
                           For Each ReportField In CurrentRow
                              If ReportField = "" Then GoTo NextStockAgeField
                              Select Case FieldNum
                                 Case 1
                                    Rep = ReportField
                                    If Rep <> 5 Then
                                       'problem!
                                    End If
                                    FieldNum += 1
                                 Case 2
                                    'Rep = 1
                                    'For Stk = 1 To NumStk
                                    '   If ReportField.Substring(Stk - 1, 1) = "1" Then
                                    '      RepStks(Rep) = Stk
                                    '      Rep += 1
                                    '   End If
                                    'Next
                                    Option3 = ReportField
                              End Select
NextStockAgeField:
                           Next

                           Option1 = ""
                           Option1 = CStr(Mortalitytype)
                           'Option2 = ""
                           'For Stk = 1 To NumRepStk
                           '   Option2 &= RepStks(Stk).ToString
                           '   If Stk <> NumRepStk Then Option2 &= ","
                           'Next
                           'Option3 = ""
                           'For Fish = 1 To NumRepFish
                           '   Option3 &= RepFish(Fish).ToString
                           '   If Fish <> NumRepFish Then Option3 &= ","
                           'Next
                           '- Other Fields not needed .. Write Table Record
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1,Option2,Option3) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & "," & _
                              Chr(34) & Option1 & Chr(34) & "," & _
                              Chr(34) & Option2 & Chr(34) & "," & _
                              Chr(34) & Option3 & Chr(34) & ")"
                           DFC.ExecuteNonQuery()

                        Case 6
                           '- Fishery Scaler Report
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & ")"
                           DFC.ExecuteNonQuery()
                           GoTo ReportNotSelected

                        Case 7
                           '- Stock Summary Report

                        Case 8
                           '- Population Statistics Report
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & ")"
                           DFC.ExecuteNonQuery()
                           GoTo ReportNotSelected

                        Case 9
                           '- ER Distribution Report
                        Case 10
                           '- Total ER Report
                        Case 11
                           '- CAM-Coho Summary Report
                        Case 12
                           '- CAM-Coho Esc Report
                        Case 13
                           '- CAM-Coho Coastal Report

                        Case 14
                           '- Selective Fishery Report
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & ")"
                           DFC.ExecuteNonQuery()

                        Case 15
                           '- PSC Coho ER Report
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & ")"
                           DFC.ExecuteNonQuery()

                        Case 16
                           '- Stock ER & Dist Report
                           NumRepGrps = CInt(CurrentField)
                           ReDim RepStks(NumRepGrps)
                           For Rep = 1 To NumRepGrps
                              FieldNum = 1
                              TermName = ""
                              '- Get Line with Stock Selections
                              CurrentRow = DRVReader.ReadFields()
                              For Each ReportField In CurrentRow
                                 If ReportField = "" Then GoTo NextStockERField
                                 Select Case FieldNum
                                    Case 1
                                       RepNum = ReportField
                                       If RepNum <> ReportNumber Then
                                          MsgBox("ERROR in Stock ER Rep of DRV File = " & OldDriverFileName & vbCrLf & "Existing DRV File Read", MsgBoxStyle.OkOnly)
                                          Exit Sub
                                       End If
                                    Case 2
                                       NumRepStk = CInt(ReportField)
                                       ReDim RepStks(NumRepStk)
                                    Case 3 To (NumRepStk + 2)
                                       RepStks(FieldNum - 2) = CInt(ReportField)
                                    Case NumRepStk + 3
                                       TermName = ReportField
                                 End Select
                                 FieldNum += 1
NextStockERField:
                              Next
                              '- Put Each StkCatch Rep into Table
                              Option1 = Rep.ToString
                              Option2 = ""
                              For Stk = 1 To NumRepStk
                                 Option2 &= RepStks(Stk).ToString
                                 If Stk <> NumRepStk Then Option2 &= ","
                              Next
                              Option3 = TermName
                              '- Write Table Record
                              DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1,Option2,Option3) " & _
                                 "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                                 Chr(34) & RepSpecies & Chr(34) & "," & _
                                 ReportNumber.ToString & "," & _
                                 Chr(34) & Option1 & Chr(34) & "," & _
                                 Chr(34) & Option2 & Chr(34) & "," & _
                                 Chr(34) & Option3 & Chr(34) & ")"
                              DFC.ExecuteNonQuery()
                           Next

                        Case 17
                           '- Fishery Stock Composition Report
                           If FieldNum = 3 Then
                              NumRepFish = CInt(CurrentField)
                              ReDim RepFish(NumRepFish)
                              FieldNum += 1
                              GoTo NextLineField
                           End If
                           '- Read Fishery Field with Zero and Ones
                           Rep = 1
                           For Fish As Integer = 1 To NumFish
                              If CurrentField.Substring(Fish - 1, 1) = "1" Then
                                 RepFish(Rep) = Fish
                                 Rep += 1
                              End If
                           Next
                           '- Put Fishery-Stock-Comp Rep into Table
                           Option1 = ""
                           For Fish As Integer = 1 To NumRepFish
                              Option1 &= RepFish(Fish).ToString
                              If Fish <> NumRepFish Then Option1 &= ","
                           Next
                           '- Other Fields not needed .. Write Table Record
                           DFC.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1) " & _
                              "VALUES(" & Chr(34) & ReportDriverName & Chr(34) & "," & _
                              Chr(34) & RepSpecies & Chr(34) & "," & _
                              ReportNumber.ToString & "," & _
                              Chr(34) & Option1 & Chr(34) & ")"
                           DFC.ExecuteNonQuery()

                     End Select

                  End If
NextLineField:
               Next

         End Select

         'Catch ex As Exception
         '   MsgBox("The DRIVER FILE file Selected has Format Problems" & vbCrLf & " RecNum=" & RecNum.ToString, MsgBoxStyle.OkOnly)
         '   Exit Sub
         'End Try

ReportNotSelected:
         LineNum += 1

      End While

      DrvTrans.Commit()
      FramDB.Close()
      DRVReader.Close()

   End Sub

   Sub MortalitySummaryReport(ByVal TypeRep As String, ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      Dim ParseOld, ParseNew, Report, RepStkNum, BY As Integer
      Dim RepStks(NumStk), NumRepStks As Integer
      Dim RepGrpName As String
      'Dim TempCatch As Double

      RepGrpName = ""
      NumRepStks = 0
      If TypeRep = "Fishery" Then
         MortalityType = CInt(Opt1)
      ElseIf TypeRep = "Stock" Then
         MortalityType = CInt(Opt1)
         ParseOld = 1
         RepStkNum = 1
         NumRepStks = CInt(Opt2)
         If NumRepStks < 50 Then
            ParseOld = 1
            For Stk As Integer = 1 To NumStk
               ParseNew = InStr(ParseOld, Opt3, ",")
               If ParseNew = 0 Then
                  RepStks(Stk) = CInt(Opt3.Substring(ParseOld - 1, Opt3.Length - ParseOld + 1))
                  NumRepStks = Stk
                  Exit For
               Else
                  RepStks(Stk) = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
               End If
               ParseOld = ParseNew + 1
            Next
         Else
            For Stk As Integer = 1 To NumStk
               If CInt(Opt3.Substring(Stk - 1, 1)) = 1 Then
                  RepStks(RepStkNum) = Stk
                  RepStkNum += 1
               End If
            Next
         End If
         RepGrpName = Opt4
      End If
      If MortalityType = 7 Then
         OptionChinookBYAEQ = 2
         Call BYERReport()
         '- Put Brood Year Age 2 Array into Mortality Arrays then Reverse after Report
         BY = 2
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     If TStep = NumSteps Then
                        LandedCatch(Stk, Age, Fish, TStep) = 0
                        Shakers(Stk, Age, Fish, TStep) = 0
                        DropOff(Stk, Age, Fish, TStep) = 0
                        NonRetention(Stk, Age, Fish, TStep) = 0
                        Encounters(Stk, Age, Fish, TStep) = 0
                        MSFLandedCatch(Stk, Age, Fish, TStep) = 0
                        MSFShakers(Stk, Age, Fish, TStep) = 0
                        MSFDropOff(Stk, Age, Fish, TStep) = 0
                        MSFNonRetention(Stk, Age, Fish, TStep) = 0
                        MSFEncounters(Stk, Age, Fish, TStep) = 0
                     Else
                        LandedCatch(Stk, Age, Fish, TStep) = BYLandedCatch(BY, Stk, Age, Fish, TStep)
                        Shakers(Stk, Age, Fish, TStep) = BYShakers(BY, Stk, Age, Fish, TStep)
                        DropOff(Stk, Age, Fish, TStep) = BYDropOff(BY, Stk, Age, Fish, TStep)
                        NonRetention(Stk, Age, Fish, TStep) = BYNonRetention(BY, Stk, Age, Fish, TStep)
                        MSFLandedCatch(Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                        MSFShakers(Stk, Age, Fish, TStep) = BYMSFShakers(BY, Stk, Age, Fish, TStep)
                        MSFDropOff(Stk, Age, Fish, TStep) = BYMSFDropOff(BY, Stk, Age, Fish, TStep)
                        MSFNonRetention(Stk, Age, Fish, TStep) = BYMSFNonRetention(BY, Stk, Age, Fish, TStep)
                     End If
                  Next
               Next
            Next
         Next
      End If

      '============== FISHERY SUMMARY REPORT #1 ===================

      For Report = 1 To 5

         If Report > 1 And MortalityType >= 6 Then
            '- Write NewPage for Multiple Reports
            rw.WriteLine(Chr(12))
         End If

         '- Report Header Information for both CHINOOK and COHO
         If MortalityType = Report Or MortalityType >= 6 Then
            PrnLine = "Species:" & String.Format("{0,-7}", SpeciesName)
            PrnLine &= " FRAM-Version:" & String.Format("{0,-4}", FramVersion)
            PrnLine &= "  RunName: " & String.Format("{0,-27}", RunIDNameSelect)
            PrnLine &= "RunDate:" & RunIDRunTimeDateSelect.ToString
            rw.WriteLine(PrnLine)
            If TypeRep = "Fishery" Then
               PrnLine = "Report: Fishery Summary Report    "
               PrnLine &= "     Driver: " & String.Format("{0,-27}", ReportDriverName)
               PrnLine &= " RepDate:" & Now.ToString
            Else
               PrnLine = "Report: Stock Catch Summary Report"
               PrnLine &= "     Driver: " & String.Format("{0,-27}", ReportDriverName)
               PrnLine &= " RepDate:" & Now.ToString
            End If
            rw.WriteLine(PrnLine)
            If TypeRep = "Fishery" Then
               rw.WriteLine("")
            Else
               PrnLine = "Stock Group: " & RepGrpName
               rw.WriteLine(PrnLine)
            End If
         Else
            GoTo NextFishSumRep
         End If
         '- Title Line
         If MortalityType <= 6 Then
            Select Case Report
               Case 1
                  rw.WriteLine("LANDED CATCH BY FISHERY")
               Case 2
                  rw.WriteLine("SHAKER MORTALITY BY FISHERY")
               Case 3
                  rw.WriteLine("CNR (NON-RETENTION) MORTALITY BY FISHERY")
               Case 4
                  If SpeciesName = "COHO" Then
                     rw.WriteLine("CATCH + CNR MORTALITY BY FISHERY")
                  Else
                     rw.WriteLine("AEQ TOTAL MORTALITY BY FISHERY")
                  End If
               Case 5
                  rw.WriteLine("TOTAL MORTALITY (CATCH+SHAKER+LEGSHKR+CNR+DRPOFF) BY FISHERY")
            End Select
         ElseIf MortalityType = 7 Then
            Select Case Report
               Case 1
                  rw.WriteLine("BROOD YEAR LANDED CATCH BY FISHERY")
               Case 2
                  rw.WriteLine("BROOD YEAR SHAKER MORTALITY BY FISHERY")
               Case 3
                  rw.WriteLine("BROOD YEAR CNR (NON-RETENTION) MORTALITY BY FISHERY")
               Case 4
                  rw.WriteLine("BROOD YEAR AEQ TOTAL MORTALITY BY FISHERY")
               Case 5
                  rw.WriteLine("BROOD YEAR TOTAL MORTALITY  BY FISHERY")
            End Select
         End If

         If SpeciesName = "COHO" Then

            Dim TotCatch(NumFish, NumSteps + 1)
            Dim SelStk As Integer

            PrnLine = ("====================")
            For TStep = 1 To NumSteps + 1
               PrnLine &= ("===========")
            Next TStep
            rw.WriteLine(PrnLine)
            PrnLine = "       Fishery      "
            For TStep = 1 To NumSteps
               PrnLine &= String.Format("{0,11}", TimeStepName(TStep))
            Next
            PrnLine &= String.Format("{0,11}", "Total")
            rw.WriteLine(PrnLine)
            PrnLine = ("====================")
            For TStep = 1 To NumSteps + 1
               PrnLine &= ("===========")
            Next TStep
            rw.WriteLine(PrnLine)

            '- Sum TotCatch Arrays for COHO Fishery Summary Reports
            Age = 3
            Select Case Report
               Case 1
                  '- Landed Catch
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumSelStk1
                        Next
                        GoTo SkipSelStk1
                     End If
SumSelStk1:
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                        Next
                     Next
SkipSelStk1:
                  Next
               Case 2
                  '- Shakers
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumSelStk2
                        Next
                        GoTo SkipSelStk2
                     End If
SumSelStk2:
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        Next
                     Next
SkipSelStk2:
                  Next
               Case 3
                  '- Non-Retention
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumSelStk3
                        Next
                        GoTo SkipSelStk3
                     End If
SumSelStk3:
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        Next
                     Next
SkipSelStk3:
                  Next
               Case 4
                  '- Catch + CNR ???? (Old PFMC Style Report)
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumSelStk4
                        Next
                        GoTo SkipSelStk4
                     End If
SumSelStk4:
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        Next
                     Next
SkipSelStk4:
                  Next
               Case 5
                  '- Total Mortality
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumSelStk5
                        Next
                        GoTo SkipSelStk5
                     End If
SumSelStk5:
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        Next
                     Next
SkipSelStk5:
                  Next
            End Select

            '- Print Body of Report
            For Fish As Integer = 1 To NumFish
               PrnLine = String.Format("{0,20}", FisheryName(Fish))
               For TStep As Integer = 1 To NumSteps + 1
                  PrnLine &= String.Format("{0,11}", CLng(TotCatch(Fish, TStep)))
               Next
               rw.WriteLine(PrnLine)
            Next
            PrnLine = ("====================")
            For TStep As Integer = 1 To NumSteps + 1
               PrnLine &= ("===========")
            Next TStep
            rw.WriteLine(PrnLine)

            '=-=-=-=-=-=-=-=---------------------------------------=-=-=-=-=-
         ElseIf SpeciesName = "CHINOOK" Then

            Dim TotCatch(NumFish, NumSteps + 2) As Double
            Dim SelStk As Integer

            PrnLine = ("=========================")
            For TStep = 1 To NumSteps + 2
               PrnLine &= ("==============")
            Next TStep
            rw.WriteLine(PrnLine)
            PrnLine = "          Fishery        "
            For TStep = 1 To NumSteps
               PrnLine &= String.Format("{0,14}", TimeStepName(TStep))
            Next
            PrnLine &= String.Format("{0,14}", "GrandTot(1-4)")
            PrnLine &= String.Format("{0,14}", "SubTotal(2-4)")
            rw.WriteLine(PrnLine)
            PrnLine = ("=========================")
            For TStep = 1 To NumSteps + 2
               PrnLine &= ("==============")
            Next TStep
            rw.WriteLine(PrnLine)

            '- Sum TotCatch Arrays for CHINOOK Fishery Summary Reports
            Select Case Report
               Case 1
                  '- Landed Catch
                  For Stk As Integer = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumChinStk1
                        Next
                        GoTo SkipChinStk1
                     End If
SumChinStk1:
                     For Age As Integer = MinAge To MaxAge
                        For Fish As Integer = 1 To NumFish
                           For TStep As Integer = 1 To NumSteps
                              TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                              TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                              If TStep <> 1 Then
                                 TotCatch(Fish, NumSteps + 2) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                              End If
                           Next
                        Next
                     Next
SkipChinStk1:
                  Next
                  Jim = 1
               Case 2
                  '- Shakers
                  For Stk = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumChinStk2
                        Next
                        GoTo SkipChinStk2
                     End If
SumChinStk2:
                     For Age As Integer = MinAge To MaxAge
                        For Fish As Integer = 1 To NumFish
                           For TStep As Integer = 1 To NumSteps
                              TotCatch(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                              TotCatch(Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                              If TStep <> 1 Then
                                 TotCatch(Fish, NumSteps + 2) += Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                              End If
                           Next
                        Next
                     Next
SkipChinStk2:
                  Next
               Case 3
                  '- Non-Retention
                  For Stk = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumChinStk3
                        Next
                        GoTo SkipChinStk3
                     End If
SumChinStk3:
                     For Age As Integer = MinAge To MaxAge
                        For Fish As Integer = 1 To NumFish
                           For TStep As Integer = 1 To NumSteps
                              TotCatch(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              TotCatch(Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              If TStep <> 1 Then
                                 TotCatch(Fish, NumSteps + 2) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              End If
                           Next
                        Next
                     Next
SkipChinStk3:
                  Next
               Case 4
                  '- Total AEQ Mortality (Replaces Old PFMC Catch + CNR)
                  For Stk = 1 To NumStk
                     If Stk = 37 Then
                        Jim = 1
                     End If
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumChinStk4
                        Next
                        GoTo SkipChinStk4
                     End If
SumChinStk4:
                     For Age As Integer = MinAge To MaxAge
                        For Fish As Integer = 1 To NumFish
                           For TStep As Integer = 1 To NumSteps
                              If TerminalFisheryFlag(Fish, TStep) = Term Then
                                 TotCatch(Fish, TStep) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))
                                 TotCatch(Fish, NumSteps + 1) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))
                                 If TStep <> 1 Then
                                    TotCatch(Fish, NumSteps + 2) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep))
                                 End If
                              Else
                                 TotCatch(Fish, TStep) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                                 TotCatch(Fish, NumSteps + 1) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                                 'TempCatch = (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
                                 'If (Stk = 37 Or Stk = 38) And TempCatch <> 0 And Fish > 3 And Fish < 16 And TStep < 4 Then
                                 '   PrnLine = String.Format("{0,3}{1,4}{2,4}{3,4}", Stk.ToString, Age.ToString, Fish.ToString, TStep.ToString)
                                 '   PrnLine &= String.Format("{0,10}", LandedCatch(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", Shakers(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", LegalShakers(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", NonRetention(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", DropOff(Stk, Age, Fish, TStep).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", (TempCatch * AEQ(Stk, Age, TStep)).ToString("#####0.00"))
                                 '   PrnLine &= String.Format("{0,10}", AEQ(Stk, Age, TStep).ToString("##0.00000"))
                                 '   rw.WriteLine(PrnLine)
                                 'End If
                                 If TStep <> 1 Then
                                    TotCatch(Fish, NumSteps + 2) += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                                 End If
                              End If
                           Next
                        Next
                     Next
SkipChinStk4:
                  Next
               Case 5
                  '- Total Mortality
                  For Stk = 1 To NumStk
                     If TypeRep = "Stock" Then
                        For SelStk = 1 To NumRepStks
                           If Stk = RepStks(SelStk) Then GoTo SumChinStk5
                        Next
                        GoTo SkipChinStk5
                     End If
SumChinStk5:
                     For Age As Integer = MinAge To MaxAge
                        For Fish As Integer = 1 To NumFish
                           For TStep As Integer = 1 To NumSteps
                              TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              If TStep <> 1 Then
                                 TotCatch(Fish, NumSteps + 2) += LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                              End If
                           Next
                        Next
                     Next
SkipChinStk5:
                  Next
            End Select

            '- Print Body of Report
            For Fish As Integer = 1 To NumFish
               If FisheryTitle(Fish).Length > 25 Then
                  PrnLine = String.Format("{0,25}", FisheryTitle(Fish).Substring(0, 25))
               Else
                  PrnLine = String.Format("{0,25}", FisheryTitle(Fish))
               End If
               For TStep = 1 To NumSteps + 2
                  If TypeRep = "Stock" Then
                     PrnLine &= String.Format("{0,14}", CLng(TotCatch(Fish, TStep)))
                  Else
                     PrnLine &= String.Format("{0,14}", CLng(TotCatch(Fish, TStep) / ModelStockProportion(Fish)))
                  End If
               Next
               rw.WriteLine(PrnLine)
            Next
            PrnLine = ("=========================")
            For TStep = 1 To NumSteps + 2
               PrnLine &= ("==============")
            Next TStep
            rw.WriteLine(PrnLine)

         End If

NextFishSumRep:
      Next

      '- Reverse Mortality Array Values when Brood Style Report Selected
      If MortalityType = 7 Then
         BY = 2
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge + 1 To MaxAge
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps - 1
                     LandedCatch(Stk, Age, Fish, TStep) = BYLandedCatch(BY, Stk, BY, Fish, TStep)
                     Shakers(Stk, Age, Fish, TStep) = BYShakers(BY, Stk, BY, Fish, TStep)
                     DropOff(Stk, Age, Fish, TStep) = BYDropOff(BY, Stk, BY, Fish, TStep)
                     NonRetention(Stk, Age, Fish, TStep) = BYNonRetention(BY, Stk, BY, Fish, TStep)
                     MSFLandedCatch(Stk, Age, Fish, TStep) = BYMSFLandedCatch(BY, Stk, BY, Fish, TStep)
                     MSFShakers(Stk, Age, Fish, TStep) = BYMSFShakers(BY, Stk, BY, Fish, TStep)
                     MSFDropOff(Stk, Age, Fish, TStep) = BYMSFDropOff(BY, Stk, BY, Fish, TStep)
                     MSFNonRetention(Stk, Age, Fish, TStep) = BYMSFNonRetention(BY, Stk, BY, Fish, TStep)
                  Next
               Next
            Next
         Next
      End If

   End Sub

   Sub TerminalRunReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      Dim TermRun(MaxAge) As Double
      Dim StkEscAdlt, StkEscJack, TotAdltRun, TotTermRun As Double
      Dim CmbStk, SelStk, SelFish, ParseOld, ParseNew As Integer
      Dim BYCalcs As Boolean


      '- Parse Option Strings from Report Driver Table Fields until All Groups are Read
      ParseOld = 1
      '- Stock Group
      For Stk = 1 To NumStk
         ParseNew = InStr(ParseOld, Opt1, ",")
         If ParseNew = 0 Then
            RepStks(NumRepGrps, Stk) = CInt(Opt1.Substring(ParseOld - 1, Opt1.Length - ParseOld + 1))
            NumRepStks(NumRepGrps) = Stk
            Exit For
         Else
            RepStks(NumRepGrps, Stk) = CInt(Opt1.Substring(ParseOld - 1, ParseNew - ParseOld))
         End If
         ParseOld = ParseNew + 1
      Next
      ParseOld = 1
      '- Fishery Group .. Can be Zero for Escapement Only
      If Opt2 = "" Then
         NumRepFish(NumRepGrps) = 0
         GoTo SkipFishGroup
      End If
      For Fish As Integer = 1 To NumFish
         ParseNew = InStr(ParseOld, Opt2, ",")
         If ParseNew = 0 Then
            RepFish(NumRepGrps, Fish) = CInt(Opt2.Substring(ParseOld - 1, Opt2.Length - ParseOld + 1))
            NumRepFish(NumRepGrps) = Fish
            Exit For
         Else
            RepFish(NumRepGrps, Fish) = CInt(Opt2.Substring(ParseOld - 1, ParseNew - ParseOld))
         End If
         ParseOld = ParseNew + 1
      Next
SkipFishGroup:
      ParseOld = 1
      '- Terminal Time Steps
      For TStep = 1 To NumSteps
         ParseNew = InStr(ParseOld, Opt3, ",")
         If ParseNew = 0 Then
            RepTStep(NumRepGrps, TStep) = CInt(Opt3.Substring(ParseOld - 1, Opt3.Length - ParseOld + 1))
            Exit For
         Else
            RepTStep(NumRepGrps, TStep) = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
         End If
         ParseOld = ParseNew + 1
      Next
      RepGrpType(NumRepGrps) = Opt4
      RepGrpName(NumRepGrps) = Opt5

      '- Loop Back to Driver until all Term Run Records are Read
      If Opt6 = "GROUP" Then Exit Sub

      '- PRINT HEADER DATA
      PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If SpeciesName = "CHINOOK" Then
         If TermRunBYAEQ = True Then
            BYCalcs = True
            OptionChinookBYAEQ = 2
            Call BYERReport()
         End If
      End If

      If BYCalcs = False Then
         If ReportDriverName.Length > 20 Then
            PrnLine = "Report : Terminal Run Report              DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report : Terminal Run Report              DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "   RepDate:" & Now.ToString
      Else
         If ReportDriverName.Length > 20 Then
            PrnLine = "Report : BROOD YEAR Terminal Run Report   DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report : BROOD YEAR Terminal Run Report   DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "   RepDate:" & Now.ToString
      End If
      rw.WriteLine(PrnLine)

      PrnLine = "Title  : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      rw.WriteLine("")
      PrnLine = "========================="
      If SpeciesName = "COHO" Then
         For Age As Integer = MaxAge To MaxAge + 1
            PrnLine &= "==============="
         Next
         rw.WriteLine(PrnLine)
      ElseIf SpeciesName = "CHINOOK" Then
         For Age As Integer = MinAge To MaxAge + 4
            PrnLine &= "========"
         Next
         rw.WriteLine(PrnLine)
      End If

      '- PRINT ESCAPEMENT DATA ---
      If SpeciesName = "COHO" Then
         PrnLine = "                Stock     Terminal Run       Escapement"
         rw.WriteLine(PrnLine)
      ElseIf SpeciesName = "CHINOOK" Then
         PrnLine = "                         "
         For Age As Integer = MinAge To MaxAge
            PrnLine &= "     Age"
         Next
         PrnLine &= "    Adlt"
         PrnLine &= "   Total"
         PrnLine &= "    Jack"
         PrnLine &= "    Adlt"
         rw.WriteLine(PrnLine)
         PrnLine = "                Stock    "
         For Age As Integer = MinAge To MaxAge
            PrnLine &= String.Format("{0,8}", Age.ToString)
         Next
         PrnLine &= "     Run"
         PrnLine &= "     Run"
         PrnLine &= "     Esc"
         PrnLine &= "     Esc "
         rw.WriteLine(PrnLine)
      End If

      PrnLine = "========================="
      If SpeciesName = "COHO" Then
         For Age As Integer = MaxAge To MaxAge + 1
            PrnLine &= "==============="
         Next
         rw.WriteLine(PrnLine)
      ElseIf SpeciesName = "CHINOOK" Then
         For Age As Integer = MinAge To MaxAge + 4
            PrnLine &= "========"
         Next
         rw.WriteLine(PrnLine)
      End If
      rw.WriteLine("")

      For CmbStk = 1 To NumRepGrps
         ReDim TermRun(MaxAge)
         StkEscAdlt = 0
         StkEscJack = 0
         TotAdltRun = 0
         TotTermRun = 0
         If RepGrpName(CmbStk).Length > 25 Then
            PrnLine = String.Format("{0,25}", RepGrpName(CmbStk).Substring(0, 25))
         Else
            PrnLine = String.Format("{0,25}", RepGrpName(CmbStk))
         End If

         If BYCalcs = True Then  '-- Brood Year Style Report

            For TStep As Integer = 1 To NumSteps - 1
               For SelStk = 1 To NumRepStks(CmbStk)
                  Stk = RepStks(CmbStk, SelStk)
                  For Age As Integer = MinAge To MaxAge
                     TermRun(Age) += BYEscape(2, Stk, Age, TStep)
                     If Age <> 2 Then
                        TotAdltRun = TotAdltRun + BYEscape(2, Stk, Age, TStep)
                        StkEscAdlt = StkEscAdlt + BYEscape(2, Stk, Age, TStep)
                     Else
                        StkEscJack = StkEscJack + BYEscape(2, Stk, Age, TStep)
                     End If
                     TotTermRun = TotTermRun + BYEscape(2, Stk, Age, TStep)
                  Next Age
               Next SelStk
            Next TStep
            '--- Get Catch of All Stocks in Terminal Fishery ---
            If NumRepFish(CmbStk) = 0 Then GoTo SkipTermCatch
            If RepGrpType(CmbStk) = "TAA" Then
               For Stk As Integer = 1 To NumStk
                  For TStep As Integer = RepTStep(CmbStk, 1) To RepTStep(CmbStk, 2)
                     For Age As Integer = MinAge To MaxAge
                        For SelFish = 1 To NumRepFish(CmbStk)
                           Fish = RepFish(CmbStk, SelFish)
                           TermRun(Age) += (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep))
                           If Age <> 2 Then
                              TotAdltRun = TotAdltRun + (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                           End If
                           TotTermRun = TotTermRun + (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        Next SelFish
                     Next Age
                  Next TStep
               Next Stk
               '---- Get Catch of Local Stock in Terminal Fishery ---
            Else
               For SelStk = 1 To NumRepStks(CmbStk)
                  Stk = RepStks(CmbStk, SelStk)
                  For TStep As Integer = RepTStep(CmbStk, 1) To RepTStep(CmbStk, 2)
                     For Age As Integer = MinAge To MaxAge
                        For SelFish = 1 To NumRepFish(CmbStk)
                           Fish = RepFish(CmbStk, SelFish)
                           TermRun(Age) = TermRun(Age) + (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep))
                           If Age <> 2 Then
                              TotAdltRun = TotAdltRun + (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep))
                           End If
                           TotTermRun = TotTermRun + (BYLandedCatch(2, Stk, Age, Fish, TStep) + BYMSFLandedCatch(2, Stk, Age, Fish, TStep))
                        Next SelFish
                     Next Age
                  Next TStep
               Next SelStk
            End If
SkipTermCatch:
         Else  '--- Normal Terminal Run Report Style
            '- First Get Escapement
            For TStep = 1 To NumSteps
               For SelStk = 1 To NumRepStks(CmbStk)
                  Stk = RepStks(CmbStk, SelStk)
                  '--- Cowlitz and Willamette Springs Mature and Escape in both
                  '--- Time 1 and 4 so TermRunRep uses Time 1 Only
                  If NumStk < 50 Then
                     If TStep = 4 And (SelStk = 25 Or SelStk = 26) Then
                        GoTo NextCRSpr
                     End If
                  Else
                     If TStep = 4 And (SelStk = 49 Or SelStk = 50 Or SelStk = 51 Or SelStk = 52) Then
                        GoTo NextCRSpr
                     End If
                  End If
                  For Age As Integer = MinAge To MaxAge
                     TermRun(Age) = TermRun(Age) + Escape(Stk, Age, TStep)
                     If Age <> 2 Then
                        TotAdltRun = TotAdltRun + Escape(Stk, Age, TStep)
                        StkEscAdlt = StkEscAdlt + Escape(Stk, Age, TStep)
                     Else
                        StkEscJack = StkEscJack + Escape(Stk, Age, TStep)
                     End If
                     TotTermRun = TotTermRun + Escape(Stk, Age, TStep)
                  Next Age
NextCRSpr:
               Next
            Next
            '- Next Get Catch 
            If NumRepFish(CmbStk) = 0 Then GoTo SkipTermCatch2
            If RepGrpType(CmbStk) = "TAA" Then
               '- TAA = All Stocks in Terminal Fishery 
               For Stk As Integer = 1 To NumStk
                  For TStep As Integer = RepTStep(CmbStk, 1) To RepTStep(CmbStk, 2)
                     For Age As Integer = MinAge To MaxAge
                        For SelFish = 1 To NumRepFish(CmbStk)
                           Fish = RepFish(CmbStk, SelFish)
                           TermRun(Age) = TermRun(Age) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                           If Age <> 2 Then
                              TotAdltRun = TotAdltRun + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                           End If
                           TotTermRun = TotTermRun + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)) / ModelStockProportion(Fish)
                        Next SelFish
                     Next Age
                  Next TStep
               Next Stk
            Else
               '- ETRS = Catch of Local Stock Only in Terminal Fishery
               For SelStk = 1 To NumRepStks(CmbStk)
                  Stk = RepStks(CmbStk, SelStk)
                  For TStep As Integer = RepTStep(CmbStk, 1) To RepTStep(CmbStk, 2)
                     For Age As Integer = MinAge To MaxAge
                        For SelFish = 1 To NumRepFish(CmbStk)
                           Fish = RepFish(CmbStk, SelFish)
                           TermRun(Age) = TermRun(Age) + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                           If Age <> 2 Then
                              TotAdltRun = TotAdltRun + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                           End If
                           TotTermRun = TotTermRun + (LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep))
                        Next
                     Next
                  Next
               Next
            End If
         End If
SkipTermCatch2:
         '- Print Term Run Sizes
         If SpeciesName = "COHO" Then
            PrnLine &= String.Format("{0,14}", CLng(TotAdltRun).ToString)
            PrnLine &= String.Format("{0,14}", CLng(StkEscAdlt).ToString)
            rw.WriteLine(PrnLine)
         ElseIf SpeciesName = "CHINOOK" Then
            For Age As Integer = MinAge To MaxAge
               PrnLine &= String.Format("{0,8}", (CLng(TermRun(Age)).ToString))
            Next Age
            PrnLine &= String.Format("{0,8}", (CLng(TotAdltRun).ToString))
            PrnLine &= String.Format("{0,8}", (CLng(TotTermRun).ToString))
            PrnLine &= String.Format("{0,8}", (CLng(StkEscJack).ToString))
            PrnLine &= String.Format("{0,8}", (CLng(StkEscAdlt).ToString))
            rw.WriteLine(PrnLine)
         End If
      Next CmbStk

      rw.WriteLine("")
      PrnLine = "========================="
      If SpeciesName = "COHO" Then
         For Age As Integer = MaxAge To MaxAge + 1
            PrnLine &= "==============="
         Next Age
         rw.WriteLine(PrnLine)
      ElseIf SpeciesName = "CHINOOK" Then
         For Age As Integer = MinAge To MaxAge + 4
            PrnLine &= "========"
         Next Age
         rw.WriteLine(PrnLine)
      End If

      '- Print Area Names Used in Terminal Run Designation 
      rw.WriteLine("")
      PrnLine = "TERMINAL RUN FISHERIES FOR COMBINED STOCKS LISTED ABOVE"
      rw.WriteLine(PrnLine)
      PrnLine = "          - Time Periods Listed in Parentheses"
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      For CmbStk = 1 To NumRepGrps
         PrnLine = String.Format("{0,-25}", RepGrpName(CmbStk))
         rw.WriteLine(PrnLine)
         If (RepGrpType(CmbStk) = "TAA") Then
            PrnLine = "    Flag=ON Catch of ALL Stocks with True-To-Model Adjustment for Adlt & Total Run"
            rw.WriteLine(PrnLine)
         Else
            PrnLine = "    Flag=OFF Catch of LOCAL Stocks WITHOUT True-To-Model Adjustment for Adlt & Total Run"
            rw.WriteLine(PrnLine)
         End If
         PrnLine = ""
         For Fish As Integer = 1 To NumRepFish(CmbStk)
            If PrnLine = "" Then
               PrnLine = String.Format("  {0,3}", Fish.ToString) & "-"
            Else
               PrnLine &= String.Format(" {0,3}", Fish.ToString) & "-"
            End If
            If FisheryName(RepFish(CmbStk, Fish)).Length > 10 Then
               PrnLine &= String.Format("{0,10}", FisheryName(RepFish(CmbStk, Fish)).Substring(0, 10))
            Else
               PrnLine &= String.Format("{0,10}", FisheryName(RepFish(CmbStk, Fish)))
            End If
            PrnLine &= String.Format("({0,1}", RepTStep(CmbStk, 1).ToString) & "-" & _
            String.Format("{0,1})", RepTStep(CmbStk, 2).ToString)
            '- Print Four Per Line
            If ((Fish Mod 4) = 0) Then
               rw.WriteLine(PrnLine)
               PrnLine = ""
            End If
         Next
         If PrnLine <> "" Then rw.WriteLine(PrnLine)
         If NumRepFish(CmbStk) = 0 Then
            PrnLine = "    No Fisheries - ESCAPEMENT Only"
            rw.WriteLine(PrnLine)
         End If
      Next CmbStk

   End Sub


   Sub MortAgeReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      '------------------------- MORTALITY BY AGE REPORT #5 ---
      'Dim ParseOld, ParseNew, NumRepStks, NumRepFish, SelFish, SelStk As Integer
      Dim RepTotMort, RepMort As Double
      Dim RepStks(NumStk), RepFish(NumFish) As Integer

      '- Parse Option Strings from Report Driver Table Fields until All Groups are Read
      MortalityType = CInt(Opt1)
      '- Check if Opt 2 string is same as Current NumFish
      If Opt2.Length <> NumFish Then
         MsgBox("DRV Error - Number of Fisheries different than Current Base Period " & vbCrLf & _
                "the DRV file and Base Period File are mismatched for MORTBYAGE Report" & vbCrLf & _
                "... Choose Apprpriate DRV", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      '- Check if Opt 3 string is same as Current NumStk
      If Opt3.Length <> NumStk Then
         MsgBox("DRV Error - Number of Stocks different than Current Base Period " & vbCrLf & _
                "the DRV file and Base Period File are mismatched for MORTBYAGE Report" & vbCrLf & _
                "... Choose Apprpriate DRV", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      For Fish As Integer = 1 To NumFish
         RepFish(Fish) = CInt(Opt2.Substring(Fish - 1, 1))
      Next
      For Stk As Integer = 1 To NumStk
         RepStks(Stk) = CInt(Opt3.Substring(Stk - 1, 1))
      Next

      'PRINT CATCH BY AGE AND TYPE OVER TIME PERIODS

      For Stk = 1 To NumStk
         If RepStks(Stk) = 0 Then GoTo NextMortAgeStk
         'Stk = RepStks(SelStk)
         '- PRINT HEADER DATA
         PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
         PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
         If RunIDNameSelect.Length > 25 Then
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
         Else
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
         End If
         PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
         rw.WriteLine(PrnLine)
         If ReportDriverName.Length > 20 Then
            PrnLine = "Report : Stock Mortality by Age Report    DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report : Stock Mortality by Age Report    DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "     RepDate:" & Now.ToString
         rw.WriteLine(PrnLine)
         PrnLine = "Stock  : " & StockTitle(Stk)
         rw.WriteLine(PrnLine)
         '-------------------------- PRINT LINE AT TOP OF TABLE COLUMN HEADERS
         If MortalityType = 1 Then
            PrnLine = "LANDED CATCH BY FISHERY, TIME, AND AGE"
         ElseIf MortalityType = 2 Then
            PrnLine = "TOTAL MORTALITY BY FISHERY, TIME, AND AGE"
         ElseIf MortalityType = 3 Then
            PrnLine = "AEQ TOTAL MORTALITY BY FISHERY, TIME, AND AGE"
         End If
         rw.WriteLine(PrnLine)

         If SpeciesName = "COHO" Then

            PrnLine = "============="
            For TStep As Integer = 1 To NumSteps + 1
               PrnLine &= "========"
            Next TStep
            rw.WriteLine(PrnLine)
            'PrnLine = "              ---------- Age 3 Data Only ----------"
            'rw.WriteLine(PrnLine)
            'rw.WriteLine("")
            PrnLine = "Fishery      "
            For TStep As Integer = 1 To NumSteps
               'PrnLine = Format(Mid(TimeStepName(TStep), 3), "@@@@@@")
               PrnLine &= String.Format("{0,8}", TimeStepName(TStep))
            Next TStep
            PrnLine &= "   Total"
            rw.WriteLine(PrnLine)
            PrnLine = "============="
            For TStep = 1 To NumSteps + 1
               PrnLine &= "========"
            Next TStep
            rw.WriteLine(PrnLine)
            For Fish As Integer = 1 To NumFish
               If RepFish(Fish) = 0 Then GoTo NextMortAgeFish1
               RepTotMort = 0
               PrnLine = String.Format("{0,-13}", FisheryName(Fish))
               For TStep As Integer = 1 To NumSteps
                  If MortalityType = 1 Then
                     '- Landed Catch
                     PrnLine &= String.Format("{0,8}", CLng(LandedCatch(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep)))
                     RepTotMort += LandedCatch(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep)
                  ElseIf MortalityType > 1 Then
                     '- Total Mortality
                     RepMort = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     PrnLine &= String.Format("{0,8}", CLng(RepMort))
                     RepTotMort += RepMort
                  End If
               Next TStep
               PrnLine &= String.Format("{0,8}", CLng(RepTotMort))
               rw.WriteLine(PrnLine)
NextMortAgeFish1:
            Next
            PrnLine = "============="
            For TStep As Integer = 1 To NumSteps + 1
               PrnLine &= "========"
            Next TStep
            rw.WriteLine(PrnLine)
            rw.WriteLine("")

         ElseIf SpeciesName = "CHINOOK" Then

            PrnLine = "============="
            For Age As Integer = MinAge To MaxAge
               PrnLine &= "=============================="
            Next
            rw.WriteLine(PrnLine)
            'rw.WriteLine("")
            PrnLine = "             "
            For Age As Integer = MinAge To MaxAge
               '---- Age 2 Time Steps ----
               PrnLine &= "     --- Age "
               PrnLine &= Age.ToString
               PrnLine &= " Time Steps ----"
            Next Age
            rw.WriteLine(PrnLine)
            PrnLine = " Fishery     "
            For Age As Integer = MinAge To MaxAge
               For TStep As Integer = 1 To NumSteps
                  PrnLine &= String.Format("{0,6}", TStep.ToString)
               Next TStep
               PrnLine &= " Total"
            Next Age
            rw.WriteLine(PrnLine)
            PrnLine = "============="
            For Age As Integer = MinAge To MaxAge
               PrnLine &= "=============================="
            Next
            rw.WriteLine(PrnLine)
            For Fish As Integer = 1 To NumFish
               If RepFish(Fish) = 0 Then GoTo NextMortAgeFish2
               PrnLine = String.Format("{0,-13}", FisheryName(Fish))
               For Age As Integer = MinAge To MaxAge
                  RepTotMort = 0
                  For TStep As Integer = 1 To NumSteps
                     If MortalityType = 1 Then
                        '- Landed Catch
                        PrnLine &= String.Format("{0,6}", CLng(LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)))
                        RepTotMort += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                     ElseIf MortalityType = 2 Then
                        '- Total Mortality
                        RepMort = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                        PrnLine &= String.Format("{0,6}", CLng(RepMort))
                        RepTotMort += RepMort
                     ElseIf MortalityType = 3 Then
                        '- AEQ Total Mortality - Terminal Fisheries AEQ = 1
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                           RepMort = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                        Else
                           RepMort = AEQ(Stk, Age, TStep) * (LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep))
                        End If
                        PrnLine &= String.Format("{0,6}", CLng(RepMort))
                        RepTotMort += RepMort
                     End If
                  Next TStep
                  PrnLine &= String.Format("{0,6}", CLng(RepTotMort))
               Next
               rw.WriteLine(PrnLine)
NextMortAgeFish2:
            Next
            PrnLine = "============="
            For Age As Integer = MinAge To MaxAge
               PrnLine &= "=============================="
            Next
            rw.WriteLine(PrnLine)
            rw.WriteLine("")
         End If
NextMortAgeStk:
      Next

   End Sub

   Sub FisheryScalerReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      '---------------------- EXPLOITATION RATE SCALE REPORT #6 ---
      Dim TempScaler As Double

      '- PRINT HEADER DATA
      PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report : Fishery Scaler Report            DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report : Fishery Scaler Report            DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "     RepDate:" & RunIDRunTimeDateSelect.ToShortTimeString
      rw.WriteLine(PrnLine)

      PrnLine = "Title  : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      PrnLine = "EXPLOITATION RATE SCALE FACTORS BY FISHERY"
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      rw.WriteLine("")
      PrnLine = "========================="
      For TStep = 1 To NumSteps
         PrnLine &= "==============="
      Next
      rw.WriteLine(PrnLine)
      PrnLine = "                         "
      For TStep As Integer = 1 To NumSteps
         PrnLine &= String.Format("{0,15}", TimeStepName(TStep).ToString)
      Next
      rw.WriteLine(PrnLine)
      PrnLine = "========================="
      For TStep As Integer = 1 To NumSteps
         PrnLine &= "==============="
      Next
      rw.WriteLine(PrnLine)

      For Fish As Integer = 1 To NumFish
         If FisheryTitle(Fish).Length > 25 Then
            PrnLine = String.Format("{0,-25}", FisheryTitle(Fish).Substring(0, 25))
         Else
            PrnLine = String.Format("{0,-25}", FisheryTitle(Fish).ToString)
         End If
         For TStep As Integer = 1 To NumSteps
            '- Special Case for Chinook Base Period Numbers
            If InStr(ReportDriverName, "BP-ReCalc") <> 0 Then
               TempScaler = FisheryScaler(Fish, TStep) * ((TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep)) / TotalLandedCatch(Fish, TStep))
               PrnLine &= String.Format("{0,15}", TempScaler.ToString("###0.0000"))
            Else
               '- Normal Case Printing
               PrnLine &= String.Format("{0,15}", FisheryScaler(Fish, TStep).ToString("###0.0000"))
            End If
         Next TStep
         rw.WriteLine(PrnLine)
      Next Fish

      PrnLine = "========================="
      For TStep As Integer = 1 To NumSteps
         PrnLine &= "==============="
      Next
      rw.WriteLine(PrnLine)

   End Sub

   Sub StockSummaryReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      '---------------------------------- STOCK SUMMARY REPORT ---
      Dim StkTotCat(NumStk, NumFish + 1), FishTotCat(NumFish + 1), TempVal As Double
      Dim Page, LastPage, BegStk, EndStk As Integer

      '------------ SUM CATCH OVER AGES AND TIMES STEPS ------------

      BegStk = 1
      EndStk = 10
      For Fish As Integer = 1 To NumFish
         For Stk As Integer = 1 To NumStk
            For TStep As Integer = 1 To NumSteps
               If TStep = 1 And SpeciesName = "CHINOOK" Then GoTo SkipStep1
               For Age As Integer = MinAge To MaxAge
                  StkTotCat(Stk, Fish) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  FishTotCat(Fish) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  StkTotCat(Stk, NumFish + 1) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  FishTotCat(NumFish + 1) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
               Next
SkipStep1:
            Next
         Next
      Next

      'STOCK SUMMARY REPORT PRINT LOOP

      '- PRINT HEADER DATA
      PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate: " + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report : Stock Summary Report             DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report : Stock Summary Report             DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate: " & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title  : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      If SpeciesName = "CHINOOK" Then
         PrnLine = "Total Mortality by Fishery by Stock for Time Periods 2-4 and ALL Ages"
      Else
         PrnLine = "Total Mortality by Fishery by Stock for ALL Time Periods and ALL Ages"
      End If
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      'PRINT TABLE HEADINGS

      LastPage = 0
      For Page = 1 To 100
         BegStk = Page * 10 - 9
         EndStk = Page * 10
         If EndStk >= NumStk Then
            EndStk = NumStk
            LastPage = 1
         End If
         If BegStk > NumStk Then Exit For

         If Page <> 1 Then
            PrnLine = Chr(12)
            rw.WriteLine(PrnLine)
            If SpeciesName = "COHO" Then
               PrnLine = "Total Mortality by Fishery by Stock for ALL Time Periods and ALL Ages    Page" + String.Format("{0,3}", Page.ToString("#0"))
            Else
               PrnLine = "Total Mortality by Fishery by Stock for Time Periods 2-4 and ALL Ages    Page" + String.Format("{0,3}", Page.ToString("#0"))
            End If
         End If
         PrnLine = "==========="
         For TStep = BegStk To EndStk
            PrnLine &= "=========="
         Next
         If LastPage = 1 Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)
         PrnLine = "  Fishery  "
         For Stk = BegStk To EndStk
            If StockName(Stk).Length > 10 Then
               PrnLine &= String.Format("{0,10}", StockName(Stk).Substring(0, 10))
            Else
               PrnLine &= String.Format("{0,10}", StockName(Stk))
            End If
         Next
         If LastPage = 1 Then PrnLine &= "     Total"
         rw.WriteLine(PrnLine)
         PrnLine = "==========="
         For TStep As Integer = BegStk To EndStk
            PrnLine &= "=========="
         Next
         If LastPage = 1 Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)
         For Fish As Integer = 1 To NumFish + 1
            If Fish = NumFish + 1 Then
               PrnLine = "-- TOTAL -"
            Else
               If FisheryName(Fish).Length > 10 Then
                  PrnLine = String.Format("{0,-10}", FisheryName(Fish).Substring(0, 10))
               Else
                  PrnLine = String.Format("{0,-10}", FisheryName(Fish))
               End If
            End If
            PrnLine &= " "
            For Stk As Integer = BegStk To EndStk
               TempVal = CLng(StkTotCat(Stk, Fish))
               PrnLine &= String.Format("{0,10}", TempVal.ToString("######0"))
            Next
            If LastPage = 1 Then
               TempVal = CLng(FishTotCat(Fish))
               PrnLine &= String.Format("{0,10}", TempVal.ToString("######0"))
            End If
            rw.WriteLine(PrnLine)
         Next
         PrnLine = "==========="
         For TStep As Integer = BegStk To EndStk
            PrnLine &= "=========="
         Next
         If LastPage = 1 Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)
      Next

      '   '-------------------------- Percent Mortality Report ----
      '   For Year() = 1 To NumYears
      '      BegStk = 1
      '      EndStk = 10
      'Print #3, Chr(12)
      'Print #3, "Species: " + Species$; Tab(20); "Version#:" + VersNumb$; Tab(44); "CMD File: " + CMDFile$; Tab(80); "Date: " + Date$
      'Print #3, "Report : Stock Summary Report"; Tab(44); "DRV File: " + DRVFile$; Tab(80); "Time: " + Time$
      'Print #3, "Title  : ", CmdTitle$
      'Print #3,
      '      '   If Species$ = "COHO" Then
      '      '      Print #3, "Percent Catch by Fishery by Stock for ALL Time Periods and ALL Ages"
      '      '   Else
      '   Print #3, "Percent Total Mortality by Fishery by Stock for Time Periods 2-4 and ALL Ages"
      '      '   End If

      '      LastPage = 0
      '      For Page = 1 To 100
      '         BegStk = Page * 10 - 9
      '         EndStk = Page * 10
      '         If EndStk >= NumStk Then
      '            EndStk = NumStk
      '            LastPage = 1
      '         End If
      '         If BegStk > NumStk Then Exit For

      '         If Page <> 1 Then
      '      Print #3, Chr(12)
      '            If Species$ = "CHINOOK" Then
      '         Print #3, "Percent Total Mortality by Fishery by Stock for Time Periods 2-4 and ALL Ages    Page" + Format(Page, " ##")
      '            Else
      '         Print #3, "Percent Catch by Fishery by Stock for ALL Time Periods and ALL Ages    Page" + Format(Page, " ##")
      '            End If
      '         End If
      '   Print #3, "=========================";
      '         For TStep = BegStk To EndStk
      '      Print #3, "==========";
      '         Next TStep
      '         '      If LastPage = 1 Then Print #3, "==========";
      '   Print #3,
      '   Print #3, "    Fishery              ";
      '         For N = BegStk To EndStk
      '      Print #3, Format(SmlStockName$(N), "@@@@@@@@@@");
      '         Next N
      '         '      If LastPage = 1 Then Print #3, "     Total";
      '   Print #3,
      '   Print #3, "=========================";
      '         For TStep = BegStk To EndStk
      '      Print #3, "==========";
      '         Next TStep
      '         '      If LastPage = 1 Then Print #3, "==========";
      '   Print #3,
      '         For fish = 1 To Numfish
      '      Print #3, Format(Left(FishName$(fish), 25), "@@@@@@@@@@@@@@@@@@@@@@@@@");
      '            For Stk = BegStk To EndStk
      '               If FishTotCat&(fish) <> 0 Then
      '                  If StkTotCat&(Stk, fish) = 0 Then
      '               Print #3, "    ------";
      '                  Else
      '                     ERVal1 = Str(Int(((StkTotCat&(Stk, fish) / FishTotCat&(fish)) * 100 * TrueToModel!(fish))))
      '                     ERVal2 = Int((((StkTotCat&(Stk, fish) / FishTotCat&(fish)) * 100 * TrueToModel!(fish)) - Int(((StkTotCat&(Stk, fish) / FishTotCat&(fish)) * 100 * TrueToModel!(fish)))) * 10)
      '               Print #3, Format(ERVal1, " @@@@@@.") + Format(ERVal2, "0") + "";
      '                     '                  Print #3, Format((StkTotCat&(Stk, fish) / FishTotCat&(fish)) * 100, "  0.000000");
      '                  End If
      '               Else
      '            Print #3, "    ------";
      '               End If
      '            Next Stk
      '            '         If LastPage = 1 Then Print #3, Format(Str(CLng(FishTotCat&(fish))), " @@@@@@@@@");
      '      Print #3,
      '         Next fish
      '   Print #3, "=========================";
      '         For TStep = BegStk To EndStk
      '      Print #3, "==========";
      '         Next TStep
      '         '      If LastPage = 1 Then Print #3, "==========";
      '   Print #3,
      '      Next Page

      '------------
   End Sub

   Sub PopulationStatisticsReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      Dim PopStat(MaxAge - 1, NumStk, NumSteps, 5) As Double
      Dim TotStat(MaxAge - 1, 4) As Double
      Dim BegStk, EndStk, Page, CohortType As Integer
      Dim RowLabel(5) As String

      RowLabel(1) = "Starting Cohort"
      RowLabel(2) = "After Nat. Mort"
      RowLabel(3) = "After PreTermnl"
      RowLabel(4) = "Mature Cohort  "
      RowLabel(5) = "Escapement     "

      '- READ POPULATION AND ESCAPEMENT NUMBERS INTO ARRAY

      '- Cohort Sizes
      For TStep As Integer = 1 To NumSteps
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               PopStat(Age - 1, Stk, TStep, 0) = Cohort(Stk, Age, 0, TStep) '- Cohort
               PopStat(Age - 1, Stk, TStep, 4) = Cohort(Stk, Age, 1, TStep) '- Mature
               PopStat(Age - 1, Stk, TStep, 3) = Cohort(Stk, Age, 2, TStep) '- Mid-Calcs
               PopStat(Age - 1, Stk, TStep, 2) = Cohort(Stk, Age, 3, TStep) '- Working
               PopStat(Age - 1, Stk, TStep, 1) = Cohort(Stk, Age, 4, TStep) '- Starting
            Next
         Next
      Next

      '- ESCAPEMENT DATA 
      For TStep As Integer = 1 To NumSteps
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               PopStat(Age - 1, Stk, TStep, 5) = Escape(Stk, Age, TStep)
            Next
         Next
      Next

      '- PRINT LOOP by Species
      If SpeciesName = "COHO" Then
         Age = 3
         '- PRINT HEADER DATA
         PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
         PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
         If RunIDNameSelect.Length > 25 Then
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
         Else
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
         End If
         PrnLine &= " RunDate: " + RunIDRunTimeDateSelect.ToString
         rw.WriteLine(PrnLine)

         If ReportDriverName.Length > 20 Then
            PrnLine = "Report : Population Statistics Report     DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report : Population Statistics Report     DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "      RepDate: " & Now.ToString
         rw.WriteLine(PrnLine)

         PrnLine = "Title  : " & RunIDNameSelect
         rw.WriteLine(PrnLine)
         PrnLine = "POPULATION STATISTICS"
         rw.WriteLine(PrnLine)
         rw.WriteLine("")

         For Page = 1 To 100
            BegStk = Page * 10 - 9
            EndStk = Page * 10
            If EndStk >= NumStk Then EndStk = NumStk
            If BegStk > NumStk Then Exit For

            PrnLine = "==============="
            For Stk = BegStk To EndStk
               PrnLine &= "=========="
            Next
            rw.WriteLine(PrnLine)
            PrnLine = "               "
            For Stk = BegStk To EndStk
               PrnLine &= String.Format("{0,10}", StockName(Stk))
            Next
            rw.WriteLine(PrnLine)
            PrnLine = "==============="
            For Stk = BegStk To EndStk
               PrnLine &= "=========="
            Next
            rw.WriteLine(PrnLine)

            For TStep = 1 To NumSteps
               PrnLine = "Time Step "
               PrnLine &= String.Format("{0,2}", TStep.ToString)
               rw.WriteLine(PrnLine)
               For CohortType = 1 To 5
                  If SpeciesName = "COHO" And TStep < 5 And CohortType > 3 Then Exit For
                  PrnLine = String.Format("{0,9}", RowLabel(CohortType))
                  For Stk = BegStk To EndStk
                     PrnLine &= String.Format("{0,10}", (CLng(PopStat(Age - 1, Stk, TStep, CohortType)).ToString("######0")))
                  Next
                  rw.WriteLine(PrnLine)
               Next
               PrnLine = "==============="
               For Stk = BegStk To EndStk
                  PrnLine &= "=========="
               Next
               rw.WriteLine(PrnLine)
            Next
         Next

      ElseIf SpeciesName = "CHINOOK" Then
         '- CHINOOK Pop Stat
         '- PRINT HEADER DATA
         PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
         PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
         If RunIDNameSelect.Length > 25 Then
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
         Else
            PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
         End If
         PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
         rw.WriteLine(PrnLine)

         If ReportDriverName.Length > 20 Then
            PrnLine = "Report : Population Statistics Report     DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report : Population Statistics Report     DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "      RepDate:" & Now.ToString
         rw.WriteLine(PrnLine)

         PrnLine = "Title  : " & RunIDNameSelect
         rw.WriteLine(PrnLine)
         PrnLine = "POPULATION STATISTICS"
         rw.WriteLine(PrnLine)
         rw.WriteLine("")

         For Page = 1 To 100
            BegStk = Page * 2 - 1
            EndStk = Page * 2
            If EndStk >= NumStk Then EndStk = NumStk
            If BegStk > NumStk Then Exit For

            PrnLine = "==============="
            For Stk As Integer = BegStk To EndStk
               For Age As Integer = MinAge To MaxAge
                  PrnLine &= "=========="
               Next
            Next
            rw.WriteLine(PrnLine)
            PrnLine = "                    "
            For Stk As Integer = BegStk To EndStk
               PrnLine &= String.Format("{0,-41}", StockTitle(Stk))
            Next
            rw.WriteLine(PrnLine)
            PrnLine = "==============="
            For Stk As Integer = BegStk To EndStk
               For Age As Integer = MinAge To MaxAge
                  PrnLine &= "=========="
               Next
            Next
            rw.WriteLine(PrnLine)

            For TStep = 1 To NumSteps
               PrnLine = "Time Step "
               PrnLine &= String.Format("{0,2}     ", TStep.ToString)
               For Stk As Integer = BegStk To EndStk
                  For Age As Integer = MinAge To MaxAge
                     PrnLine &= String.Format("   Age {0,1}  ", Age.ToString)
                  Next
               Next
               rw.WriteLine(PrnLine)
               For CohortType = 1 To 5
                  PrnLine = String.Format("{0,9}", RowLabel(CohortType))
                  For Stk As Integer = BegStk To EndStk
                     For Age As Integer = MinAge To MaxAge
                        PrnLine &= String.Format("{0,10}", CLng(PopStat(Age - 1, Stk, TStep, CohortType)).ToString("######0"))
                     Next
                  Next
                  rw.WriteLine(PrnLine)
               Next
               PrnLine = "==============="
               For Stk As Integer = BegStk To EndStk
                  For Age As Integer = MinAge To MaxAge
                     PrnLine &= "=========="
                  Next
               Next
               rw.WriteLine(PrnLine)
            Next
         Next
      End If

   End Sub

   Sub SelectiveFisheryReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      '-------------- Selective Fishery Report ----------

      Dim StkTotal As Double
      Dim TotUEnc, TotUCat, TotUCNR, TotUNon, TotUShk, TotUSub As Double
      Dim TotMEnc, TotMCat, TotMCNR, TotMNon, TotMShk, TotMSub As Double
      Dim TempName As String
      Dim StrPos As Integer

      '- Determine if any Selective Fisheries are in Selected Recordset
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then GoTo SelectiveFisheryFound
         Next
      Next

      '- No Selective Fishery Found
      '- PRINT HEADER DATA
      PrnLine = "Species: " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report : Selective Fishery Report         DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report : Selective Fishery Report         DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title  : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "---- No Selective Fisheries Specified for this RecordSet --------"
      rw.WriteLine(PrnLine)
      Exit Sub

SelectiveFisheryFound:

      '- Print Separate Report for each Selective Fishery
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) > 6 Then
               '- PRINT HEADER DATA
               PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
               PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
               If RunIDNameSelect.Length > 25 Then
                  PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
               Else
                  PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
               End If
               PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
               rw.WriteLine(PrnLine)

               If ReportDriverName.Length > 20 Then
                  PrnLine = "Report  : Selective Fishery Report         DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
               Else
                  PrnLine = "Report  : Selective Fishery Report         DRV File: " & String.Format("{0,-20}", ReportDriverName)
               End If
               PrnLine &= "      RepDate:" & RunIDRunTimeDateSelect.ToShortTimeString
               rw.WriteLine(PrnLine)

               PrnLine = "Title   : " & RunIDNameSelect
               rw.WriteLine(PrnLine)
               rw.WriteLine("")
               PrnLine = "Fishery : " & FisheryTitle(Fish)
               rw.WriteLine(PrnLine)
               PrnLine = "TimeStep: " & TimeStepTitle(TStep)
               rw.WriteLine(PrnLine)
               rw.WriteLine("")

               If SpeciesName = "CHINOOK" Then
                  PrnLine = "  Stock         UnMark  UnMark  UnMark  UnMark  UnMark  Marked  Marked  Marked  Marked  Marked "
                  rw.WriteLine(PrnLine)
                  PrnLine = "  Name      Age Handled  Catch  NonRete Dropoff SubLegl Handled  Catch  NonRete Dropoff SubLegl"
                  rw.WriteLine(PrnLine)
                  PrnLine = " ---------- --- ------- ------- ------- ------- ------- ------- ------- ------- ------- -------"
                  rw.WriteLine(PrnLine)
               Else '- Coho
                  PrnLine = "  Stock                                  UnMark  UnMark  UnMark  UnMark  Marked  Marked  Marked  Marked "
                  rw.WriteLine(PrnLine)
                  PrnLine = "  Name                               Age Handled  Catch  NonRete Dropoff Handled  Catch  NonRete Dropoff"
                  rw.WriteLine(PrnLine)
                  PrnLine = " ----------------------------------- --- ------- ------- ------- ------- ------- ------- ------- -------"
                  rw.WriteLine(PrnLine)
               End If
               '- Get Data for each Stock/Age
               For Stk As Integer = 1 To NumStk Step 2
                  StkTotal = 0
                  For Age As Integer = MinAge To MaxAge
                     TotUEnc += MSFEncounters(Stk, Age, Fish, TStep)
                     TotUCat += MSFLandedCatch(Stk, Age, Fish, TStep)
                     TotUNon += MSFNonRetention(Stk, Age, Fish, TStep)
                     TotUShk += MSFDropOff(Stk, Age, Fish, TStep)
                     TotUSub += MSFShakers(Stk, Age, Fish, TStep)
                     TotMEnc += MSFEncounters(Stk + 1, Age, Fish, TStep)
                     TotMCat += MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                     TotMNon += MSFNonRetention(Stk + 1, Age, Fish, TStep)
                     TotMShk += MSFDropOff(Stk + 1, Age, Fish, TStep)
                     TotMSub += MSFShakers(Stk + 1, Age, Fish, TStep)

                     StkTotal = MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + _
                                MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep)

                     If CLng(StkTotal) <> 0 Then
                        If SpeciesName = "CHINOOK" Then
                           PrnLine = String.Format("{0,-10}", StockName(Stk))
                           PrnLine &= String.Format("{0,4}", Age.ToString)
                           PrnLine &= String.Format("{0,8}", CLng(MSFEncounters(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFLandedCatch(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFNonRetention(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFDropOff(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFShakers(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFEncounters(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFLandedCatch(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFNonRetention(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFDropOff(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFShakers(Stk + 1, Age, Fish, TStep).ToString))
                           rw.WriteLine(PrnLine)
                        Else
                           StrPos = InStr(StockTitle(Stk), "UnMarked")
                           TempName = StockTitle(Stk).Substring(0, StrPos - 1)
                           PrnLine = String.Format(" {0,-35}", TempName)
                           PrnLine &= String.Format("{0,4}", Age.ToString)
                           PrnLine &= String.Format("{0,8}", CLng(MSFEncounters(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFLandedCatch(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFNonRetention(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(DropOff(Stk, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(Encounters(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFLandedCatch(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFNonRetention(Stk + 1, Age, Fish, TStep).ToString))
                           PrnLine &= String.Format("{0,8}", CLng(MSFDropOff(Stk + 1, Age, Fish, TStep).ToString))
                           rw.WriteLine(PrnLine)
                        End If
                     End If
                  Next
               Next

               '--- Print Totals Line ---
               If SpeciesName = "CHINOOK" Then
                  PrnLine = "                ------- ------- ------- ------- ------- ------- ------- ------- ------- -------"
                  rw.WriteLine(PrnLine)
                  PrnLine = " FRAM Stocks  "
                  PrnLine &= String.Format("{0,8}", CLng(TotUEnc).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUCat).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUNon).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUShk).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUSub).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMEnc).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMCat).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMNon).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMShk).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMSub).ToString)
                  rw.WriteLine(PrnLine)
                  If ModelStockProportion(Fish) < 1 Then
                     PrnLine = " ALL Stocks   "
                     PrnLine &= String.Format("{0,8}", CLng(TotUEnc / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotUCat / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotUNon / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotUShk / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotUSub / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotMEnc / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotMCat / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotMNon / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotMShk / ModelStockProportion(Fish)).ToString)
                     PrnLine &= String.Format("{0,8}", CLng(TotMSub / ModelStockProportion(Fish)).ToString)
                     rw.WriteLine(PrnLine)
                  End If
               Else
                  PrnLine = "                                         ------- ------- ------- ------- ------- ------- ------- -------"
                  rw.WriteLine(PrnLine)
                  PrnLine = String.Format("{0,48}", CLng(TotUEnc).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUCat).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUNon).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotUShk).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMEnc).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMCat).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMNon).ToString)
                  PrnLine &= String.Format("{0,8}", CLng(TotMShk).ToString)
                  rw.WriteLine(PrnLine)
               End If
               TotUEnc = 0.0
               TotUCat = 0.0
               TotUCNR = 0.0
               TotUNon = 0.0
               TotUShk = 0.0
               TotUSub = 0.0
               TotMEnc = 0.0
               TotMCat = 0.0
               TotMCNR = 0.0
               TotMNon = 0.0
               TotMShk = 0.0
               TotMSub = 0.0
               rw.WriteLine("")
               rw.WriteLine("")
            End If
         Next
      Next

   End Sub

   Sub PSCCohoER(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      'added stocks 14-17 Sep 26, 2007 AHB

      Dim PSCGroup(17, 5) As Integer
      Dim PSCER(17, 9), StockTotal As Double
      Dim StkGroup, StkList As Integer
      Dim PSCStockName(17) As String

      PSCStockName(1) = "Skagit"
      PSCStockName(2) = "Stillaguamish"
      PSCStockName(3) = "Snohomish"
      PSCStockName(4) = "Hood Canal"
      PSCStockName(5) = "US Strait JDF"
      PSCStockName(6) = "Quillayute"
      PSCStockName(7) = "Hoh"
      PSCStockName(8) = "Queets"
      PSCStockName(9) = "Grays Harbor"
      PSCStockName(10) = "Lower Fraser"
      PSCStockName(11) = "Upper Fraser"
      PSCStockName(12) = "Georgia Mainland"
      PSCStockName(13) = "Georgia Vanc Isl"
      PSCStockName(14) = "Puyallup Hatch"
      PSCStockName(15) = "Skookum Hatch"
      PSCStockName(16) = "Deschutes Wild"
      PSCStockName(17) = "SPS Net Pens"

      '- Number of FRAM Stocks to Group
      PSCGroup(1, 0) = 2
      PSCGroup(2, 0) = 1
      PSCGroup(3, 0) = 1
      PSCGroup(4, 0) = 5
      PSCGroup(5, 0) = 3
      PSCGroup(6, 0) = 2
      PSCGroup(7, 0) = 1
      PSCGroup(8, 0) = 1
      PSCGroup(9, 0) = 3
      PSCGroup(10, 0) = 1
      '- Old Base Period had 256 Stocks ... Newer Base has changed Stock Numbers
      If NumStk = 256 Then
         PSCGroup(11, 0) = 2
      Else
         PSCGroup(11, 0) = 1
      End If
      PSCGroup(12, 0) = 1
      PSCGroup(13, 0) = 1
      PSCGroup(14, 0) = 1
      PSCGroup(15, 0) = 1
      PSCGroup(16, 0) = 1
      PSCGroup(17, 0) = 1
      '- FRAM Stock Codes (UnMarked Wild)
      PSCGroup(1, 1) = 17 '- Skagit Wild
      PSCGroup(1, 2) = 23 '- Baker Wild
      PSCGroup(2, 1) = 29 '- Stillaguamish Wild
      PSCGroup(3, 1) = 35 '- Snohomish Wild
      PSCGroup(4, 1) = 43 '- Pt Gamble Wild
      PSCGroup(4, 2) = 45 '- Area 12/12B Wild
      PSCGroup(4, 3) = 51 '- Area 12A Wild
      PSCGroup(4, 4) = 55 '- Area 12C/D Wild
      PSCGroup(4, 5) = 59 '- Skokomish Wild
      '   PSCGroup(5, 1) = 107 '- Dungeness Wild  Changed Jan 2004 for PSC Rep
      '   PSCGroup(5, 2) = 111 '- Elwha Wild
      PSCGroup(5, 1) = 115 '- East JDF Wild
      PSCGroup(5, 2) = 117 '- West JDF Wild
      PSCGroup(5, 3) = 121 '- Area 9 Misc Wild
      PSCGroup(6, 1) = 127 '- Quillayute Summer Wild
      PSCGroup(6, 2) = 131 '- Quillayute Fall Wild
      PSCGroup(7, 1) = 135 '- Hoh Wild
      PSCGroup(8, 1) = 139 '- Queets Wild
      PSCGroup(9, 1) = 149 '- Chehalis Wild
      PSCGroup(9, 2) = 153 '- Humptulips Wild
      PSCGroup(9, 3) = 157 '- Grays Harbor Misc Wild
      PSCGroup(10, 1) = 227 '- Lower Fraser Wild
      PSCGroup(11, 1) = 231 '- Upper Fraser Wild
      If NumStk = 256 Then
         PSCGroup(11, 2) = 235 '- Thompson Wild
      End If
      PSCGroup(12, 1) = 207 '- Georgia Str Mainland Wild
      PSCGroup(13, 1) = 211 '- Georgia Str Vanc Isl Wild
      PSCGroup(14, 1) = 83 '- Puyallup Hatchery
      PSCGroup(15, 1) = 5 '- Skookum Hatchery
      PSCGroup(16, 1) = 63 '- Deschutes Wild
      PSCGroup(17, 1) = 65 '- South Sound Net Pens

      '- Get ESCAPEMENT Data and put into array dimension-1
      Age = 3
      For TStep = 4 To 5
         For StkGroup = 1 To 17
            For StkList = 1 To PSCGroup(StkGroup, 0)
               Stk = PSCGroup(StkGroup, StkList)
               PSCER(StkGroup, 1) = PSCER(StkGroup, 1) + Escape(Stk, Age, TStep)
            Next
         Next
      Next

      '- Get Catch by US/Canada and put into appropriate array dimensions
      For Fish As Integer = 1 To NumFish
         For StkGroup = 1 To 17
            For StkList = 1 To PSCGroup(StkGroup, 0)
               Stk = PSCGroup(StkGroup, StkList)
               For TStep As Integer = 1 To NumSteps
                  If NumStk = 256 Then
                     If Fish > 166 And Fish < 202 Then '- Canadian Catch
                        PSCER(StkGroup, 3) = PSCER(StkGroup, 3) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     Else
                        PSCER(StkGroup, 2) = PSCER(StkGroup, 2) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                  Else
                     If Fish > 166 And Fish < 194 Then
                        '- Canadian Catch
                        PSCER(StkGroup, 6) = PSCER(StkGroup, 6) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish < 167 Or Fish > 193 Then
                        '- US Catch
                        PSCER(StkGroup, 2) = PSCER(StkGroup, 2) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish > 0 And Fish < 23 Or Fish > 32 And Fish < 44 Or Fish = 79 Or Fish > 193 And Fish < 199 Then 'US Ocean
                        PSCER(StkGroup, 3) = PSCER(StkGroup, 3) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish = 44 Or Fish > 79 And Fish < 167 Then
                        '- Puget Sound
                        PSCER(StkGroup, 4) = PSCER(StkGroup, 4) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish > 22 And Fish < 33 Or Fish > 44 And Fish < 79 Then
                        '- US Other
                        PSCER(StkGroup, 5) = PSCER(StkGroup, 5) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish > 170 And Fish < 176 Or Fish > 177 And Fish < 182 Or Fish > 186 And Fish < 189 Or Fish = 190 Then 'BC Ocean
                        PSCER(StkGroup, 7) = PSCER(StkGroup, 7) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish = 176 Or Fish = 183 Or Fish > 190 And Fish < 193 Then
                        '- Georgia Strait
                        PSCER(StkGroup, 8) = PSCER(StkGroup, 8) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     If Fish > 166 And Fish < 171 Or Fish = 177 Or Fish = 182 Or Fish > 183 And Fish < 187 Or Fish = 189 Or Fish = 193 Then
                        '- BC Other
                        PSCER(StkGroup, 9) = PSCER(StkGroup, 9) + LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                  End If
               Next
            Next
         Next
      Next

      '- PRINT HEADER DATA
      PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report  :PSC Coho Exploitation Rate Report DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report  :PSC Coho Exploitation Rate Report DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title   : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "     Stock         U.S.ER     US Ocean   US PS      U.S.Other  Cnd ER     B.C.Ocean  Geo.St.    B.C.Other   Total ER"
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      For Stk = 1 To 17
         PrnLine = String.Format("{0,-16}", PSCStockName(Stk))
         StockTotal = (PSCER(Stk, 1) + PSCER(Stk, 2) + PSCER(Stk, 6))
         If StockTotal = 0 Then
            PrnLine &= "  0.000000  0.000000  ----------  ---------- ---------- ----------"
            rw.WriteLine(PrnLine)
         Else
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 2) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 3) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 4) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 5) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 6) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 7) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 8) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", (PSCER(Stk, 9) / StockTotal).ToString("0.000000"))
            PrnLine &= String.Format("{0,11}", ((PSCER(Stk, 2) + PSCER(Stk, 6)) / StockTotal).ToString("0.000000"))
            rw.WriteLine(PrnLine)
         End If
      Next Stk

   End Sub

   Sub CohoStockER(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      Dim PrnData(NumFish + 2, NumSteps + 1) As Double
      Dim TotalMort(NumFish + 2, NumSteps + 1) As Double
      Dim TotalStockMort, TotalCohort As Double

      '- This report now does only one stock per report selection
      Stk = CInt(Opt1)

      Age = 3
      '- Add Escapement to Print Matrix ---
      For TStep = 1 To NumSteps
         PrnData(NumFish + 2, TStep) += Escape(Stk, Age, TStep)
         PrnData(NumFish + 2, NumSteps + 1) += Escape(Stk, Age, TStep)
      Next

      '- Total Fishery Mortality by Fishery
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            TotalMort(Fish, TStep) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotalMort(Fish, NumSteps + 1) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotalMort(NumFish + 1, TStep) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotalMort(NumFish + 1, NumSteps + 1) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
         Next
      Next

      '- Stock Total Mortality 
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            PrnData(Fish, TStep) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
            PrnData(NumFish + 1, TStep) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
            PrnData(Fish, NumSteps + 1) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
            PrnData(NumFish + 1, NumSteps + 1) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
         Next
      Next
      TotalStockMort = PrnData(NumFish + 1, NumSteps + 1)
      TotalCohort = TotalStockMort + PrnData(NumFish + 2, NumSteps + 1)

      '- PRINT HEADER DATA
      PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report  : Coho Exploitation Rate Report    DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report  : Coho Exploitation Rate Report    DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title   : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "Stock   : " & StockTitle(Stk)
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      PrnLine = "ESTIMATED TOTAL MORTALITY EXPLOITATION RATE BY FISHERY-TIME STRATA"
      rw.WriteLine(PrnLine)
      PrnLine = "Distribution of Exploitation Rate (Percent)"
      rw.WriteLine(PrnLine)
      PrnLine = "========================================================================="
      rw.WriteLine(PrnLine)
      PrnLine = " Fishery       "
      For TStep = 1 To NumSteps
         PrnLine += String.Format("{0,10}", TimeStepName(TStep))
      Next
      PrnLine += "   Total "
      rw.WriteLine(PrnLine)

      PrnLine = "========================================================================="
      rw.WriteLine(PrnLine)
      PrnLine = " (-----) No Fishery    (*****) No Stock Impact in Fishery"
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      '- Print Coho ER Values
      For Fish As Integer = 1 To NumFish + 1
         If PrnData(Fish, NumSteps + 1) = 0 Then GoTo NextFTMFish
         If Fish <= NumFish Then
            PrnLine = String.Format("{0,-13}", FisheryName(Fish))
         Else
            rw.WriteLine("")
            PrnLine = "Monthly Total"
         End If
         For TStep As Integer = 1 To NumSteps + 1
            If TotalMort(Fish, TStep) Then
               If PrnData(Fish, TStep) Then
                  PrnLine += String.Format("{0,10}", (PrnData(Fish, TStep) * 100.0! / TotalCohort).ToString("##0.0000"))
               Else
                  PrnLine += "     *****"
               End If
            Else
               PrnLine += "     -----"
            End If
         Next
         rw.WriteLine(PrnLine)
NextFTMFish:
      Next
      rw.WriteLine("")
      PrnLine = "Escapement   "
      For TStep = 1 To NumSteps + 1
         If PrnData(NumFish + 2, TStep) Then
            PrnLine += String.Format("{0,10}", CLng(PrnData(NumFish + 2, TStep)).ToString)
         Else
            PrnLine += "         0"
         End If
      Next
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "TotalMort    "
      For TStep As Integer = 1 To NumSteps + 1
         If PrnData(NumFish + 1, TStep) Then
            PrnLine += String.Format("{0,10}", CLng(PrnData(NumFish + 1, TStep)).ToString)
         Else
            PrnLine += "         0"
         End If
      Next
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

   End Sub

   Sub ChinookStockER(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)

      Dim TotStkMort(NumFish + 2, MaxAge, NumSteps + 1) As Double
      Dim TotAEQMort(NumFish + 2, MaxAge, NumSteps + 1) As Double
      Dim TotFisheryMort(NumFish + 2, NumSteps + 1) As Double
      Dim SumTotStkMort, SumTotAEQMort, TotalAEQCohort, SumTimeStep, SumStkFishMort As Double
      Dim AnyAgeMorts As Integer

      Stk = CInt(Opt1)

      '- Escapement
      For TStep As Integer = 1 To NumSteps
         For Age As Integer = MinAge To MaxAge
            TotStkMort(NumFish + 2, Age, TStep) += Escape(Stk, Age, TStep)
            TotStkMort(NumFish + 2, Age, NumSteps + 1) += Escape(Stk, Age, TStep)
            TotAEQMort(NumFish + 2, Age, TStep) += Escape(Stk, Age, TStep)
            TotAEQMort(NumFish + 2, Age, NumSteps + 1) += Escape(Stk, Age, TStep)
         Next
      Next

      '- Total Mortality by Fishery
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            TotFisheryMort(Fish, TStep) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotFisheryMort(Fish, NumSteps + 1) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotFisheryMort(NumFish + 1, TStep) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
            TotFisheryMort(NumFish + 1, NumSteps + 1) += TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
         Next TStep
      Next Fish

      '- Stock Total Mortality 
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            For Age As Integer = MinAge To MaxAge
               TotStkMort(Fish, Age, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
               TotStkMort(NumFish + 1, Age, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
               TotStkMort(Fish, Age, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
               TotStkMort(NumFish + 1, Age, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
               '- Check for Terminal Fishery Flag for AEQ Mortality
               If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                  TotAEQMort(Fish, Age, TStep) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  TotAEQMort(NumFish + 1, Age, TStep) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  TotAEQMort(Fish, Age, NumSteps + 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  TotAEQMort(NumFish + 1, Age, NumSteps + 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
               Else
                  TotAEQMort(Fish, Age, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                  TotAEQMort(NumFish + 1, Age, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                  TotAEQMort(Fish, Age, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                  TotAEQMort(NumFish + 1, Age, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
               End If
            Next
         Next
      Next

      '- Sum Stock Total Mortality and AEQ Mortality
      SumTotStkMort = 0
      For Age As Integer = MinAge To MaxAge
         SumTotStkMort += TotStkMort(NumFish + 1, Age, NumSteps + 1)
         SumTotAEQMort += TotAEQMort(NumFish + 1, Age, NumSteps + 1)
      Next
      TotalAEQCohort = SumTotAEQMort
      For Age As Integer = MinAge To MaxAge
         TotalAEQCohort = TotalAEQCohort + TotStkMort(NumFish + 2, Age, NumSteps + 1)
      Next Age

      '- Total Mortality Exploitation Rate Distribution 

      '- PRINT HEADER DATA
      PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,15}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  RUN Name: " & String.Format("{0,-25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  RUN Name: " & String.Format("{0,-25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report  : Chinook Exploitation Rate Report DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report  : Chinook Exploitation Rate Report DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title   : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "Stock   : " & StockTitle(Stk)
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      PrnLine = "ESTIMATED AEQ-TOTAL MORTALITY EXPLOITATION RATE BY FISHERY-TIME STRATA"
      rw.WriteLine(PrnLine)
      PrnLine = "Distribution of AEQ Exploitation Rate (Percent)"
      rw.WriteLine(PrnLine)
      PrnLine = "================================================================="
      rw.WriteLine(PrnLine)
      PrnLine = " Fishery       "
      For TStep As Integer = 1 To NumSteps
         PrnLine += String.Format("{0,10}", TimeStepName(TStep))
      Next
      PrnLine += "     Total"
      rw.WriteLine(PrnLine)

      PrnLine = "================================================================="
      rw.WriteLine(PrnLine)
      PrnLine = " (-----) No Fishery    (*****) No Stock Impact in Fishery"
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      '..... Generate Stock Impact Report

      For Fish As Integer = 1 To NumFish + 1
         SumStkFishMort = 0
         For Age As Integer = MinAge To MaxAge
            SumStkFishMort += TotStkMort(Fish, Age, NumSteps + 1)
         Next
         If SumStkFishMort = 0 Then GoTo NextTFM1
         If Fish <= NumFish Then
            If FisheryName(Fish).Length < 20 Then
               PrnLine = String.Format("{0,11}", FisheryName(Fish))
            Else
               PrnLine = String.Format("{0,11}", FisheryName(Fish).Substring(0, 20))
            End If
         Else
            rw.WriteLine("")
            PrnLine = "TStep Total"
         End If
         AnyAgeMorts = 0
         For Age As Integer = MinAge To MaxAge
            If TotStkMort(Fish, Age, NumSteps + 1) = 0 Then
               GoTo NextTFM2
            Else
               AnyAgeMorts += 1
            End If
            If AnyAgeMorts > 1 Then
               PrnLine = String.Format("           {0,3} ", Age.ToString)
            Else
               PrnLine += String.Format("{0,3} ", Age.ToString)
            End If
            For TStep As Integer = 1 To NumSteps + 1
               If TotFisheryMort(Fish, TStep) Then
                  If TotAEQMort(Fish, Age, TStep) Then
                     PrnLine += String.Format("{0,10}", (TotAEQMort(Fish, Age, TStep) * 100.0! / TotalAEQCohort).ToString("###0.0000"))
                  Else
                     PrnLine += "     *****"
                  End If
               Else
                  PrnLine += "     -----"
               End If
            Next TStep
            rw.WriteLine(PrnLine)
NextTFM2:
         Next Age
         rw.WriteLine("")
NextTFM1:
      Next Fish

      PrnLine = "TStep Total ALL"
      For TStep As Integer = 1 To NumSteps + 1
         SumTimeStep = 0
         For Age As Integer = MinAge To MaxAge
            SumTimeStep += TotAEQMort(NumFish + 1, Age, TStep)
         Next Age
         PrnLine += String.Format("{0,10}", (SumTimeStep * 100.0! / TotalAEQCohort).ToString("###0.0000"))
      Next TStep
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      For Age As Integer = MinAge To MaxAge
         PrnLine = "Escapement   "
         PrnLine += Age.ToString
         PrnLine += " "
         For TStep As Integer = 1 To NumSteps + 1
            If TotStkMort(NumFish + 2, Age, TStep) Then
               PrnLine += String.Format("{0,10}", CLng(TotAEQMort(NumFish + 2, Age, TStep)).ToString)
            Else
               PrnLine += "         0"
            End If
         Next
         rw.WriteLine(PrnLine)
      Next
      rw.WriteLine("")
      For Age As Integer = MinAge To MaxAge
         PrnLine = "TotalAEQMort "
         PrnLine += Age.ToString
         PrnLine += " "
         For TStep As Integer = 1 To NumSteps + 1
            If TotStkMort(NumFish + 1, Age, TStep) Then
               PrnLine += String.Format("{0,10}", CLng(TotAEQMort(NumFish + 1, Age, TStep)).ToString)
            Else
               PrnLine += "         0"
            End If
         Next TStep
         rw.WriteLine(PrnLine)
      Next Age
      rw.WriteLine("")
      For Age As Integer = MinAge To MaxAge
         PrnLine = "TotalStkMort "
         PrnLine += Age.ToString
         PrnLine += " "
         For TStep = 1 To NumSteps + 1
            If TotStkMort(NumFish + 1, Age, TStep) Then
               PrnLine += String.Format("{0,10}", CLng(TotStkMort(NumFish + 1, Age, TStep)).ToString)
            Else
               PrnLine += "         0"
            End If
         Next TStep
         rw.WriteLine(PrnLine)
      Next Age

   End Sub

   Sub ERDistributionReport(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      '------------------- ER / MORTALITY DISTRIBUTION REPORT ---

      'Mortality Type  1=Catch; 2=Total Mortality; 3=Total AEQ Mortality)
      '                4= Landed Catch Plus Escapement May 96
      '                5= Total Mortality Plus Escapement May 96
      '                6= Total AEQ Mortality Plus Escapement Mar. 96

      Dim MortProp As Double
      Dim RepStks(NumStk), RepFish(1, NumFish) As Integer
      Dim RepFishGroupName(1) As String
      Dim RepTMort(1, 1) As Double
      Dim ParseOld, ParseNew, NumRepStks, NumRepFish, FishGrpNum, FishNum As Integer
      Dim StkPos, AggrFish, Sect, NumSect, BegStk, EndStk, EndCol, GrpNum As Integer

      '- Parse Option Strings from Report Driver Table Fields until All Groups are Read
      MortalityType = CInt(Opt1)
      NumRepStks = 0
      ParseOld = 1

      '- Stock Group Numbers
      For Stk = 1 To NumStk
         NumRepStks += 1
         ParseNew = InStr(ParseOld, Opt2, ",")
         If ParseNew = 0 Then
            '- Last Stock in List
            RepStks(NumRepStks) = CInt(Opt2.Substring(ParseOld - 1, Opt2.Length - ParseOld + 1))
            Exit For
         Else
            RepStks(NumRepStks) = CInt(Opt2.Substring(ParseOld - 1, ParseNew - ParseOld))
         End If
         ParseOld = ParseNew + 1
      Next

      '- Fishery Groups
      FishGrpNum = 0
      GrpNum = 0
      NumRepFish = 0
      ParseOld = 1
      For Fish As Integer = 0 To NumFish * 2
         '- Parse Option String using Comma Separators
         ParseNew = InStr(ParseOld, Opt3, ",")
         If ParseNew = 0 Then
            '- Last Fishery in List
            FishNum = CInt(Opt3.Substring(ParseOld - 1, Opt3.Length - ParseOld + 1))
            RepFish(FishGrpNum, GrpNum) = FishNum
            Exit For
         Else
            If Fish = 0 Then
               '- First Number is Number of Fishery Groups
               NumRepFish = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
               ReDim RepFish(NumRepFish, NumFish)
               ReDim RepFishGroupName(NumRepFish)
            Else
               If GrpNum = 0 And FishGrpNum = 0 Then
                  '- First Number of First Group (Number of Fisheries in Group)
                  FishGrpNum += 1
                  RepFish(FishGrpNum, 0) = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
               Else
                  GrpNum += 1
                  If GrpNum > RepFish(FishGrpNum, 0) Then
                     GrpNum = 0
                     FishGrpNum += 1
                     '- First Number of Group (Number of Fisheries in Group)
                     RepFish(FishGrpNum, 0) = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
                  Else
                     RepFish(FishGrpNum, GrpNum) = CInt(Opt3.Substring(ParseOld - 1, ParseNew - ParseOld))
                  End If
               End If
            End If
         End If
         ParseOld = ParseNew + 1
      Next

      '- Fishery Aggregate Group Names
      ParseOld = 1
      For Fish As Integer = 1 To NumRepFish
         ParseNew = InStr(ParseOld, Opt4, ",")
         If ParseNew = 0 Then
            RepFishGroupName(Fish) = Opt4.Substring(ParseOld - 1, Opt4.Length - ParseOld + 1)
            Exit For
         Else
            RepFishGroupName(Fish) = Opt4.Substring(ParseOld - 1, ParseNew - ParseOld)
         End If
         ParseOld = ParseNew + 1
      Next

      '- Sum Mortalities and Escapement by Mortality Option Selection, stock, and Fishery Grouping
      If MortalityType > 3 Then
         '- Add space for escapement for these Mortality Types
         ReDim RepTMort(NumRepStks, NumRepFish + 1)
      Else
         ReDim RepTMort(NumRepStks, NumRepFish)
      End If
      StkPos = 0
      For StkPos = 1 To NumRepStks
         Stk = RepStks(StkPos)
         For AggrFish = 1 To NumRepFish
            For FishNum = 1 To RepFish(AggrFish, 0)
               Fish = RepFish(AggrFish, FishNum)
               For Age As Integer = MinAge To MaxAge
                  For TStep As Integer = 1 To NumSteps
                     '- Don't Use Time 1 for Chinook (Double Count)
                     If SpeciesName = "CHINOOK" And TStep = 1 Then GoTo SkipTime1
                     '- Landed Catch
                     If MortalityType = 1 Or MortalityType = 4 Then
                        RepTMort(StkPos, 0) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                        RepTMort(StkPos, AggrFish) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                     End If
                     '- Total Mortality
                     If MortalityType = 2 Or MortalityType = 5 Then
                        RepTMort(StkPos, 0) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                        RepTMort(StkPos, AggrFish) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     End If
                     '- Total AEQ Mortality
                     If MortalityType = 3 Or MortalityType = 6 Then
                        If TerminalFisheryFlag(Fish, TStep) = PTerm Then
                           RepTMort(StkPos, 0) += (LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)) * AEQ(Stk, Age, TStep)
                           RepTMort(StkPos, AggrFish) += (LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        Else
                           RepTMort(StkPos, 0) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                           RepTMort(StkPos, AggrFish) += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                        End If
                     End If
SkipTime1:
                  Next
               Next
            Next
         Next
         If MortalityType > 3 Then
            '- Add Escapement for Selected Mortality Types
            For Age As Integer = MinAge To MaxAge
               For TStep = 1 To NumSteps
                  RepTMort(StkPos, 0) += Escape(Stk, Age, TStep)
                  RepTMort(StkPos, NumRepFish + 1) += Escape(Stk, Age, TStep)
               Next
            Next
         End If
NextStk1:
      Next

      '------ PRINT PROPORTION OF STOCK MORTALITY BY Aggregated Fishery ---------------

      '- Determine number of pages - 8 Stocks per page
      If (NumRepStks > 8) Then
         NumSect = NumRepStks \ 8
         If (NumRepStks Mod 8) <> 0 Then NumSect = NumSect + 1
      Else
         NumSect = 1
      End If

      '- PRINT HEADER DATA
      PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)

      If ReportDriverName.Length > 20 Then
         PrnLine = "Report  : Mortality Distribution Report    DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report  : Mortality Distribution Report    DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)

      PrnLine = "Title   : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

      Select Case MortalityType
         Case 1
            PrnLine = "Distribution of Landed Catch"
         Case 2
            PrnLine = "Distribution of Total Mortality"
         Case 3
            PrnLine = "Distribution of AEQ Total Mortality"
         Case 4
            PrnLine = "Distribution of Landed Catch Plus Escapement"
         Case 5
            PrnLine = "Distribution of Total Mortality Plus Escapement"
         Case 6
            PrnLine = "Distribution of AEQ Total Mortality Plus Escapement"
      End Select
      rw.WriteLine(PrnLine)

      '- Loop for Each Section (Page)

      For Sect = 1 To NumSect
         If Sect = NumSect Then
            EndCol = NumRepStks - (Sect * 8 - 8)
         Else
            EndCol = 8
         End If
         PrnLine = "========================="
         For Stk = 1 To EndCol
            PrnLine &= "=========="
         Next
         rw.WriteLine(PrnLine)
         rw.WriteLine("")
         PrnLine = String.Format("Fishery", "@@@@@@@@@@@@@@@@@@@@@@@@@")
         StkPos = 0
         PrnLine = "                         "
         BegStk = Sect * 8 - 7
         If (Sect * 8 + 1) > NumRepStks Then
            EndStk = NumRepStks
         Else
            EndStk = Sect * 8
         End If
         StkPos = 0
         For Stk = BegStk To EndStk
            PrnLine &= String.Format("{0,10}", StockName(RepStks(Stk)))
         Next Stk
         rw.WriteLine(PrnLine)
         PrnLine = "========================="
         For Stk = 1 To EndCol
            PrnLine &= "=========="
         Next
         rw.WriteLine(PrnLine)

         For AggrFish = 1 To NumRepFish
            PrnLine = String.Format("{0,25}", RepFishGroupName(AggrFish), "@@@@@@@@@@@@@@@@@@@@@@@@@")
            For Stk = BegStk To EndStk
               If RepTMort(Stk, 0) = 0 Then
                  MortProp = 0
               Else
                  MortProp = RepTMort(Stk, AggrFish) / RepTMort(Stk, 0)
               End If
               PrnLine &= String.Format("{0,10}", MortProp.ToString("0.00000"))
            Next
            rw.WriteLine(PrnLine)
            PrnLine = ""
         Next
         '- Escapement Line
         If MortalityType > 3 Then
            PrnLine = String.Format("{0,25}", "** Escapement **", "@@@@@@@@@@@@@@@@@@@@@@@@@")
            For Stk = BegStk To EndStk
               If RepTMort(Stk, 0) = 0 Then
                  MortProp = 0
               Else
                  MortProp = RepTMort(Stk, NumRepFish + 1) / RepTMort(Stk, 0)
               End If
               PrnLine &= String.Format("{0,10}", MortProp.ToString("0.00000"))
NextStk5:
            Next
            rw.WriteLine(PrnLine)
            PrnLine = ""
         End If

         PrnLine = "========================="
         For Stk = 1 To EndCol
            PrnLine &= "=========="
         Next
         rw.WriteLine(PrnLine)
         rw.WriteLine("")

      Next Sect

   End Sub

   Sub FisheryStockComposition(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      '- Fishery Stock Composition Report Summed for ALL Ages 

      Dim ParseOld, ParseNew, NumSelFish, FishNum, NumContribStk, NumContribStkAge As Integer
      Dim FisheryRepSelect(NumFish) As Integer
      Dim TotFisheryMort(NumSteps + 1), TempVal, StkTempVal, StkAgeTempVal, StkTempVal24 As Double

      '- Parse Option1 line for Fishery Selections
      ParseOld = 1
      NumSelFish = 0
      For Fish As Integer = 1 To NumFish
         ParseNew = InStr(ParseOld, Opt1, ",")
         NumSelFish += 1
         If ParseNew = 0 Then
            FisheryRepSelect(Fish) = CInt(Opt1.Substring(ParseOld - 1, Opt1.Length - ParseOld + 1))
            Exit For
         Else
            FisheryRepSelect(Fish) = CInt(Opt1.Substring(ParseOld - 1, ParseNew - ParseOld))
         End If
         ParseOld = ParseNew + 1
      Next

      For FishNum = 1 To NumSelFish
         Fish = FisheryRepSelect(FishNum)

         '- Sum Total Mortality by Fishery and Time Step
         ReDim TotFisheryMort(NumSteps + 2)
         NumContribStk = 0
         NumContribStkAge = 0
         For Stk As Integer = 1 To NumStk
            StkTempVal = 0
            For Age As Integer = MinAge To MaxAge
               StkAgeTempVal = 0
               For TStep As Integer = 1 To NumSteps
                  TempVal = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  TotFisheryMort(TStep) += TempVal
                  TotFisheryMort(NumSteps + 1) += TempVal
                  If SpeciesName = "CHINOOK" And TStep > 1 Then TotFisheryMort(NumSteps + 2) += TempVal
                  StkTempVal += TempVal
                  StkAgeTempVal += TempVal
               Next
               If StkAgeTempVal > 0 Then NumContribStkAge += 1
            Next
            If StkTempVal > 0 Then NumContribStk += 1
         Next

         '- PRINT HEADER DATA
         PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
         PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
         If RunIDNameSelect.Length > 25 Then
            PrnLine &= "  Run Name: " & String.Format("{0,-25}", RunIDNameSelect.Substring(0, 25))
         Else
            PrnLine &= "  Run Name: " & String.Format("{0,-25}", RunIDNameSelect)
         End If
         PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
         rw.WriteLine(PrnLine)
         If ReportDriverName.Length > 20 Then
            PrnLine = "Report  : Fishery Stock Composition Report DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
         Else
            PrnLine = "Report  : Fishery Stock Composition Report DRV File: " & String.Format("{0,-20}", ReportDriverName)
         End If
         PrnLine &= "      RepDate:" & Now.ToString
         rw.WriteLine(PrnLine)
         PrnLine = "Title   : " & RunIDNameSelect
         rw.WriteLine(PrnLine)
         rw.WriteLine("")
         PrnLine = "Fishery : " & FisheryTitle(Fish)
         rw.WriteLine(PrnLine)
         PrnLine = "===================================="
         For TStep = 1 To NumSteps + 1
            PrnLine &= "=========="
         Next
         If TotFisheryMort(NumSteps + 1) = 0 Then
            PrnLine = "No Fishing Mortality for this Fishery"
            rw.WriteLine(PrnLine)
            rw.WriteLine("")
            GoTo NextRepFishery
         End If

         If SpeciesName = "CHINOOK" Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)
         PrnLine = "      Stock Name                    "
         For TStep As Integer = 1 To NumSteps
            PrnLine &= String.Format("{0,10}", TimeStepName(TStep))
         Next
         If SpeciesName = "CHINOOK" Then
            PrnLine &= "  Time 2-4     Total"
         Else
            PrnLine &= "    Total"
         End If
         rw.WriteLine(PrnLine)
         PrnLine = "===================================="
         For TStep As Integer = 1 To NumSteps + 1
            PrnLine &= "=========="
         Next
         If SpeciesName = "CHINOOK" Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)

         '- Stock Composition Lines
         For Stk As Integer = 1 To NumStk
            StkTempVal = 0
            StkTempVal24 = 0
            '- Check if Stock Contributes to this Fishery
            For TStep As Integer = 1 To NumSteps
               For Age As Integer = MinAge To MaxAge
                  TempVal = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  StkTempVal += TempVal
                  If SpeciesName = "CHINOOK" And TStep > 1 Then StkTempVal24 += TempVal
               Next
            Next
            '- If Yes, Print line by Time Step
            If StkTempVal > 0 Then
               If StockTitle(Stk).Length > 36 Then
                  PrnLine = String.Format("{0,36}", StockTitle(Stk).Substring(0, 36))
               Else
                  PrnLine = String.Format("{0,36}", StockTitle(Stk))
               End If
               For TStep As Integer = 1 To NumSteps
                  If TotFisheryMort(TStep) = 0 Then
                     PrnLine &= "      ****"
                  Else
                     TempVal = 0
                     For Age As Integer = MinAge To MaxAge
                        TempVal += LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                     Next
                     PrnLine &= String.Format("{0,10}", (TempVal / TotFisheryMort(TStep) * 100).ToString("##0.00"))
                  End If
               Next
               If SpeciesName = "CHINOOK" Then
                  If TotFisheryMort(NumSteps + 2) = 0 Then
                     PrnLine &= "     ****"
                  Else
                     PrnLine &= String.Format("{0,10}", ((StkTempVal24 / TotFisheryMort(NumSteps + 2)) * 100).ToString("##0.00"))
                  End If
               End If
               PrnLine &= String.Format("{0,10}", ((StkTempVal / TotFisheryMort(NumSteps + 1)) * 100).ToString("##0.00"))
               rw.WriteLine(PrnLine)
            End If
         Next
         PrnLine = "===================================="
         For TStep As Integer = 1 To NumSteps + 1
            PrnLine &= "=========="
         Next
         If SpeciesName = "CHINOOK" Then PrnLine &= "=========="
         rw.WriteLine(PrnLine)
         rw.WriteLine("")
NextRepFishery:
      Next

   End Sub

   Sub StockImpactsPer1000(ByVal Opt1 As String, ByVal Opt2 As String, ByVal Opt3 As String, ByVal Opt4 As String, ByVal Opt5 As String, ByVal Opt6 As String)
      Dim TempVal, FisherySum, FisheryStepSum, FisheryTotalSum As Double

      Stk = CInt(Opt1)
      '- PRINT HEADER DATA
      PrnLine = "Species : " + String.Format("{0,7}", SpeciesName)
      PrnLine &= String.Format("{0,20}{1,4}", "Version#:", FramVersion)
      If RunIDNameSelect.Length > 25 Then
         PrnLine &= "  Run Name: " & String.Format("{0,-25}", RunIDNameSelect.Substring(0, 25))
      Else
         PrnLine &= "  Run Name: " & String.Format("{0,-25}", RunIDNameSelect)
      End If
      PrnLine &= " RunDate:" + RunIDRunTimeDateSelect.ToString
      rw.WriteLine(PrnLine)
      If ReportDriverName.Length > 20 Then
         PrnLine = "Report  : Stock Impacts Per 1000 Report    DRV File: " & String.Format("{0,-20}", ReportDriverName.Substring(0, 20))
      Else
         PrnLine = "Report  : Stock Impacts Per 1000 Report    DRV File: " & String.Format("{0,-20}", ReportDriverName)
      End If
      PrnLine &= "      RepDate:" & Now.ToString
      rw.WriteLine(PrnLine)
      PrnLine = "Title   : " & RunIDNameSelect
      rw.WriteLine(PrnLine)
      rw.WriteLine("")
      PrnLine = "Stock   : " & StockTitle(Stk)
      rw.WriteLine(PrnLine)
      PrnLine = "======================================"
      For TStep = 1 To NumSteps + 1
         PrnLine &= "=========="
      Next
      rw.WriteLine(PrnLine)
      PrnLine = "      Fishery Name                 Age"
      For TStep = 1 To NumSteps
         PrnLine &= String.Format("{0,10}", TimeStepName(TStep))
      Next
      PrnLine &= String.Format("{0,10}", "Total")
      rw.WriteLine(PrnLine)
      PrnLine = "======================================"
      For TStep As Integer = 1 To NumSteps + 1
         PrnLine &= "=========="
      Next
      rw.WriteLine(PrnLine)

      For Fish As Integer = 1 To NumFish
         FisheryTotalSum = 0
         For Age As Integer = MinAge To MaxAge
            FisherySum = 0
            If Age = MinAge Then
               If FisheryTitle(Fish).Length < 36 Then
                  PrnLine = String.Format("{0,36}", FisheryTitle(Fish))
               Else
                  PrnLine = String.Format("{0,36}", FisheryTitle(Fish).Substring(0, 36))
               End If
               PrnLine &= String.Format("{0,2}", Age.ToString)
            Else
               PrnLine = String.Format("{0,38}", Age.ToString)
            End If

            For TStep = 1 To NumSteps
               FisheryStepSum = (TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)) / ModelStockProportion(Fish)
               FisheryTotalSum += FisheryStepSum
               If FisheryStepSum = 0 Then
                  PrnLine &= String.Format("{0,10}", "-----")
               Else
                  TempVal = LandedCatch(Stk, 3, Fish, TStep) + NonRetention(Stk, 3, Fish, TStep) + Shakers(Stk, 3, Fish, TStep) + DropOff(Stk, 3, Fish, TStep) + MSFLandedCatch(Stk, 3, Fish, TStep) + MSFNonRetention(Stk, 3, Fish, TStep) + MSFShakers(Stk, 3, Fish, TStep) + MSFDropOff(Stk, 3, Fish, TStep)
                  FisherySum += TempVal
                  If TempVal = 0 Then
                     PrnLine &= String.Format("{0,10}", "*****")
                  Else
                     'PrnLine &= String.Format("{0,10}", (TempVal * (1000 / (FisheryStepSum / ModelStockProportion(Fish)))).ToString(" ###0.00"))
                     PrnLine &= String.Format("{0,10}", (TempVal * (1000 / FisheryStepSum)).ToString(" ###0.00"))
                  End If
               End If
            Next
            If FisheryTotalSum = 0 Then
               PrnLine &= String.Format("{0,10}", "-----")
            Else
               PrnLine &= String.Format("{0,10}", (FisherySum * (1000 / FisheryTotalSum)).ToString(" ###0.00"))
            End If
            rw.WriteLine(PrnLine)
         Next
      Next
      PrnLine = "====================================="
      For TStep = 1 To NumSteps + 1
         PrnLine &= "=========="
      Next
      rw.WriteLine(PrnLine)
      rw.WriteLine("")

   End Sub

   Sub PSCCohoSpreadsheet()

      Dim Stock, RecNum, RunYear As Integer
      Dim PSCStock(16, 4) As Integer
      Dim PSCFishery(NumFish) As Integer
      Dim PSCFishName(28), PSCSSName As String
      Dim PSCCatch(27, 30) As Double
      Dim PSCBase(198, 5) As Double
      Dim PSCBaseER(,) As Double
      Dim LCat, CNR, Drop, Shak, InitCoht, Esc As Double
      Dim MLCat, MCNR, MDrop, MShak As Double
      Dim CmdStr As String

      '-- UnMarked List of Stocks for PSC Periodic Report

      PSCStock(1, 0) = 2      '- Skagit
      PSCStock(1, 1) = 17
      PSCStock(1, 2) = 23
      PSCStock(2, 0) = 1      '- Stillaguamish
      PSCStock(2, 1) = 29
      PSCStock(3, 0) = 1      '- Snohomish
      PSCStock(3, 1) = 35
      PSCStock(4, 0) = 4      '- Hood Canal
      PSCStock(4, 1) = 45
      PSCStock(4, 2) = 51
      PSCStock(4, 3) = 55
      PSCStock(4, 4) = 59
      PSCStock(5, 0) = 4      '- US JDF
      PSCStock(5, 1) = 107
      PSCStock(5, 2) = 111
      PSCStock(5, 3) = 115
      PSCStock(5, 4) = 117
      PSCStock(6, 0) = 1      '- Quillayute
      PSCStock(6, 1) = 131
      PSCStock(7, 0) = 1      '- Hoh
      PSCStock(7, 1) = 135
      PSCStock(8, 0) = 1      '- Queets
      PSCStock(8, 1) = 139
      PSCStock(9, 0) = 3      '- Grays harbor
      PSCStock(9, 1) = 149
      PSCStock(9, 2) = 153
      PSCStock(9, 3) = 157
      PSCStock(10, 0) = 1     '- Lower Fraser
      PSCStock(10, 1) = 227
      PSCStock(11, 0) = 1     '- Upper Fraser
      PSCStock(11, 1) = 231
      PSCStock(12, 0) = 1     '- GS Mainland
      PSCStock(12, 1) = 207
      PSCStock(13, 0) = 1     '- GS Vanc Isl
      PSCStock(13, 1) = 211
      PSCStock(14, 0) = 1     '- SW Vanc Isl
      PSCStock(14, 1) = 219
      PSCStock(15, 0) = 1     '- Col River Early (colreh)
      PSCStock(15, 1) = 166
      PSCStock(16, 0) = 1     '- Col River Late (colrlh)
      PSCStock(16, 1) = 176

      '- Marked List for Bill Gazey Project ONLY !!!!

      'PSCStock(1, 0) = 2      '- Skagit
      'PSCStock(1, 1) = 20
      'PSCStock(1, 2) = 22
      'PSCStock(2, 0) = 1      '- Stillaguamish
      'PSCStock(2, 1) = 32
      'PSCStock(3, 0) = 1      '- Snohomish
      'PSCStock(3, 1) = 36
      'PSCStock(4, 0) = 3      '- Hood Canal
      'PSCStock(4, 1) = 48
      'PSCStock(4, 2) = 54
      'PSCStock(4, 3) = 58
      ''PSCStock(4, 4) = 60
      'PSCStock(5, 0) = 2      '- US JDF
      'PSCStock(5, 1) = 110
      'PSCStock(5, 2) = 114
      ''PSCStock(5, 3) = 116
      ''PSCStock(5, 4) = 118
      'PSCStock(6, 0) = 1      '- Quillayute
      'PSCStock(6, 1) = 134
      'PSCStock(7, 0) = 1      '- Hoh
      'PSCStock(7, 1) = 138
      'PSCStock(8, 0) = 1      '- Queets
      'PSCStock(8, 1) = 142
      'PSCStock(9, 0) = 2      '- Grays harbor
      'PSCStock(9, 1) = 152
      'PSCStock(9, 2) = 156
      ''PSCStock(9, 3) = 158
      'PSCStock(10, 0) = 1     '- Lower Fraser
      'PSCStock(10, 1) = 226
      'PSCStock(11, 0) = 1     '- Upper Fraser
      'PSCStock(11, 1) = 230
      'PSCStock(12, 0) = 1     '- GS Mainland
      'PSCStock(12, 1) = 206
      'PSCStock(13, 0) = 1     '- GS Vanc Isl
      'PSCStock(13, 1) = 210
      'PSCStock(14, 0) = 1     '- SW Vanc Isl
      'PSCStock(14, 1) = 218
      'PSCStock(15, 0) = 1     '- Col River Early (colreh)
      'PSCStock(15, 1) = 166
      'PSCStock(16, 0) = 1     '- Col River Late (colrlh)
      'PSCStock(16, 1) = 176

      '- FRAM Coho Fishery Assignment to PSC-Fishery List
      PSCFishery(1) = 18 '-No Cal Trm	No Calif Cst Terminal Catch
      PSCFishery(2) = 18 '-Cn Cal Trm	Cntrl Cal Cst Term Catch
      PSCFishery(3) = 18 '-Ft Brg Spt	Fort Bragg Sport
      PSCFishery(4) = 18 '-Ft Brg Trl	Fort Bragg Troll
      PSCFishery(5) = 18 '-Ca KMZ Spt	KMZ Sport
      PSCFishery(6) = 18 '-Ca KMZ Trl	KMZ Troll
      PSCFishery(7) = 18 '-So Cal Spt	So Calif. Sport
      PSCFishery(8) = 18 '-So Cal Trl	So Calif. Troll
      PSCFishery(9) = 18 '-So Ore Trm	So Ore Coast Terminal Catch
      PSCFishery(10) = 18 '-Or Prv Trm	Ore Private Hat Term Catch
      PSCFishery(11) = 18 '-SMi Or Trm	So Mid Ore Coast Term Catch
      PSCFishery(12) = 18 '-NMi Or Trm	No Mid Ore Coast Term Catch
      PSCFishery(13) = 18 '-No Ore Trm	North Ore Coast Term Catch
      PSCFishery(14) = 18 '-Or Cst Trm	Oregon Coast Term Catch
      PSCFishery(15) = 18 '-Brkngs Spt	Brookings Sport
      PSCFishery(16) = 18 '-Brkngs Trl	Brookings Troll
      PSCFishery(17) = 18 '-Newprt Spt	Newport Sport
      PSCFishery(18) = 18 '-Newprt Trl	Newport Troll
      PSCFishery(19) = 18 '-Coos B Spt	Coos Bay Sport
      PSCFishery(20) = 18 '-Coos B Trl	Coos Bay Troll
      PSCFishery(21) = 18 '-Tillmk Spt	Tillamook Sport
      PSCFishery(22) = 18 '-Tillmk Trl	Tillamook Troll
      PSCFishery(23) = 24 '-Buoy10 Spt	Col. Rvr. Buoy 10 Sport
      PSCFishery(24) = 24 '-L ColR Spt	Col. Rvr. Lower R Sport
      PSCFishery(25) = 24 '-L ColR Net	Col. Rvr. Lower R Net
      PSCFishery(26) = 24 '-Yngs B Net	Col. Rvr. Youngs Bay Net
      PSCFishery(27) = 24 '-LCROrT Spt	Col. Rvr. Ore Trib Spt
      PSCFishery(28) = 24 '-Clackm Spt	Clackamas R Sport
      PSCFishery(29) = 24 '-SandyR Spt	Sandy R Sport
      PSCFishery(30) = 24 '-LCRWaT Spt	Col. Rvr. Wash Trib Spt
      PSCFishery(31) = 24 '-UpColR Spt	Col. Rvr. Sport Above Bonneville
      PSCFishery(32) = 24 '-UpColR Net	Col. Rvr. Net Above Bonneville
      PSCFishery(33) = 17 '-A1-Ast Spt	WA Area 1 & Astoria Sport
      PSCFishery(34) = 16 '-A1-Ast Trl	WA Area 1 & Astoria Troll
      PSCFishery(35) = 16 '-Area2TrlNT	WA Area 2 Non-Treaty Troll
      PSCFishery(36) = 16 '-Area2TrlTR	WA Area 2 Treaty Troll
      PSCFishery(37) = 17 '-Area 2 Spt	WA Area 2 Sport
      PSCFishery(38) = 16 '-Area3TrlNT	WA Area 3 Non-Treaty Troll
      PSCFishery(39) = 16 '-Area3TrlTR	WA Area 3 Treaty Troll
      PSCFishery(40) = 17 '-Area 3 Spt	WA Area 3 Sport
      PSCFishery(41) = 17 '-Area 4 Spt	WA Area 4 Sport
      PSCFishery(42) = 16 '-A4/4BTrlNT	WA Area 4/4B Non-Treaty Troll
      PSCFishery(43) = 16 '-A4/4BTrlTR	WA Area 4/4B Treaty Troll
      PSCFishery(44) = 19 '-A 5-6C Trl	WA Area 5-6-6C Troll
      PSCFishery(45) = 24 '-Willpa Spt	Willapa Bay Sport (2.1)
      PSCFishery(46) = 24 '-Wlp Tb Spt	Willapa Tributary Sport
      PSCFishery(47) = 24 '-WlpaBT Net	Willapa Bay & FW Trib Net
      PSCFishery(48) = 24 '-GryHbr Spt	Grays Harbor Sport (2.2)
      PSCFishery(49) = 24 '-SGryHb Spt	South Grays Harbor Sport
      PSCFishery(50) = 24 '-GryHbr Net	Grays Harbor Estuary Net
      PSCFishery(51) = 24 '-Hump R Spt	Humptulips R Sport
      PSCFishery(52) = 24 '-LwCheh Net	Lower Chehalis R Net
      PSCFishery(53) = 24 '-Hump R C&S	Humptulips R C&S
      PSCFishery(54) = 24 '-Chehal Spt	Chehalis R Sport
      PSCFishery(55) = 24 '-Hump R Net	Humptulips R Net
      PSCFishery(56) = 24 '-UpCheh Net	Upper Chehalis R Net
      PSCFishery(57) = 24 '-Chehal C&S	Chehalis R C&S
      PSCFishery(58) = 24 '-Wynoch Spt	Wynochee R Sport
      PSCFishery(59) = 24 '-Hoquam Spt	Hoquiam R Sport
      PSCFishery(60) = 24 '-Wishkh Spt	Wishkah R Sport
      PSCFishery(61) = 24 '-Satsop Spt	Satsop R Sport
      PSCFishery(62) = 24 '-Quin R Spt	Quinault R Sport
      PSCFishery(63) = 24 '-Quin R Net	Quinault R Net
      PSCFishery(64) = 24 '-Quin R C&S	Quinault R C&S
      PSCFishery(65) = 24 '-Queets Spt	Queets R Sport
      PSCFishery(66) = 24 '-Clrwtr Spt	Clearwater R Sport
      PSCFishery(67) = 24 '-Salm R Spt	Salmon R Sport (Queets)
      PSCFishery(68) = 24 '-Queets Net	Queets R Net
      PSCFishery(69) = 24 '-Queets C&S	Queets R C&S
      PSCFishery(70) = 24 '-Quilly Spt	Quillayute R Sport
      PSCFishery(71) = 24 '-Quilly Net	Quillayute R Net
      PSCFishery(72) = 24 '-Quilly C&S	Quillayute R C&S
      PSCFishery(73) = 24 '-Hoh R  Spt	Hoh R Sport
      PSCFishery(74) = 24 '-Hoh R  Net	Hoh R Net
      PSCFishery(75) = 24 '-Hoh R  C&S	Hoh R C&S
      PSCFishery(76) = 24 '-Mak FW Spt	Makah Tributary Sport
      PSCFishery(77) = 24 '-Mak FW Net	Makah Freshwater Net
      PSCFishery(78) = 24 '-Makah  C&S	Makah C&S
      PSCFishery(79) = 16 '-A 4-4A Net	WA Area 4-4A Net
      PSCFishery(80) = 19 '-A4B6CNetNT	WA Area 4B-5-6C Non-Treaty Net
      PSCFishery(81) = 19 '-A4B6CNetTR	WA Area 4B-5-6C Treaty Net
      PSCFishery(82) = 19 '-Ar6D NetNT	6D Non-Treaty Net (Dungeness Bay & R)
      PSCFishery(83) = 19 '-Ar6D NetTR	6D Treaty Net (Dungeness Bay & R)
      PSCFishery(84) = 19 '-Elwha  Net	Elwha R Net
      PSCFishery(85) = 19 '-WJDF T Net	West JDF Straits Trib Net
      PSCFishery(86) = 19 '-EJDF T Net	East JDF Straits Trib Net
      PSCFishery(87) = 20 '-A6-7ANetNT	WA Area 7-7A Non-Treaty Net
      PSCFishery(88) = 20 '-A6-7ANetTR	WA Area 7-7A Treaty Net
      PSCFishery(89) = 19 '-EJDF FWSpt	East JDF Straits Trib Sport
      PSCFishery(90) = 19 '-WJDF FWSpt	West JDF Straits Trib Sport
      PSCFishery(91) = 19 '-Area 5 Spt	WA Area 5 Sport (Sekiu)
      PSCFishery(92) = 19 '-Area 6 Spt	WA Area 6 Sport (Port Angeles)
      PSCFishery(93) = 21 '-Area 7 Spt	WA Area 7 Sport (San Juan Islands)
      PSCFishery(94) = 19 '-Dung R Spt	Dungeness R Sport
      PSCFishery(95) = 19 '-ElwhaR Spt	Elwha R Sport
      PSCFishery(96) = 20 '-A7BCDNetNT	WA Area 7B-7C-7D Non-Treaty Net
      PSCFishery(97) = 20 '-A7BCDNetTR	WA Area 7B-7C-7D Treaty Net
      PSCFishery(98) = 23 '-Nook R Net	Nooksack R Net
      PSCFishery(99) = 24 '-Nook R Spt	Nooksack R Sport
      PSCFishery(100) = 24 '-Samh R Spt	Samish R Sport
      PSCFishery(101) = 23 '-Ar 8 NetNT	WA Area 8 Non-Treaty Net (Skagit)
      PSCFishery(102) = 23 '-Ar 8 NetTR	WA Area 8 Treaty Net (Skagit)
      PSCFishery(103) = 24 '-Skag R Net	Skagit R Net
      PSCFishery(104) = 24 '-SkgR TsNet	Skagit River Test Net
      PSCFishery(105) = 24 '-SwinCh Net	Swinomish Channel Net
        PSCFishery(106) = 22 '-Ar 8-1 Spt	WA Area 8.1 Sport (Skagit)
      PSCFishery(107) = 22 '-Area 9 Spt	WA Area 9 Sport (Admirality Inlet)
      PSCFishery(108) = 24 '-Skag R Spt	Skagit R Sport
      PSCFishery(109) = 23 '-Ar8A NetNT	WA Area 8A Non-Treaty Net
      PSCFishery(110) = 23 '-Ar8A NetTR	WA Area 8A Treaty Net
      PSCFishery(111) = 23 '-Ar8D NetNT	WA Area 8D Non-Treaty Net (Tulalip Bay)
      PSCFishery(112) = 23 '-Ar8D NetTR	WA Area 8D Treaty Net (Tulalip Bay)
      PSCFishery(113) = 24 '-Stil R Net	Stillaguamish R Net
      PSCFishery(114) = 24 '-Snoh R Net	Snohomish R Net
        PSCFishery(115) = 22 '-Ar 8-2 Spt	WA Area 8.2 Sport (Everett)
      PSCFishery(116) = 24 '-Stil R Spt	Stillaguamish R Sport
      PSCFishery(117) = 24 '-Snoh R Spt	Snohomish R Sport
      PSCFishery(118) = 22 '-Ar 10  Spt	WA Area 10 Sport (Seattle)
      PSCFishery(119) = 23 '-Ar10 NetNT	WA Area 10 Non-Treaty Net (Seattle)
      PSCFishery(120) = 23 '-Ar10 NetTR	WA Area 10 Treaty Net (Seattle)
      PSCFishery(121) = 23 '-Ar10ANetNT	WA Area 10A Non-Treaty Net (Elliott Bay)
      PSCFishery(122) = 23 '-Ar10ANetTR	WA Area 10A Treaty Net (Elliott Bay)
      PSCFishery(123) = 23 '-Ar10ENetNT	WA Area 10E Non-Treaty Net (East Kitsap)
      PSCFishery(124) = 23 '-Ar10ENetTR	WA Area 10E Treaty Net (East Kitsap)
      PSCFishery(125) = 23 '-10F-G  Net	WA Area 10F-G Treaty Net (Lake Union)
      PSCFishery(126) = 24 '-Duwm R Net	Green/Duwamish R Net
      PSCFishery(127) = 24 '-Duwm R Spt	Green/Duwamish R Sport
      PSCFishery(128) = 24 '-L WaSm Spt	Lk Wash/Sammamish/Tribs Spt
      PSCFishery(129) = 22 '-Ar 11  Spt	WA Area 11 Sport (Tacoma)
      PSCFishery(130) = 23 '-Ar11 NetNT	WA Area 11 Non-Treaty Net (E/W Pass)
      PSCFishery(131) = 23 '-Ar11 NetTR	WA Area 11 Treaty Net (E/W Pass)
      PSCFishery(132) = 23 '-Ar11ANetNT	WA Area 11A Non-Treaty Net (Comm. Bay)
      PSCFishery(133) = 23 '-Ar11ANetTR	WA Area 11A Treaty Net (Comm. Bay)
      PSCFishery(134) = 24 '-Puyl R Net	Puyallup R Net
      PSCFishery(135) = 24 '-Puyl R Spt	Puyallup R Sport
      PSCFishery(136) = 22 '-Ar 13  Spt	WAArea 13 Marine Sport
      PSCFishery(137) = 23 '-Ar13 NetNT	Area 13 Non-Treaty Net (So Puget Sound)
      PSCFishery(138) = 23 '-Ar13 NetTR	Area 13 Treaty Net (So Puget Sound)
      PSCFishery(139) = 23 '-Ar13CNetNT	Area 13C Non-Treaty Net (Chambers Bay)
      PSCFishery(140) = 23 '-Ar13CNetTR	Area 13C Treaty Net (Chambers Bay)
      PSCFishery(141) = 23 '-Ar13ANetNT	Area 13A Non-Treaty Net (Carr Inlet)
      PSCFishery(142) = 23 '-Ar13ANetTR	Area 13A Treaty Net (Carr Inlet)
      PSCFishery(143) = 23 '-Ar13DNetNT	Area 13D Non-Treaty Net
      PSCFishery(144) = 23 '-Ar13DNetTR	Area 13D Treaty Net
      PSCFishery(145) = 23 '-A13FKNetNT	Area 13F-13K Non-Treaty Net
      PSCFishery(146) = 23 '-A13FKNetTR	Area 13F-13K Treaty Net
      PSCFishery(147) = 24 '-Nisq R Net	Nisqually R Net
      PSCFishery(148) = 24 '-McAlls Net	McAllister Creek Net
      PSCFishery(149) = 24 '-13D-K TSpt	13D-13K Trib Sport
      PSCFishery(150) = 24 '-Nisq R Spt	Nisqually R Sport
      PSCFishery(151) = 24 '-Desc R Spt	Deschutes R Sport
      PSCFishery(152) = 22 '-Ar 12  Spt	Area 12 Marine Sport
      PSCFishery(153) = 23 '-1212BNetNT	Area 12-12B Hood Canal Non-Treaty Net
      PSCFishery(154) = 23 '-1212BNetTR	Area 12-12B Hood Canal Treaty Net
      PSCFishery(155) = 23 '-A9-9ANetNT	Area 9/9A Non-Treaty Net
      PSCFishery(156) = 23 '-A9-9ANetTR	Area 9/9A Treaty Net (On Res)
      PSCFishery(157) = 23 '-Ar12ANetNT	Area 12A Non-Treaty Net (Quilcene Bay)
      PSCFishery(158) = 23 '-Ar12ANetTR	Area 12A Treaty Net (Quilcene Bay)
      PSCFishery(159) = 23 '-A12CDNetNT	Area 12C-12D Non-Treaty Net (SE Hood Canal)
      PSCFishery(160) = 23 '-A12CDNetTR	Area 12C-12D Treaty Net (SE Hood Canal)
      PSCFishery(161) = 24 '-Skok R Net	Skokomish R Net
      PSCFishery(162) = 24 '-Quilcn Net	Quilcene R Net
      PSCFishery(163) = 24 '-1212B TSpt	12, 12B Trib FW Sport
      PSCFishery(164) = 24 '-12A Tb Spt	12A Trib FW Sport
      PSCFishery(165) = 24 '-12C-D TSpt	12C, 12D Trib FW Sport
      PSCFishery(166) = 24 '-Skok R Spt	Skokomish R Sport
      PSCFishery(167) = 13 '-FRSLOW Trm	Lower Fraser R Term Catch
      PSCFishery(168) = 13 '-FRSUPP Trm	Upper Fraser R Term Catch
      PSCFishery(169) = 13 '-Fraser Spt	Lower Fraser River Sport
      PSCFishery(170) = 1  '-JStrBC Trl	Johnstone Strait Troll
      PSCFishery(171) = 1  '-No BC  Trl	BC Northern Troll
      PSCFishery(172) = 1  '-NoC BC Trl	BC North Central Troll
      PSCFishery(173) = 1  '-SoC BC Trl	BC South Central Troll
      PSCFishery(174) = 4  '-NW VI  Trl	NW Vancouver Island Troll
      PSCFishery(175) = 4  '-SW VI  Trl	SW Vancouver Island Troll
      PSCFishery(176) = 9  '-GeoStr Trl	Georgia Straits Troll
      PSCFishery(177) = 12 '-BC JDF Trl	BC Juan de Fuca Troll
      PSCFishery(178) = 2  '-No BC  Net	BC Northern Net
      PSCFishery(179) = 2  '-Cen BC Net	BC Central Net
      PSCFishery(180) = 5  '-NW VI  Net	NW Vancouver Island Net
      PSCFishery(181) = 5  '-SW VI  Net	SW Vancouver Island Net
      PSCFishery(182) = 7  '-Johnst Net	Johnstone Straits Net
      PSCFishery(183) = 10 '-GeoStr Net	Georgia Straits Net
      PSCFishery(184) = 13 '-Fraser Net	Fraser R Gill Net
      PSCFishery(185) = 12 '-BC JDF Net	BC Juan de Fuca Net
      PSCFishery(186) = 8  '-JStrBC Spt	Johnstone Strait Sport
      PSCFishery(187) = 3  '-No BC  Spt	BC Northern Sport
      PSCFishery(188) = 3  '-Cen BC Spt	BC Central Sport
      PSCFishery(189) = 11 '-BC JDF Spt	BC Juan de Fuca Sport
      PSCFishery(190) = 6  '-WC VI  Spt	West Coast Vanc Is Sport
      PSCFishery(191) = 9  '-NGaStr Spt	North Georgia Straits Sport
      PSCFishery(192) = 9  '-SGaStr Spt	South Georgia Straits Sport
      PSCFishery(193) = 9  '-Albern Spt	Alberni Canal Sport
      PSCFishery(194) = 15 '-SW AK  Trl	SEAK Southwest Troll
      PSCFishery(195) = 15 '-SE AK  Trl	SEAK Southeast Troll
      PSCFishery(196) = 15 '-NW AK  Trl	SEAK Northwest Troll
      PSCFishery(197) = 15 '-NE AK  Trl	SEAK Northeast Troll
      PSCFishery(198) = 15 '-Alaska Net	Southeast Alaska Net

      PSCFishName(0) = "Fishery Name"
      PSCFishName(1) = "BC No/Cent Troll"
      PSCFishName(2) = "BC No/Cent Net"
      PSCFishName(3) = "BC No/Cent Sport"
      PSCFishName(4) = "BC WCVI Troll"
      PSCFishName(5) = "BC WCVI Net"
      PSCFishName(6) = "BC WCVI Sport"
      PSCFishName(7) = "BC JnStr Net&Trl"
      PSCFishName(8) = "BC JnstStr Sport"
      PSCFishName(9) = "BC GeStr Spt&Trl"
      PSCFishName(10) = "BC GeoStr Net"
      PSCFishName(11) = "BC JDF Sport"
      PSCFishName(12) = "BC JDF Net&Trl"
      PSCFishName(13) = "BC Frasr Net&Spt"
      PSCFishName(14) = "BC Sub-Total"

      PSCFishName(15) = "SEAK All"
      PSCFishName(16) = "WA Ocn Troll"
      PSCFishName(17) = "WA Ocn Sport"
      PSCFishName(18) = "SOF All"
      PSCFishName(19) = "US JDF All"
      PSCFishName(20) = "SanJnIsl Net"
      PSCFishName(21) = "SanJnIsl Sport"
      PSCFishName(22) = "PS Sport (8-13)"
      PSCFishName(23) = "PS Net (8-13)"
      PSCFishName(24) = "FW Net & Sport"
      PSCFishName(25) = "US Sub-Total"
      PSCFishName(26) = "TOTAL"
      PSCFishName(27) = "Cohort Size"
      PSCFishName(28) = "Escapement"

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      '- This Spreadsheet is for PSC Coho Periodic Report
      PSCSSName = FVSdatabasepath & "\PSCCohoReportData.XLSX"
      '- This Spreadsheet is for Bill Gazey's PSC Coho Report
      'PSCSSName = FVSdatabasepath & "\PSCCohoReportData-UNMARKED.XLSX"

      '- Test if PSC Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(PSCSSName).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(PSCSSName)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:
      xlWorkSheet = xlWorkBook.Sheets("Skagit")
      xlApp.Application.DisplayAlerts = False

      '- Loop through PSC Coho Stocks for Catch, Cohort, & Escapement

      'For Stock = 1 To 16 sixteen stocks for Bill Gazey project
      For Stock = 1 To 13
         ReDim PSCCatch(27, 30)
         ReDim PSCBase(NumFish, NumSteps)
         CmdStr = "SELECT RunID.RunYear, Mortality.StockID, Mortality.FisheryID, Mortality.TimeStep," & _
            " Mortality.LandedCatch,Mortality.NonRetention, Mortality.Shaker, Mortality.DropOff," & _
            " Mortality.MSFLandedCatch,Mortality.MSFNonRetention, Mortality.MSFShaker, Mortality.MSFDropOff" & _
            " FROM Mortality INNER JOIN RunID ON Mortality.RunID = RunID.RunID" & _
            " WHERE ("
         For Stk = 1 To PSCStock(Stock, 0)
            CmdStr &= "(Mortality.StockID)=" & PSCStock(Stock, Stk).ToString & " OR (Mortality.StockID)=" & (PSCStock(Stock, Stk) + 1).ToString
            If Stk = PSCStock(Stock, 0) Then
               CmdStr &= ") ORDER BY RunID.RunYear;"
               Exit For
            Else
               CmdStr &= " OR "
            End If
         Next

         Dim Mcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim MortalityDA As New System.Data.OleDb.OleDbDataAdapter
         MortalityDA.SelectCommand = Mcm
         Dim Mcb As New OleDb.OleDbCommandBuilder
         Mcb = New OleDb.OleDbCommandBuilder(MortalityDA)
         If FramDataSet.Tables.Contains("PSCMort") Then
            FramDataSet.Tables("PSCMort").Clear()
         End If
         MortalityDA.Fill(FramDataSet, "PSCMort")
         Dim NumM As Integer
         NumM = FramDataSet.Tables("PSCMort").Rows.Count
         '- Loop through Table Records for Actual Values
         For RecNum = 0 To NumM - 1
            RunYear = FramDataSet.Tables("PSCMort").Rows(RecNum)(0)
            Fish = FramDataSet.Tables("PSCMort").Rows(RecNum)(2)
            TStep = FramDataSet.Tables("PSCMort").Rows(RecNum)(3)
            LCat = FramDataSet.Tables("PSCMort").Rows(RecNum)(4)
            CNR = FramDataSet.Tables("PSCMort").Rows(RecNum)(5)
            Shak = FramDataSet.Tables("PSCMort").Rows(RecNum)(6)
            Drop = FramDataSet.Tables("PSCMort").Rows(RecNum)(7)
            MLCat = FramDataSet.Tables("PSCMort").Rows(RecNum)(8)
            MCNR = FramDataSet.Tables("PSCMort").Rows(RecNum)(9)
            MShak = FramDataSet.Tables("PSCMort").Rows(RecNum)(10)
            MDrop = FramDataSet.Tables("PSCMort").Rows(RecNum)(11)
            PSCCatch(PSCFishery(Fish) - 1, RunYear - 1985) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
            '- Sub-Total Lines
            If PSCFishery(Fish) < 14 Then
               PSCCatch(13, RunYear - 1985) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
            Else
               PSCCatch(24, RunYear - 1985) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
            End If
            '- TOTALS Line
            PSCCatch(25, RunYear - 1985) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop

            '- Coho Base
            If RunYear = 1985 Then
               PSCBase(Fish - 1, TStep - 1) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               PSCBase(Fish - 1, NumSteps) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               PSCBase(NumFish, TStep - 1) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               PSCBase(NumFish, NumSteps) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
            End If

         Next
         MortalityDA = Nothing

         '- Read Cohort Data
         CmdStr = "SELECT RunID.RunYear, Cohort.StockID, Cohort.TimeStep, Cohort.StartCohort" & _
            " FROM Cohort INNER JOIN RunID ON Cohort.RunID = RunID.RunID" & _
            " WHERE(("
         For Stk = 1 To PSCStock(Stock, 0)
            CmdStr &= "(Cohort.StockID)=" & PSCStock(Stock, Stk).ToString & " OR (Cohort.StockID)=" & (PSCStock(Stock, Stk) + 1).ToString
            If Stk = PSCStock(Stock, 0) Then
               CmdStr &= ") And (Cohort.TimeStep = 1));"
               Exit For
            Else
               CmdStr &= " OR "
            End If
         Next

         Dim CScm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim CohortDA As New System.Data.OleDb.OleDbDataAdapter
         CohortDA.SelectCommand = CScm
         Dim CScb As New OleDb.OleDbCommandBuilder
         CScb = New OleDb.OleDbCommandBuilder(CohortDA)
         If FramDataSet.Tables.Contains("PSCCohort") Then
            FramDataSet.Tables("PSCCohort").Clear()
         End If
         CohortDA.Fill(FramDataSet, "PSCCohort")
         Dim NumCS As Integer
         NumCS = FramDataSet.Tables("PSCCohort").Rows.Count
         '- Loop through Table Records for Actual Values
         For RecNum = 0 To NumCS - 1
            RunYear = FramDataSet.Tables("PSCCohort").Rows(RecNum)(0)
            Stk = FramDataSet.Tables("PSCCohort").Rows(RecNum)(1)
            TStep = FramDataSet.Tables("PSCCohort").Rows(RecNum)(2)
            InitCoht = FramDataSet.Tables("PSCCohort").Rows(RecNum)(3)
            PSCCatch(26, RunYear - 1985) += InitCoht
         Next
         CohortDA = Nothing

         '- Read Escapement Data
         CmdStr = "SELECT Escapement.StockID, Escapement.Escapement, RunID.RunYear" & _
            " FROM Escapement INNER JOIN RunID ON Escapement.RunID = RunID.RunID" & _
            " WHERE("
         For Stk = 1 To PSCStock(Stock, 0)
            CmdStr &= "(Escapement.StockID)=" & PSCStock(Stock, Stk).ToString & " OR (Escapement.StockID)=" & (PSCStock(Stock, Stk) + 1).ToString
            If Stk = PSCStock(Stock, 0) Then
               CmdStr &= ") ORDER BY RunID.RunYear;"
               Exit For
            Else
               CmdStr &= " OR "
            End If
         Next
         Dim EScm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim EscapementDA As New System.Data.OleDb.OleDbDataAdapter
         EscapementDA.SelectCommand = EScm
         Dim EScb As New OleDb.OleDbCommandBuilder
         EScb = New OleDb.OleDbCommandBuilder(EscapementDA)
         If FramDataSet.Tables.Contains("PSCEsc") Then
            FramDataSet.Tables("PSCEsc").Clear()
         End If
         EscapementDA.Fill(FramDataSet, "PSCEsc")
         Dim NumES As Integer
         NumES = FramDataSet.Tables("PSCEsc").Rows.Count

         If Stock = 11 Then
            Jim = 1
         End If

         '- Loop through Table Records for Actual Values
         For RecNum = 0 To NumES - 1
            RunYear = FramDataSet.Tables("PSCEsc").Rows(RecNum)(2)
            Esc = FramDataSet.Tables("PSCEsc").Rows(RecNum)(1)
            PSCCatch(27, RunYear - 1985) += Esc
         Next
         EscapementDA = Nothing

         '- Put PSCCatch Array into PSC Spreadsheet by PSC Stock

         Select Case Stock
            Case 1
               xlWorkSheet = xlWorkBook.Sheets("Skagit")
            Case 2
               xlWorkSheet = xlWorkBook.Sheets("Stillaguamish")
            Case 3
               xlWorkSheet = xlWorkBook.Sheets("Snohomish")
            Case 4
               xlWorkSheet = xlWorkBook.Sheets("Hood Canal")
            Case 5
               xlWorkSheet = xlWorkBook.Sheets("USJDF")
            Case 6
               xlWorkSheet = xlWorkBook.Sheets("Quillayute")
            Case 7
               xlWorkSheet = xlWorkBook.Sheets("Hoh")
            Case 8
               xlWorkSheet = xlWorkBook.Sheets("Queets")
            Case 9
               xlWorkSheet = xlWorkBook.Sheets("Grays Harbor")
            Case 10
               xlWorkSheet = xlWorkBook.Sheets("Lower Fraser")
            Case 11
               xlWorkSheet = xlWorkBook.Sheets("Upper Fraser")
            Case 12
               xlWorkSheet = xlWorkBook.Sheets("GS Mainland")
            Case 13
               xlWorkSheet = xlWorkBook.Sheets("GS VancIsl")
            Case 14
               xlWorkSheet = xlWorkBook.Sheets("SW VancIsl")
            Case 15
               xlWorkSheet = xlWorkBook.Sheets("Col Riv Early")
            Case 16
               xlWorkSheet = xlWorkBook.Sheets("Col Riv Late")
         End Select

         'Transfer array to the worksheet starting at cell A2.
         xlWorkSheet.Range("B2").Resize(28, 30).Value = PSCCatch
         'xlWorkSheet.Range("A1:A24").Resize(24).Value = PSCFishName
         Dim CellAddress As String
         For Fish As Integer = 1 To 28
            CellAddress = "A" & (Fish + 1).ToString
            xlWorkSheet.Range(CellAddress).Value = PSCFishName(Fish)
         Next

         '- Transfer Base ER's
         Dim NumERFish As Integer
         NumERFish = 0
         '- Find Number of Fisheries with ER's
         For Fish As Integer = 1 To NumFish
            If PSCBase(Fish - 1, NumSteps) > 0 Then
               NumERFish += 1
            End If
         Next
         ReDim PSCBaseER(NumERFish + 1, NumSteps)
         NumERFish = 0
         '- Copy Fishery ER's
         NumERFish = 0
         For Fish As Integer = 1 To NumFish + 1
            If PSCBase(Fish - 1, NumSteps) > 0 Then
               For TStep As Integer = 0 To NumSteps
                  PSCBaseER(NumERFish, TStep) = PSCBase(Fish - 1, TStep)
                  '- Divide by Total Mort plus Escapement
                  PSCBaseER(NumERFish, TStep) = PSCBaseER(NumERFish, TStep) / (PSCCatch(25, 0) + PSCCatch(27, 0))
               Next
               NumERFish += 1
            End If
         Next
         '- 
         'Transfer array to the worksheet starting at cell B65.
         xlWorkSheet.Range("A65:AE265").ClearContents()
         xlWorkSheet.Range("B65").Resize(NumERFish, NumSteps + 1).Value = PSCBaseER
         'Copy Fishery Names to Spreadsheet
         NumERFish = 0
         For Fish As Integer = 1 To NumFish
            If PSCBase(Fish - 1, NumSteps) > 0 Then
               CellAddress = "A" & (NumERFish + 65).ToString
               xlWorkSheet.Range(CellAddress).Value = FisheryTitle(Fish)
               NumERFish += 1
            End If
         Next
         CellAddress = "A" & (NumERFish + 65).ToString
         xlWorkSheet.Range(CellAddress).Value = "Total"

      Next

      '- Call PSCCohoHatcheryInterception then return to Close Spreadsheet
      PSCCohoHatcheryInterception()

      '- Done with PSC WorkBook for this run .. Close and release object
      xlApp.Application.DisplayAlerts = False
      xlWorkBook.Save()
      If WorkBookWasNotOpen = True Then
         xlWorkBook.Close()
      End If
      If ExcelWasNotRunning = True Then
         xlApp.Application.Quit()
         xlApp.Quit()
      Else
         xlApp.Visible = True
         xlApp.WindowState = Excel.XlWindowState.xlMinimized
      End If
      xlApp.Application.DisplayAlerts = True
      xlApp = Nothing

   End Sub


   Sub PSCCohoHatcheryInterception()

      Dim PSCCohoHatchery(NumStk) As Integer
      Dim RecNum, RunYear As Integer
      Dim PSCIntercept(30, 4) As Double
      Dim LCat, CNR, Drop, Shak As Double
      Dim MLCat, MCNR, MDrop, MShak As Double
      Dim CmdStr As String

      '- Array of All Stocks  1 = US Hatchery  2 = Canadian Hatchery Stocks
      PSCCohoHatchery(1) = 0  '- U-nkskrw	Nooksack River Wild UnMarked
      PSCCohoHatchery(2) = 0  '- M-nkskrw	Nooksack River Wild Marked
      PSCCohoHatchery(3) = 1  '- U-kendlh	Kendall Creek Hatchery UnMarked
      PSCCohoHatchery(4) = 1  '- M-kendlh	Kendall Creek Hatchery Marked
      PSCCohoHatchery(5) = 1  '- U-skokmh	Skookum Creek Hatchery UnMarked
      PSCCohoHatchery(6) = 1  '- M-skokmh	Skookum Creek Hatchery Marked
      PSCCohoHatchery(7) = 1  '- U-lumpdh	Lummi Ponds Hatchery UnMarked
      PSCCohoHatchery(8) = 1  '- M-lumpdh	Lummi Ponds Hatchery Marked
      PSCCohoHatchery(9) = 1  '- U-bhambh	Bellingham Bay Net Pens UnMarked
      PSCCohoHatchery(10) = 1  '- M-bhambh	Bellingham Bay Net Pens Marked
      PSCCohoHatchery(11) = 0  '- U-samshw	Samish River Wild UnMarked
      PSCCohoHatchery(12) = 0  '- M-samshw	Samish River Wild Marked
      PSCCohoHatchery(13) = 0  '- U-ar77aw	Area 7/7A Independent Wild UnMarked
      PSCCohoHatchery(14) = 0  '- M-ar77aw	Area 7/7A Independent Wild Marked
      PSCCohoHatchery(15) = 1  '- U-whatch	Whatcom Creek Hatchery UnMarked
      PSCCohoHatchery(16) = 1  '- M-whatch	Whatcom Creek Hatchery Marked
      PSCCohoHatchery(17) = 0  '- U-skagtw	Skagit River Wild UnMarked
      PSCCohoHatchery(18) = 0  '- M-skagtw	Skagit River Wild Marked
      PSCCohoHatchery(19) = 1  '- U-skagth	Skagit River Hatchery UnMarked
      PSCCohoHatchery(20) = 1  '- M-skagth	Skagit River Hatchery Marked
      PSCCohoHatchery(21) = 1  '- U-skgbkh	Baker (Skagit) Hatchery UnMarked
      PSCCohoHatchery(22) = 1  '- M-skgbkh	Baker (Skagit) Hatchery Marked
      PSCCohoHatchery(23) = 0  '- U-skgbkw	Baker (Skagit) Wild UnMarked
      PSCCohoHatchery(24) = 0  '- U-skgbkw	Baker (Skagit) Wild UnMarked
      PSCCohoHatchery(25) = 1  '- U-swinch	Swinomish Channel Hatchery UnMarked
      PSCCohoHatchery(26) = 1  '- M-swinch	Swinomish Channel Hatchery Marked
      PSCCohoHatchery(27) = 1  '- U-oakhbh	Oak Harbor Net Pens UnMarked
      PSCCohoHatchery(28) = 1  '- M-oakhbh	Oak Harbor Net Pens Marked
      PSCCohoHatchery(29) = 0  '- U-stillw	Stillaguamish River Wild UnMarked
      PSCCohoHatchery(30) = 0  '- M-stillw	Stillaguamish River Wild Marked
      PSCCohoHatchery(31) = 1  '- U-stillh	Stillaguamish River Hatchery UnMarked
      PSCCohoHatchery(32) = 1  '- M-stillh	Stillaguamish River Hatchery Marked
      PSCCohoHatchery(33) = 1  '- U-tuliph	Tulalip Hatchery UnMarked
      PSCCohoHatchery(34) = 1  '- M-tuliph	Tulalip Hatchery Marked
      PSCCohoHatchery(35) = 0  '- U-snohow	Snohomish River Wild UnMarked
      PSCCohoHatchery(36) = 0  '- M-snohow	Snohomish River Wild Marked
      PSCCohoHatchery(37) = 1  '- U-snohoh	Snohomish River Hatchery UnMarked
      PSCCohoHatchery(38) = 1  '- M-snohoh	Snohomish River Hatchery Marked
      PSCCohoHatchery(39) = 1  '- U-ar8anh	Area 8A Net Pens UnMarked
      PSCCohoHatchery(40) = 1  '- M-ar8anh	Area 8A Net Pens Marked
      PSCCohoHatchery(41) = 1  '- U-ptgamh	Port Gamble Net Pens UnMarked
      PSCCohoHatchery(42) = 1  '- M-ptgamh	Port Gamble Net Pens Marked
      PSCCohoHatchery(43) = 0  '- U-ptgamw	Port Gamble Bay Wild UnMarked
      PSCCohoHatchery(44) = 0  '- M-ptgamw	Port Gamble Bay Wild Marked
      PSCCohoHatchery(45) = 0  '- U-ar12bw	Area 12/12B Wild UnMarked
      PSCCohoHatchery(46) = 0  '- M-ar12bw	Area 12/12B Wild Marked
      PSCCohoHatchery(47) = 1  '- U-qlcnbh	Quilcene Hatchery UnMarked
      PSCCohoHatchery(48) = 1  '- M-qlcnbh	Quilcene Hatchery Marked
      PSCCohoHatchery(49) = 0  '- U-qlcenh	Quilcene Bay Net Pens UnMarked
      PSCCohoHatchery(50) = 0  '- M-qlcenh	Quilcene Bay Net Pens Marked
      PSCCohoHatchery(51) = 0  '- U-ar12aw	Area 12A Wild UnMarked
      PSCCohoHatchery(52) = 0  '- M-ar12aw	Area 12A Wild Marked
      PSCCohoHatchery(53) = 1  '- U-hoodsh	Hoodsport Hatchery UnMarked
      PSCCohoHatchery(54) = 1  '- M-hoodsh	Hoodsport Hatchery Marked
      PSCCohoHatchery(55) = 0  '- U-ar12dw	Area 12C/12D Wild UnMarked
      PSCCohoHatchery(56) = 0  '- M-ar12dw	Area 12C/12D Wild Marked
      PSCCohoHatchery(57) = 1  '- U-gadamh	George Adams Hatchery UnMarked
      PSCCohoHatchery(58) = 1  '- M-gadamh	George Adams Hatchery Marked
      PSCCohoHatchery(59) = 0  '- U-skokrw	Skokomish River Wild UnMarked
      PSCCohoHatchery(60) = 0  '- M-skokrw	Skokomish River Wild Marked
      PSCCohoHatchery(61) = 0  '- U-ar13bw	Area 13B Miscellaneous Wild UnMarked
      PSCCohoHatchery(62) = 0  '- M-ar13bw	Area 13B Miscellaneous Wild Marked
      PSCCohoHatchery(63) = 0  '- U-deschw	Deschutes River (WA) Wild UnMarked
      PSCCohoHatchery(64) = 0  '- M-deschw	Deschutes River (WA) Wild Marked
      PSCCohoHatchery(65) = 1  '- U-ssdnph	South Puget Sound Net Pens UnMarked
      PSCCohoHatchery(66) = 1  '- M-ssdnph	South Puget Sound Net Pens Marked
      PSCCohoHatchery(67) = 1  '- U-nisqlh	Nisqually River Hatchery UnMarked
      PSCCohoHatchery(68) = 1  '- M-nisqlh	Nisqually River Hatchery Marked
      PSCCohoHatchery(69) = 0  '- U-nisqlw	Nisqually River Wild UnMarked
      PSCCohoHatchery(70) = 0  '- M-nisqlw	Nisqually River Wild Marked
      PSCCohoHatchery(71) = 1  '- U-foxish	Fox Island Net Pens UnMarked
      PSCCohoHatchery(72) = 1  '- M-foxish	Fox Island Net Pens Marked
      PSCCohoHatchery(73) = 1  '- U-mintch	Minter Creek Hatchery UnMarked
      PSCCohoHatchery(74) = 1  '- M-mintch	Minter Creek Hatchery Marked
      PSCCohoHatchery(75) = 0  '- U-ar13mw	Area 13 Miscellaneous Wild UnMarked
      PSCCohoHatchery(76) = 0  '- M-ar13mw	Area 13 Miscellaneous Wild Marked
      PSCCohoHatchery(77) = 1  '- U-chambh	Chambers Creek Hatchery UnMarked
      PSCCohoHatchery(78) = 1  '- M-chambh	Chambers Creek Hatchery Marked
      PSCCohoHatchery(79) = 1  '- U-ar13mh	Area 13 Miscellaneous Hatchery UnMarked
      PSCCohoHatchery(80) = 1  '- M-ar13mh	Area 13 Miscellaneous Hatchery Marked
      PSCCohoHatchery(81) = 0  '- U-ar13aw	Area 13A Miscellaneous Wild UnMarked
      PSCCohoHatchery(82) = 0  '- M-ar13aw	Area 13A Miscellaneous Wild Marked
      PSCCohoHatchery(83) = 1  '- U-puyalh	Puyallup River Hatchery UnMarked
      PSCCohoHatchery(84) = 1  '- M-puyalh	Puyallup River Hatchery Marked
      PSCCohoHatchery(85) = 0  '- U-puyalw	Puyallup River Wild UnMarked
      PSCCohoHatchery(86) = 0  '- M-puyalw	Puyallup River Wild Marked
      PSCCohoHatchery(87) = 1  '- U-are11h	Area 11 Hatchery UnMarked
      PSCCohoHatchery(88) = 1  '- M-are11h	Area 11 Hatchery Marked
      PSCCohoHatchery(89) = 0  '- U-ar11mw	Area 11 Miscellaneous Wild UnMarked
      PSCCohoHatchery(90) = 0  '- M-ar11mw	Area 11 Miscellaneous Wild Marked
      PSCCohoHatchery(91) = 1  '- U-ar10eh	Area 10E Hatchery UnMarked
      PSCCohoHatchery(92) = 1  '- M-ar10eh	Area 10E Hatchery Marked
      PSCCohoHatchery(93) = 0  '- U-ar10ew	Area 10E Miscellaneous Wild UnMarked
      PSCCohoHatchery(94) = 0  '- M-ar10ew	Area 10E Miscellaneous Wild Marked
      PSCCohoHatchery(95) = 1  '- U-greenh	Green River Hatchery UnMarked
      PSCCohoHatchery(96) = 1  '- M-greenh	Green River Hatchery Marked
      PSCCohoHatchery(97) = 0  '- U-greenw	Green River Wild UnMarked
      PSCCohoHatchery(98) = 0  '- M-greenw	Green River Wild Marked
      PSCCohoHatchery(99) = 1  '- U-lakwah	Lake Washington Hatchery UnMarked
      PSCCohoHatchery(100) = 1  '- M-lakwah	Lake Washington Hatchery Marked
      PSCCohoHatchery(101) = 0  '- U-lakwaw	Lake Washington Wild UnMarked
      PSCCohoHatchery(102) = 0  '- M-lakwaw	Lake Washington Wild Marked
      PSCCohoHatchery(103) = 1  '- U-are10h	Area 10 Hatchery UnMarked
      PSCCohoHatchery(104) = 1  '- M-are10h	Area 10 Hatchery Marked
      PSCCohoHatchery(105) = 0  '- U-ar10mw	Area 10 Miscellaneous Wild UnMarked
      PSCCohoHatchery(106) = 0  '- M-ar10mw	Area 10 Miscellaneous Wild Marked
      PSCCohoHatchery(107) = 0  '- U-dungew	Dungeness River Wild UnMarked
      PSCCohoHatchery(108) = 0  '- M-dungew	Dungeness River Wild Marked
      PSCCohoHatchery(109) = 1  '- U-dungeh	Dungeness Hatchery UnMarked
      PSCCohoHatchery(110) = 1  '- M-dungeh	Dungeness Hatchery Marked
      PSCCohoHatchery(111) = 0  '- U-elwhaw	Elwha River Wild UnMarked
      PSCCohoHatchery(112) = 0  '- M-elwhaw	Elwha River Wild Marked
      PSCCohoHatchery(113) = 1  '- U-elwhah	Elwha Hatchery UnMarked
      PSCCohoHatchery(114) = 1  '- M-elwhah	Elwha Hatchery Marked
      PSCCohoHatchery(115) = 0  '- U-ejdfmw	East JDF Miscellaneous Wild UnMarked
      PSCCohoHatchery(116) = 0  '- M-ejdfmw	East JDF Miscellaneous Wild Marked
      PSCCohoHatchery(117) = 0  '- U-wjdfmw	West JDF Miscellaneous Wild UnMarked
      PSCCohoHatchery(118) = 0  '- M-wjdfmw	West JDF Miscellaneous Wild Marked
      PSCCohoHatchery(119) = 1  '- U-ptangh	Port Angeles Net Pens UnMarked
      PSCCohoHatchery(120) = 1  '- M-ptangh	Port Angeles Net Pens Marked
      PSCCohoHatchery(121) = 0  '- U-area9w	Area 9 Miscellaneous Wild UnMarked
      PSCCohoHatchery(122) = 0  '- M-area9w	Area 9 Miscellaneous Wild Marked
      PSCCohoHatchery(123) = 0  '- U-makahw	Makah Coastal Wild UnMarked
      PSCCohoHatchery(124) = 0  '- M-makahw	Makah Coastal Wild Marked
      PSCCohoHatchery(125) = 1  '- U-makahh	Makah Coastal Hatchery UnMarked
      PSCCohoHatchery(126) = 1  '- M-makahh	Makah Coastal Hatchery Marked
      PSCCohoHatchery(127) = 0  '- U-quilsw	Quillayute River Summer Natural UnMarked
      PSCCohoHatchery(128) = 0  '- M-quilsw	Quillayute River Summer Natural Marked
      PSCCohoHatchery(129) = 1  '- U-quilsh	Quillayute River Summer Hatchery UnMarked
      PSCCohoHatchery(130) = 1  '- M-quilsh	Quillayute River Summer Hatchery Marked
      PSCCohoHatchery(131) = 0  '- U-quilfw	Quillayute River Fall Natural UnMarked
      PSCCohoHatchery(132) = 0  '- M-quilfw	Quillayute River Fall Natural Marked
      PSCCohoHatchery(133) = 1  '- U-quilfh	Quillayute River Fall Hatchery UnMarked
      PSCCohoHatchery(134) = 1  '- M-quilfh	Quillayute River Fall Hatchery Marked
      PSCCohoHatchery(135) = 0  '- U-hohrvw	Hoh River Wild UnMarked
      PSCCohoHatchery(136) = 0  '- M-hohrvw	Hoh River Wild Marked
      PSCCohoHatchery(137) = 1  '- U-hohrvh	Hoh River Hatchery UnMarked
      PSCCohoHatchery(138) = 1  '- M-hohrvh	Hoh River Hatchery Marked
      PSCCohoHatchery(139) = 0  '- U-quetfw	Queets River Fall Natural UnMarked
      PSCCohoHatchery(140) = 0  '- M-quetfw	Queets River Fall Natural Marked
      PSCCohoHatchery(141) = 1  '- U-quetfh	Queets River Fall Hatchery UnMarked
      PSCCohoHatchery(142) = 1  '- M-quetfh	Queets River Fall Hatchery Marked
      PSCCohoHatchery(143) = 1  '- U-quetph	Queets River Suppl. Hatchery UnMarked
      PSCCohoHatchery(144) = 1  '- M-quetph	Queets River Suppl. Hatchery Marked
      PSCCohoHatchery(145) = 0  '- U-quinfw	Quinault River Fall Natural UnMarked
      PSCCohoHatchery(146) = 0  '- M-quinfw	Quinault River Fall Natural Marked
      PSCCohoHatchery(147) = 1  '- U-quinfh	Quinault River Fall Hatchery UnMarked
      PSCCohoHatchery(148) = 1  '- M-quinfh	Quinault River Fall Hatchery Marked
      PSCCohoHatchery(149) = 0  '- U-chehlw	Chehalis River Wild UnMarked
      PSCCohoHatchery(150) = 0  '- M-chehlw	Chehalis River Wild Marked
      PSCCohoHatchery(151) = 1  '- U-chehlh	Chehalis River Hatchery UnMarked
      PSCCohoHatchery(152) = 1  '- M-chehlh	Chehalis River Hatchery Marked
      PSCCohoHatchery(153) = 0  '- U-humptw	Humptulips River Wild UnMarked
      PSCCohoHatchery(154) = 0  '- M-humptw	Humptulips River Wild Marked
      PSCCohoHatchery(155) = 1  '- U-humpth	Humptulips River Hatchery UnMarked
      PSCCohoHatchery(156) = 1  '- M-humpth	Humptulips River Hatchery Marked
      PSCCohoHatchery(157) = 0  '- U-gryhmw	Grays Harbor Miscellaneous Wild UnMarked
      PSCCohoHatchery(158) = 0  '- M-gryhmw	Grays Harbor Miscellaneous Wild Marked
      PSCCohoHatchery(159) = 1  '- U-gryhbh	Grays Harbor Net Pens UnMarked
      PSCCohoHatchery(160) = 1  '- M-gryhbh	Grays Harbor Net Pens Marked
      PSCCohoHatchery(161) = 0  '- U-willaw	Willapa Bay Natural UnMarked
      PSCCohoHatchery(162) = 0  '- M-willaw	Willapa Bay Natural Marked
      PSCCohoHatchery(163) = 1  '- U-willah	Willapa Bay Hatchery UnMarked
      PSCCohoHatchery(164) = 1  '- M-willah	Willapa Bay Hatchery Marked
      PSCCohoHatchery(165) = 1  '- U-colreh	Columbia River Early Hatchery UnMarked
      PSCCohoHatchery(166) = 1  '- M-colreh	Columbia River Early Hatchery Marked
      PSCCohoHatchery(167) = 1  '- U-youngh	Youngs Bay Hatchery UnMarked
      PSCCohoHatchery(168) = 1  '- M-youngh	Youngs Bay Hatchery Marked
      PSCCohoHatchery(169) = 0  '- U-crorew	Lower Col R Oregon Wild UnMarked
      PSCCohoHatchery(170) = 0  '- M-crorew	Lower Col R Oregon Wild Marked
      PSCCohoHatchery(171) = 0  '- U-washew	Wash Early Wild UnMarked
      PSCCohoHatchery(172) = 0  '- M-washew	Wash Early Wild Marked
      PSCCohoHatchery(173) = 0  '- U-washlw	Wash Late Wild UnMarked
      PSCCohoHatchery(174) = 0  '- M-washlw	Wash Late Wild Marked
      PSCCohoHatchery(175) = 1  '- U-colrlh	Columbia River Late Hatchery UnMarked
      PSCCohoHatchery(176) = 1  '- M-colrlh	Columbia River Late Hatchery Marked
      PSCCohoHatchery(177) = 1  '- U-orenoh	Oregon North Coast Hatchery UnMarked
      PSCCohoHatchery(178) = 1  '- M-orenoh	Oregon North Coast Hatchery Marked
      PSCCohoHatchery(179) = 0  '- U-orenow	Oregon North Coast Wild UnMarked
      PSCCohoHatchery(180) = 0  '- M-orenow	Oregon North Coast Wild Marked
      PSCCohoHatchery(181) = 1  '- U-orenmh	Oregon North-Mid Coast Hatchery UnMarked
      PSCCohoHatchery(182) = 1  '- M-orenmh	Oregon North-Mid Coast Hatchery Marked
      PSCCohoHatchery(183) = 0  '- U-orenmw	Oregon North-Mid Coast Wild UnMarked
      PSCCohoHatchery(184) = 0  '- M-orenmw	Oregon North-Mid Coast Wild Marked
      PSCCohoHatchery(185) = 1  '- U-oresmh	Oregon South-Mid Coast Hatchery UnMarked
      PSCCohoHatchery(186) = 1  '- M-oresmh	Oregon South-Mid Coast Hatchery Marked
      PSCCohoHatchery(187) = 0  '- U-oresmw	Oregon South-Mid Coast Wild UnMarked
      PSCCohoHatchery(188) = 0  '- M-oresmw	Oregon South-Mid Coast Wild Marked
      PSCCohoHatchery(189) = 1  '- U-oranah	Oregon Anadromous Hatchery UnMarked
      PSCCohoHatchery(190) = 1  '- M-oranah	Oregon Anadromous Hatchery Marked
      PSCCohoHatchery(191) = 1  '- U-oraqah	Oregon Aqua-Foods Hatchery UnMarked
      PSCCohoHatchery(192) = 1  '- M-oraqah	Oregon Aqua-Foods Hatchery Marked
      PSCCohoHatchery(193) = 1  '- U-oresoh	Oregon South Coast Hatchery UnMarked
      PSCCohoHatchery(194) = 1  '- M-oresoh	Oregon South Coast Hatchery Marked
      PSCCohoHatchery(195) = 0  '- U-oresow	Oregon South Coast Wild UnMarked
      PSCCohoHatchery(196) = 0  '- M-oresow	Oregon South Coast Wild Marked
      PSCCohoHatchery(197) = 1  '- U-calnoh	California North Coast Hatchery UnMarked
      PSCCohoHatchery(198) = 1  '- M-calnoh	California North Coast Hatchery Marked
      PSCCohoHatchery(199) = 0  '- U-calnow	California North Coast Wild UnMarked
      PSCCohoHatchery(200) = 0  '- M-calnow	California North Coast Wild Marked
      PSCCohoHatchery(201) = 1  '- U-calcnh	California Central Coast Hatchery UnMarked
      PSCCohoHatchery(202) = 1  '- M-calcnh	California Central Coast Hatchery Marked
      PSCCohoHatchery(203) = 0  '- U-calcnw	California Central Coast Wild UnMarked
      PSCCohoHatchery(204) = 0  '- M-calcnw	California Central Coast Wild Marked
      PSCCohoHatchery(205) = 2  '- U-gsmndh	Georgia Strait Mainland Hatchery UnMarked
      PSCCohoHatchery(206) = 2  '- M-gsmndh	Georgia Strait Mainland Hatchery Marked
      PSCCohoHatchery(207) = 0  '- U-gsmndw	Georgia Strait Mainland Wild UnMarked
      PSCCohoHatchery(208) = 0  '- M-gsmndw	Georgia Strait Mainland Wild Marked
      PSCCohoHatchery(209) = 2  '- U-gsvcih	Georgia Strait Vanc. Isl. Hatchery UnMarked
      PSCCohoHatchery(210) = 2  '- M-gsvcih	Georgia Strait Vanc. Isl. Hatchery Marked
      PSCCohoHatchery(211) = 0  '- U-gsvciw	Georgia Strait Vanc. Isl. Wild UnMarked
      PSCCohoHatchery(212) = 0  '- M-gsvciw	Georgia Strait Vanc. Isl. Wild Marked
      PSCCohoHatchery(213) = 2  '- U-jnstrh	Johnstone Strait Hatchery UnMarked
      PSCCohoHatchery(214) = 2  '- M-jnstrh	Johnstone Strait Hatchery Marked
      PSCCohoHatchery(215) = 0  '- U-jnstrw	Johnstone Strait Wild UnMarked
      PSCCohoHatchery(216) = 0  '- M-jnstrw	Johnstone Strait Wild Marked
      PSCCohoHatchery(217) = 2  '- U-swvcih	SW Vancouver Island Hatchery UnMarked
      PSCCohoHatchery(218) = 2  '- M-swvcih	SW Vancouver Island Hatchery Marked
      PSCCohoHatchery(219) = 0  '- U-swvciw	SW Vancouver Island Wild UnMarked
      PSCCohoHatchery(220) = 0  '- M-swvciw	SW Vancouver Island Wild Marked
      PSCCohoHatchery(221) = 2  '- U-nwvcih	NW Vancouver Island Hatchery UnMarked
      PSCCohoHatchery(222) = 2  '- M-nwvcih	NW Vancouver Island Hatchery Marked
      PSCCohoHatchery(223) = 0  '- U-nwvciw	NW Vancouver Island Wild UnMarked
      PSCCohoHatchery(224) = 0  '- M-nwvciw	NW Vancouver Island Wild Marked
      PSCCohoHatchery(225) = 2  '- U-frslwh	Lower Fraser River Hatchery UnMarked
      PSCCohoHatchery(226) = 2  '- M-frslwh	Lower Fraser River Hatchery Marked
      PSCCohoHatchery(227) = 0  '- U-frslww	Lower Fraser River Wild UnMarked
      PSCCohoHatchery(228) = 0  '- M-frslww	Lower Fraser River Wild Marked
      PSCCohoHatchery(229) = 2  '- U-frsuph	Upper Fraser River Hatchery UnMarked
      PSCCohoHatchery(230) = 2  '- M-frsuph	Upper Fraser River Hatchery Marked
      PSCCohoHatchery(231) = 0  '- U-frsupw	Upper Fraser River Wild UnMarked
      PSCCohoHatchery(232) = 0  '- M-frsupw	Upper Fraser River Wild Marked
      PSCCohoHatchery(233) = 0  '- U-bccnhw	BC Central Coast Hatchery/Wild UnMarked
      PSCCohoHatchery(234) = 0  '- M-bccnhw	BC Central Coast Hatchery/Wild Marked
      PSCCohoHatchery(235) = 0  '- U-bcnchw	BC North Coast Hatchery/Wild UnMarked
      PSCCohoHatchery(236) = 0  '- M-bcnchw	BC North Coast Hatchery/Wild Marked
      PSCCohoHatchery(237) = 0  '- U-tranhw	Trans Boundary Hatchery/Wild UnMarked
      PSCCohoHatchery(238) = 0  '- M-tranhw	Trans Boundary Hatchery/Wild Marked
      PSCCohoHatchery(239) = 0  '- U-niakhw	Alaska Northern Inside Hat/Wild UnMarked
      PSCCohoHatchery(240) = 0  '- M-niakhw	Alaska Northern Inside Hat/Wild Marked
      PSCCohoHatchery(241) = 0  '- U-noakhw	Alaska Northern Outside Hat/Wild UnMarked
      PSCCohoHatchery(242) = 0  '- M-noakhw	Alaska Northern Outside Hat/Wild Marked
      PSCCohoHatchery(243) = 0  '- U-siakhw	Alaska Southern Inside Hat/Wild UnMarked
      PSCCohoHatchery(244) = 0  '- M-siakhw	Alaska Southern Inside Hat/Wild Marked
      PSCCohoHatchery(245) = 0  '- U-soakhw	Alaska Southern Outside Hat/Wild UnMarked
      PSCCohoHatchery(246) = 0  '- M-soakhw	Alaska Southern Outside Hat/Wild Marked

      '- Fishery Numbers  167-193 are Canadian
      '- Stock Numbers  Odd Numbers are UnMarked, Even Numbers are Marked

      'CmdStr = "SELECT RunID.RunYear, Mortality.StockID, Mortality.FisheryID, Mortality.TimeStep, Mortality.LandedCatch," & _
      '   "Mortality.NonRetention, Mortality.Shaker, Mortality.LegalShaker, Mortality.DropOff" & _
      '   " FROM Mortality INNER JOIN RunID ON Mortality.RunID = RunID.RunID ORDER BY RunID.RunYear;"
      CmdStr = "SELECT RunID.RunYear, Mortality.StockID, Mortality.FisheryID, Mortality.TimeStep," & _
         " Mortality.LandedCatch,Mortality.NonRetention, Mortality.Shaker, Mortality.DropOff," & _
         " Mortality.MSFLandedCatch,Mortality.MSFNonRetention, Mortality.MSFShaker, Mortality.MSFDropOff" & _
         " FROM Mortality INNER JOIN RunID ON Mortality.RunID = RunID.RunID ORDER BY RunID.RunYear;"

      Dim Mcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim MortalityDA As New System.Data.OleDb.OleDbDataAdapter
      MortalityDA.SelectCommand = Mcm
      Dim Mcb As New OleDb.OleDbCommandBuilder
      Mcb = New OleDb.OleDbCommandBuilder(MortalityDA)
      If FramDataSet.Tables.Contains("PSCIntercept") Then
         FramDataSet.Tables("PSCIntercept").Clear()
      End If
      MortalityDA.Fill(FramDataSet, "PSCIntercept")
      Dim NumM As Integer
      NumM = FramDataSet.Tables("PSCIntercept").Rows.Count

      '- Loop through Table Records for Actual Values
      For RecNum = 0 To NumM - 1
         RunYear = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(0)
         Stk = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(1)
         Fish = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(2)
         TStep = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(3)
         LCat = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(4)
         CNR = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(5)
         Shak = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(6)
         Drop = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(7)
         MLCat = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(8)
         MCNR = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(9)
         MShak = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(10)
         MDrop = FramDataSet.Tables("PSCIntercept").Rows(RecNum)(11)

         If PSCCohoHatchery(Stk) = 0 Then GoTo NextCohoStock

         If PSCCohoHatchery(Stk) = 1 Then
            '- US Hatchery Fish
            If Fish >= 167 And Fish <= 193 Then
               '- Canadian Catch of US Hatchery Fish
               If Stk Mod 2 = 0 And RunYear > 1997 Then
                  '- Marked
                  PSCIntercept(RunYear - 1985, 2) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               Else
                  '- UnMarked
                  PSCIntercept(RunYear - 1985, 3) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               End If
            End If
         Else
            '- Canadian Hatchery Fish
            If Fish < 167 Or Fish > 193 Then
               '- US Catch of Canadian Hatchery Fish
               If Stk Mod 2 = 0 And RunYear > 1997 Then
                  '- Marked
                  PSCIntercept(RunYear - 1985, 0) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               Else
                  '- UnMarked
                  PSCIntercept(RunYear - 1985, 1) += LCat + CNR + Shak + Drop + MLCat + MCNR + MShak + MDrop
               End If
            End If
         End If

NextCohoStock:

      Next
      MortalityDA = Nothing

      xlWorkSheet = xlWorkBook.Sheets("Hatchery Intercept")

      'Transfer array to the worksheet starting at cell B3.
      xlWorkSheet.Range("B3").Resize(30, 4).Value = PSCIntercept
      Dim CellAddress As String
      For Fish As Integer = 0 To 29
         CellAddress = "A" & (Fish + 3).ToString
         If Fish = 0 Then
            xlWorkSheet.Range(CellAddress).Value = "Base"
         Else
            xlWorkSheet.Range(CellAddress).Value = (Fish + 1985).ToString
         End If
      Next

      Exit Sub

   End Sub


End Module
