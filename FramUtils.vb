Imports System.IO.File
Imports System.Data.OleDb

Module FramUtils

   Public Sub Resize_Form(ByVal Loadform As Form)
      Dim cControl As Control
      Dim i_i As Integer
      Dim Boldness As Boolean

      '- Don't ReSize if Screen Bounds are the same as Development Screen
      If (Screen.PrimaryScreen.Bounds.Height = 1024) And (Screen.PrimaryScreen.Bounds.Width = 1280) Then
         Exit Sub
      End If

      With Loadform
         For i_i = 0 To .Controls.Count - 1
            If TypeOf .Controls(i_i) Is ComboBox Then ' cannot change Height 
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width

            ElseIf TypeOf .Controls(i_i) Is Panel Then
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width
               .Controls(i_i).Height = .Controls(i_i).Height / FormHeight * Loadform.Height
               'Do the same for Panel's Children
               For Each cControl In .Controls(i_i).Controls

                  cControl.Left = cControl.Left / FormWidth * Loadform.Width
                  cControl.Top = cControl.Top / FormHeight * Loadform.Height
                  cControl.Width = cControl.Width / FormWidth * Loadform.Width
                  cControl.Height = cControl.Height / FormHeight * Loadform.Height
                  cControl.Font = New Font(cControl.Font.Name, cControl.Font.Size / FormWidth * Loadform.Width)

                  If TypeOf cControl Is PictureBox Then 'Make it stretch
                     cControl.BackgroundImageLayout = ImageLayout.Stretch
                  End If

               Next
            ElseIf TypeOf .Controls(i_i) Is System.Windows.Forms.TabControl Then
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width
               .Controls(i_i).Height = .Controls(i_i).Height / FormHeight * Loadform.Height

            ElseIf TypeOf .Controls(i_i) Is GroupBox Then
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width
               .Controls(i_i).Height = .Controls(i_i).Height / FormHeight * Loadform.Height

            ElseIf TypeOf .Controls(i_i) Is DataGridView Then
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               Jim = .Controls(i_i).Width
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width
               Jim = .Controls(i_i).Width
               .Controls(i_i).Height = .Controls(i_i).Height / FormHeight * Loadform.Height

            Else
               .Controls(i_i).Left = .Controls(i_i).Left / FormWidth * Loadform.Width
               .Controls(i_i).Top = .Controls(i_i).Top / FormHeight * Loadform.Height
               .Controls(i_i).Width = .Controls(i_i).Width / FormWidth * Loadform.Width
               .Controls(i_i).Height = .Controls(i_i).Height / FormHeight * Loadform.Height
            End If

            If .Controls(i_i).Font.Bold = True Then
               Boldness = True
            Else
               Boldness = False
            End If
            If Boldness = True Then
               .Controls(i_i).Font = New Font(.Controls(i_i).Font.Name, .Controls(i_i).Font.Size / FormWidth * Loadform.Width, FontStyle.Bold)
            Else
               .Controls(i_i).Font = New Font(.Controls(i_i).Font.Name, .Controls(i_i).Font.Size / FormWidth * Loadform.Width, FontStyle.Regular)
            End If
            'If FormHeight > Loadform.Height Then
            '   If Boldness = True Then
            '      .Controls(i_i).Font = New Font(.Controls(i_i).Font.Name, .Controls(i_i).Font.Size / FormHeight * Loadform.Height, FontStyle.Bold)
            '   Else
            '      .Controls(i_i).Font = New Font(.Controls(i_i).Font.Name, .Controls(i_i).Font.Size / FormHeight * Loadform.Height, FontStyle.Regular)
            '   End If
            'End If
         Next i_i
      End With

   End Sub


   Sub ReDimBaseArrays()
      '- ReDim Base Arrays
      ReDim BaseCohortSize(NumStk, MaxAge)
      ReDim BaseExploitationRate(NumStk, MaxAge, NumFish, NumSteps)
      ReDim AnyBaseRate(NumFish, NumSteps)
      ReDim BaseSubLegalRate(NumStk, MaxAge, NumFish, NumSteps)
      ReDim MaturationRate(NumStk, MaxAge, NumSteps)
      ReDim AEQ(NumStk, MaxAge, NumSteps)
      ReDim StockFishRateScalers(NumStk, NumFish, NumSteps)
      ReDim ModelStockProportion(NumFish)
      '- ModelStockProportion only applies to CHINOOK
      If SpeciesName = "COHO" Then
         For Fish = 1 To NumFish
            ModelStockProportion(Fish) = 1
         Next
      End If
      ReDim ShakerMortRate(NumFish, NumSteps)
      ReDim EncounterRateAdjustment(MaxAge, NumFish, NumSteps)
      '- All EncounterRateAdjustments are ONE unless specified differently
      For Age = 0 To MaxAge
         For Fish = 0 To NumFish
            For TStep = 0 To NumSteps
               EncounterRateAdjustment(Age, Fish, TStep) = 1
            Next
         Next
      Next
      ReDim NaturalMortality(MaxAge, NumSteps)
      ReDim IncidentalRate(NumFish, NumSteps)
      ReDim TerminalFisheryFlag(NumFish, NumSteps)
      ReDim VonBertL(NumStk, 1)
      ReDim VonBertT(NumStk, 1)
      ReDim VonBertK(NumStk, 1)
      ReDim VonBertCV(NumStk, MaxAge, 1)
      '- ReDim Stock Arrays
      ReDim StockID(NumStk)
      ReDim ProductionRegion(NumStk)
      ReDim ManagementUnit(NumStk)
      ReDim StockName(NumStk)
      ReDim StockTitle(NumStk)
      '- ReDim Fishery Arrays
      ReDim FisheryID(NumFish)
      ReDim FisheryName(NumFish)
      ReDim FisheryTitle(NumFish)
      ReDim ChinookBaseEncounterAdjustment(NumFish, NumSteps)
      ReDim ChinookBaseSizeLimit(NumFish, NumSteps)
      '- ReDim Time Step Arrays
      ReDim TimeStepID(NumSteps)
      ReDim TimeStepName(NumSteps)
      ReDim TimeStepTitle(NumSteps)
      ReDim MidTimeStep(NumSteps)
   End Sub

   Sub ReDimCalcArrays()
      '- ReDim Calculation Arrays
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
      ReDim Cohort(NumStk, MaxAge, 4, NumSteps)
      ReDim Escape(NumStk, MaxAge, NumSteps)
      ReDim TotalLandedCatch(NumFish * 2, NumSteps)
      ReDim TotalNonRetention(NumFish, NumSteps)
      ReDim TotalEncounters(NumFish, NumSteps)
      ReDim TotalShakers(NumFish, NumSteps)
      ReDim TotalDropOff(NumFish, NumSteps)
      ReDim CohoTime4Cohort(NumStk)
      '- ReDim Input Variables
      ReDim StockRecruit(NumStk, MaxAge, 2)
      ReDim FisheryScaler(NumFish, NumSteps)
      ReDim FisheryQuota(NumFish, NumSteps)
      ReDim MSFFisheryScaler(NumFish, NumSteps)
      ReDim MSFFisheryQuota(NumFish, NumSteps)
      ReDim FisheryFlag(NumFish, NumSteps)
      ReDim NonRetentionFlag(NumFish, NumSteps)
      ReDim NonRetentionInput(NumFish, NumSteps, 4)
      ReDim MarkSelectiveMortRate(NumFish, NumSteps)
      ReDim MarkSelectiveMarkMisID(NumFish, NumSteps)
      ReDim MarkSelectiveUnMarkMisID(NumFish, NumSteps)
      ReDim MarkSelectiveIncRate(NumFish, NumSteps)
      ReDim MinSizeLimit(NumFish, NumSteps)
      ReDim MaxSizeLimit(NumFish, NumSteps)
      ReDim PropLegCatch(NumStk, MaxAge)
      ReDim PropSubPop(NumStk, MaxAge)
      ReDim CNRShakers(NumStk, MaxAge)
      ReDim PSCMaxER(17)
      ReDim BackwardsTarget(NumStk)

      '=============================================
      'Pete 12/13 ReDim Code for External Sublegals vars
      ReDim TargetRatio(NumFish, MaxAge, NumSteps)
      ReDim RunEncounterRateAdjustment(NumFish, MaxAge, NumSteps)
      ReDim UpdWhen(NumFish, MaxAge, NumSteps)
      ReDim UpdBy(NumFish, MaxAge, NumSteps)
      ReDim Kfat(NumFish, MaxAge, NumSteps)


      If SpeciesName = "COHO" Then
         ReDim BackwardsFlag(NumStk)
      ElseIf SpeciesName = "CHINOOK" Then
            If NumStk = 38 Or NumStk > 75 Then
                ReDim BackwardsChinook(NumStk + NumStk / 2 + 20, MaxAge)
                ReDim BackwardsFlag(NumStk + NumStk + 20 / 2)
            Else
                ReDim BackwardsChinook(NumStk + 32, MaxAge)
                ReDim BackwardsFlag(NumStk + 32)
            End If
      End If
      ReDim StockFishRateScalers(NumStk, NumFish, NumSteps)
      '- Default value for Rate Scalers is ONE
      For Stk = 1 To NumStk
         For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
               StockFishRateScalers(Stk, Fish, TStep) = 1
            Next
         Next
      Next

      '- Set Edit Change Variables to False
      ChangeAnyInput = False
      ChangeBackFram = False
      ChangeFishScalers = False
      ChangeNonRetention = False
      ChangePSCMaxER = False
      ChangeSizeLimit = False
      ChangeStockFishScaler = False
      ChangeStockRecruit = False

   End Sub

   Sub ReadOldCommandFile()

      Dim TextLine As String
      '- First ReDim (Zero) All Input and Calculation Arrays
      '- ReDim Calculation Arrays
      Call ReDimCalcArrays()

      '- Text File Reader
      Dim CMDReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(OldCMDFile)
      CMDReader.TextFieldType = FileIO.FieldType.Delimited
      CMDReader.SetDelimiters(",")

      Dim CurrentRow As String()
      Dim CurrentField As String
      Dim CmdStr As String
      Dim LineNum, LoopLen, FieldNum, Result As Integer
      Dim RunName, Comments, CMDBasePeriodName, BaseName As String
      Dim NumComms, RecNum As Integer

      LineNum = 1
      Comments = ""
      RunName = ""

      '- Read CMD Header Information ...
      '   Species, RunName, Comments, BasePeriod, and Stock Recruits
      While Not CMDReader.EndOfData
         Try
            CurrentRow = CMDReader.ReadFields()
            LoopLen = CurrentRow.Length
            FieldNum = 0
            '- Fields are Line Number Specific
            Select Case LineNum
               Case 1
                  '- SpeciesName
                  SpeciesName = Trim(CurrentRow(0))
                  If SpeciesName = "COHO" Then
                     MsgBox("The Old COHO Command Files and Base Periods" & vbCrLf & "are not supported", MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
                  If SpeciesName = "NEWCOHO" Then
                     '- NEWCOHO was used in Previous FRAM Versions
                     SpeciesName = "COHO"
                  ElseIf SpeciesName = "CHINOOK" Then
                     Jim = 1
                  Else
                     MsgBox("Species Name not recognized for this Command File!!!", MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
               Case 2
                  RunName = My.Computer.FileSystem.GetFileInfo(OldCMDFile).Name
                  RunIDNameSelect = RunName
                  '- Use RunName for RunTitle Field when Reading Old CMD File
                  RunIDTitleSelect = CurrentRow(0)
                  If RunIDTitleSelect.Length > 100 Then
                     RunIDTitleSelect = RunIDTitleSelect.Substring(1, 99)
                  End If
               Case 3
                  NumComms = CInt(CurrentRow(0))
               Case 4 To (NumComms + 3)
                  For Each CurrentField In CurrentRow
                     If InStr(CurrentField, "Calibration", CompareMethod.Text) > 0 Then
                        '- Test for blank Line in Comments Section of Old CMD files
                        LineNum += 1
                        GoTo CmdBlankLine
                     End If
                     Comments &= CurrentField
                  Next
                  RunIDCommentsSelect = Comments
               Case (NumComms + 4)

                  '- BASE PERIOD for this CMD File

CmdBlankLine:
                  CMDBasePeriodName = CurrentRow(0)
                  '- Get Filename without the extension for comparison
                  BaseName = My.Computer.FileSystem.GetFileInfo(CMDBasePeriodName).Name
                  LoopLen = InStr(BaseName, ".OUT", CompareMethod.Text)
                  If LoopLen <> 0 Then
                     BaseName = Mid(BaseName, 1, LoopLen - 1)
                  End If
                  '- Values from the Associated Base Period File determine
                  '- the Format and Number of Lines for the remainder of this
                  '- Command text file.
                  '- First - Determine if an Existing Base Period Recordset exists
                  '- Second- Ask if Existing or New Base Period should be used
                  CmdStr = "SELECT * FROM BaseID WHERE SpeciesName = " & Chr(34) & SpeciesName & Chr(34) & " ORDER BY BasePeriodID"
                  Dim BPcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                  Dim BaseDA As New System.Data.OleDb.OleDbDataAdapter
                  BaseDA.SelectCommand = BPcm
                  Dim BPcb As New OleDb.OleDbCommandBuilder
                  BPcb = New OleDb.OleDbCommandBuilder(BaseDA)
                  If FramDataSet.Tables.Contains("BasePeriodIDList") Then
                     FramDataSet.Tables("BasePeriodIDList").Clear()
                  End If
                  BaseDA.Fill(FramDataSet, "BasePeriodIDList")
                  Dim NumBP As Integer
                  NumBP = FramDataSet.Tables("BasePeriodIDList").Rows.Count
                  '- Loop through Table Records for Base Period Match
                  For RecNum = 0 To NumBP - 1
                     BasePeriodID = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(1)
                     BasePeriodName = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(2)
                     If InStr(BasePeriodName, BaseName, CompareMethod.Text) <> 0 Then
                        Result = MsgBox("Found Matching Base Period Name in Database" & vbCrLf & _
                        "DatabaseName=" & BasePeriodName & vbCrLf & _
                        "NumStks=" & FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(4) & vbCrLf & _
                        "NumFish=" & FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(5) & vbCrLf & _
                        "NumSteps=" & FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(6) & vbCrLf & _
                        "Do You want to use this BasePeriod Version ???", MsgBoxStyle.YesNo)
                        If Result = vbNo Then
                           MsgBox("Please Read Selected Base Period File before Reading CMD File", MsgBoxStyle.OkOnly)
                           BaseDA = Nothing
                           Exit Sub
                        End If
                        '- Found Correct Existing Base Period File
                        '- Eventually change this to Call BasePeriodFillArrays
                        'SpeciesName = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(2)
                        NumStk = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(4)
                        NumFish = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(5)
                        NumSteps = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(6)
                        NumAge = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(7)
                        MinAge = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(8)
                        MaxAge = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(9)
                        BasePeriodDate = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(10)
                        BasePeriodComments = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(11)
                        StockVersion = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(12)
                        FisheryVersion = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(13)
                        TimeStepVersion = FramDataSet.Tables("BasePeriodIDList").Rows(RecNum)(14)
                        Call ReDimBaseArrays()
                        Call ReDimCalcArrays()
                        Exit For
                     End If
                  Next
                  If RecNum = NumBP Then
                     MsgBox("Can't Find the BASE PERIOD FILE for this CMD File in Database" & _
                     "Add BASE PERIOD FILE or Change CMD File OUT Name to Existing Name", MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
                  BaseDA = Nothing

               Case (NumComms + 5) To (NumComms + 7)
                  '- Header Lines .. Not Needed
                  Jim = 1

               Case NumComms + 8 To (((MaxAge - 1) * NumStk) + (NumComms + 8))
                  '- Recruit Scalers
                  RecNum = LineNum - (NumComms + 7)
                  If SpeciesName = "COHO" Then
                     '- COHO has two lines - Ages 2&3
                     If (RecNum Mod 2) = 0 Then
                        Stk = RecNum / 2
                        Age = 3
                        StockRecruit(Stk, Age, 1) = (CurrentRow(0))
                     End If
                  ElseIf SpeciesName = "CHINOOK" Then
                     '- CHINOOK has four lines - Ages 2 to 5
                     If (RecNum Mod 4) = 1 Then
                        Stk = (RecNum + 3) / 4
                        Age = 2
                     ElseIf (RecNum Mod 4) = 2 Then
                        Stk = (RecNum + 2) / 4
                        Age = 3
                     ElseIf (RecNum Mod 4) = 3 Then
                        Stk = (RecNum + 1) / 4
                        Age = 4
                     ElseIf (RecNum Mod 4) = 0 Then
                        Stk = RecNum / 4
                        Age = 5
                     End If
                     StockRecruit(Stk, Age, 1) = (CurrentRow(0))
                  End If

            End Select

         Catch ex As Exception
            MsgBox("The COMMAND FILE file Selected has Format Problems" & vbCrLf & " RecNum=" & RecNum.ToString, MsgBoxStyle.OkOnly)
            Exit Sub
         End Try
         LineNum += 1
         '- Exit DO WHILE when Header Information has been Read
         If LineNum = (((MaxAge - 1) * NumStk) + (NumComms + 8)) Then Exit While
      End While

      '- Read CMD Time-Step Specific Information ...
      '   Species, RunName, Comments, BasePeriod, and Stock Recruits

      Dim Flag, NumCNR, NumSFS, NumFld As Integer
      '- Note: Many CMD lines DO NOT have comma delimiters !!! Must use Space DeLimiter
      CMDReader.SetDelimiters(",", " ")
      For TStep = 1 To NumSteps
         '- Size Limit Title Line
         CurrentRow = CMDReader.ReadFields()
NewCohoOldFormat:
         For Fish = 1 To NumFish
            '- One Size Limit Line per Fishery
            CurrentRow = CMDReader.ReadFields()
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  MinSizeLimit(Fish, TStep) = CurrentField
                  Exit For
               End If
            Next
         Next
         '- Quota/Fishery Scaler Title Line
         CurrentRow = CMDReader.ReadFields()
         For Fish = 1 To NumFish
            '- Two Lines per Fishery
            CurrentRow = CMDReader.ReadFields()
            '- First Field in First Line has FisheryFlag
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  FisheryFlag(Fish, TStep) = CurrentField
                  Exit For
               End If
            Next
            '- Second Line has Variable Parameters
            CurrentRow = CMDReader.ReadFields()
            If FisheryFlag(Fish, TStep) = 0 Then
               '- Fishery Scaler
               For Each CurrentField In CurrentRow
                  If CurrentField <> "" Then
                     FisheryScaler(Fish, TStep) = CurrentField
                     Exit For
                  End If
               Next
               FisheryFlag(Fish, TStep) = 1
            ElseIf FisheryFlag(Fish, TStep) = 1 Then
               '- Fishery Quota
               For Each CurrentField In CurrentRow
                  If CurrentField <> "" Then
                     FisheryQuota(Fish, TStep) = CurrentField
                     Exit For
                  End If
               Next
               FisheryFlag(Fish, TStep) = 2
            ElseIf FisheryFlag(Fish, TStep) = 9 Then
               '- Mark-Selective Fishery Parameters .. 7 Fields
               NumFld = 1
               For Each CurrentField In CurrentRow
                  If CurrentField <> "" Then
                     Select Case NumFld
                        Case 1
                           Flag = CurrentField
                        Case 2
                           If Flag = 0 Then
                              MSFFisheryScaler(Fish, TStep) = CurrentField
                              FisheryFlag(Fish, TStep) = 7
                           ElseIf Flag = 1 Then
                              MSFFisheryQuota(Fish, TStep) = CurrentField
                              FisheryFlag(Fish, TStep) = 8
                           End If
                        Case 3
                           Jim = 1
                        Case 4
                           MarkSelectiveMortRate(Fish, TStep) = CurrentField
                        Case 5
                           MarkSelectiveMarkMisID(Fish, TStep) = CurrentField
                        Case 6
                           MarkSelectiveUnMarkMisID(Fish, TStep) = CurrentField
                        Case 7
                           MarkSelectiveIncRate(Fish, TStep) = CurrentField
                     End Select
                     NumFld += 1
                  End If
               Next
            End If
         Next
         '- CNR Title Line
         CurrentRow = CMDReader.ReadFields()
         '- Number of CNR Fisheries for this Time Step
         CurrentRow = CMDReader.ReadFields()
         For Each CurrentField In CurrentRow
            If CurrentField <> "" Then
               NumCNR = CurrentField
               Exit For
            End If
         Next
         For Flag = 1 To NumCNR
            '- First Field in First Line has Fishery Number
            CurrentRow = CMDReader.ReadFields()
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  Fish = CurrentField
                  Exit For
               End If
            Next
            '- Second Line has CNR Flag
            CurrentRow = CMDReader.ReadFields()
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  NonRetentionFlag(Fish, TStep) = CurrentField
                  '- No Longer using Zero Based Flagging
                  If SpeciesName = "COHO" Then
                     NonRetentionFlag(Fish, TStep) = 1
                  ElseIf SpeciesName = "CHINOOK" Then
                     NonRetentionFlag(Fish, TStep) = NonRetentionFlag(Fish, TStep) + 1
                  End If
                  Exit For
               End If
            Next
            '- Next Four Lines have CNR Parameters
            For LineNum = 1 To 4
               CurrentRow = CMDReader.ReadFields()
               For Each CurrentField In CurrentRow
                  If CurrentField <> "" Then
                     NonRetentionInput(Fish, TStep, LineNum) = CurrentField
                     Exit For
                  End If
               Next
            Next
         Next
         '- Stock/Fishery Scaler Title Line
         CurrentRow = CMDReader.ReadFields()
         '- Check for Old Format for COHO (i.e. No SHRS Values)
         If TStep = NumSteps And IsNothing(CurrentRow) Then
            Call CopyNewRecordset()
            Exit Sub
         End If
         TextLine = ""
         For Stk = 0 To CurrentRow.Length - 1
            TextLine &= CurrentRow(Stk).ToString
         Next
         If InStr(TextLine, "StockSpecificExploitation") >= 1 Then
            '- Check for Old Format for COHO (i.e. No SHRS Values)
         Else
            TStep += 1
            GoTo NewCohoOldFormat
         End If
         '- Number of Stock/Fishery Scalers for this Time Step
         CurrentRow = CMDReader.ReadFields()
         For Each CurrentField In CurrentRow
            If CurrentField <> "" Then
               NumSFS = CurrentField
               Exit For
            End If
         Next
         For Flag = 1 To NumSFS
            '- One Line per SF-Scaler .. expect single Space at start of line
            CurrentRow = CMDReader.ReadFields()
            FieldNum = 1
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  Select Case FieldNum
                     Case 1
                        Stk = CurrentField
                     Case 2
                        Fish = CurrentField
                     Case 3
                        StockFishRateScalers(Stk, Fish, TStep) = CurrentField
                        Exit For
                  End Select
                  FieldNum += 1
               End If
            Next
         Next
      Next TStep

      '- PSC COHO Max ER Values for Coho Tech Report
      If SpeciesName = "COHO" Then
         If CMDReader.EndOfData Then
            '- If No PSC Max ER's (Old CMD Format) Set All to 0.5
            For Stk = 1 To 17
               PSCMaxER(Stk) = 0.5
            Next
            Call CopyNewRecordset()
            Exit Sub
         End If
         '- PSC Title Line
         CurrentRow = CMDReader.ReadFields()
         For Stk = 1 To 17
            CurrentRow = CMDReader.ReadFields()
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  PSCMaxER(Stk) = CurrentField
                  Exit For
               End If
            Next
            If CMDReader.EndOfData Then
               Call CopyNewRecordset()
               Exit Sub
            End If
         Next
      End If
      If CMDReader.EndOfData Then
         Call CopyNewRecordset()
         Exit Sub
      End If

      '- Backwards FRAM Parameters
      '- Blank Line & BF Title Line
      'CurrentRow = CMDReader.ReadFields()
      CurrentRow = CMDReader.ReadFields()
      For Each CurrentField In CurrentRow
         If CurrentField = "Backwards" Then
            Jim = 1
            Exit For
         End If
      Next
      If SpeciesName = "COHO" Then
         For Stk = 1 To NumStk
            CurrentRow = CMDReader.ReadFields()
            FieldNum = 1
            For Each CurrentField In CurrentRow
               If CurrentField <> "" Then
                  If FieldNum = 1 Then
                     LineNum = CurrentField
                  ElseIf FieldNum = 2 Then
                     BackwardsTarget(Stk) = CurrentField
                  ElseIf FieldNum = 3 Then
                     BackwardsFlag(Stk) = CurrentField
                     Exit For
                  End If
                  FieldNum += 1
               End If
            Next
         Next
        ElseIf SpeciesName = "CHINOOK" Then

            If NumStk = 38 Or NumStk = 76 Then
                NumChinTermRuns = 37
            ElseIf NumStk = 33 Or NumStk = 66 Then
                NumChinTermRuns = 32
            Else
                NumChinTermRuns = NumStk / 2 - 1
            End If

            For Stk = 1 To NumStk + NumChinTermRuns
                CurrentRow = CMDReader.ReadFields()
                FieldNum = 1
                For Each CurrentField In CurrentRow
                    If CurrentField <> "" Then
                        If FieldNum = 1 Then
                            LineNum = CurrentField
                        ElseIf FieldNum = 2 Then
                            BackwardsChinook(Stk, 3) = CurrentField
                        ElseIf FieldNum = 3 Then
                            BackwardsChinook(Stk, 4) = CurrentField
                        ElseIf FieldNum = 4 Then
                            BackwardsChinook(Stk, 5) = CurrentField
                        ElseIf FieldNum = 5 Then
                            BackwardsFlag(Stk) = CurrentField
                            Exit For
                        End If
                        FieldNum += 1
                    End If
                Next
            Next
      End If

      Call CopyNewRecordset()

   End Sub

   Sub CopyNewRecordset()

      Dim RunIDRead, BaseIDRead As Integer

      '- Save Current RunID and BaseID
      RunIDRead = RunIDSelect
      BaseIDRead = BasePeriodIDSelect

      '- Fill Database Tables with New RecordSet Values from Arrays
      '- This Routine is called from ReadOldCMDFile and CopyRecordset
      '- Both Calls are from FVS_FramUtils

      '- NewRunID Variable set in CopyRecordset Selection from FVS_EditRecordsetInfo
      If RecordsetSelectionType = 4 Or RecordsetSelectionType = 5 Then GoTo SkipRID
      If RunIDSelect = 0 Then
         '- Empty RunID Recordset
         NewRunID = 1
         GoTo SkipRID
      End If

      '- RunID DataBase Table New Record --------
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      Dim MaxOldID As Integer
      '- Get Current Max RunID Value, Add One for New Recordset RunID Value
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM RunID ORDER BY RunID DESC"
      FramDB.Open()
      drd1 = cmd1.ExecuteReader
      drd1.Read()
      MaxOldID = drd1.GetInt32(1)
      cmd1.Dispose()
      drd1.Dispose()
      FramDB.Close()

      NewRunID = MaxOldID + 1

      Dim FramTrans As OleDb.OleDbTransaction
      Dim RIC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
        RIC.Connection = FramDB

      RIC.Transaction = FramTrans
        RIC.CommandText = "INSERT INTO RunID (RunID,SpeciesName,RunName,RunTitle,BasePeriodID,RunComments,CreationDate,ModifyInputDate,RunTimeDate) " & _
            "VALUES(" & RunIDSelect.ToString & "," & _
                Chr(34) & SpeciesName.ToString & Chr(34) & "," & _
              Chr(34) & RunIDNameSelect.ToString & Chr(34) & "," & _
              Chr(34) & RunIDTitleSelect.ToString & Chr(34) & "," & _
              BasePeriodID.ToString & "," & _
              Chr(34) & RunIDCommentsSelect.ToString & Chr(34) & "," & _
              Chr(35) & RunIDCreationDateSelect.ToString & Chr(35) & "," & _
              Chr(35) & Now().ToString & Chr(35) & "," & _
              Chr(35) & RunIDRunTimeDateSelect.ToString & Chr(35) & "," & _
              Chr(34) & RunIDYearSelect & Chr(34) & ");"



        '"VALUES(" & NewRunID.ToString & "," & _
        'Chr(34) & SpeciesName.ToString & Chr(34) & "," & _
        'Chr(34) & RunIDNameSelect.ToString & Chr(34) & "," & _
        'Chr(34) & RunIDTitleSelect.ToString & Chr(34) & "," & _
        'BasePeriodID.ToString & "," & _
        'Chr(34) & RunIDCommentsSelect.ToString & Chr(34) & "," & _
        'Chr(35) & Now().ToString & Chr(35) & "," & _
        'Chr(35) & RunIDModifyInputDateSelect.ToString & Chr(35) & "," & _
        'Chr(35) & RunIDRunTimeDateSelect.ToString & Chr(35) & ")"
      RIC.ExecuteNonQuery()
      FramTrans.Commit()
      FramDB.Close()

SkipRID:
      '- RunID Information Already Saved in ModelRun Copy
      If RecordsetSelectionType = 4 Or RecordsetSelectionType = 5 Then
         NewRunID = RunIDSelect
      End If

      '- STOCKRECRUIT DataBase Table Save --------
      Dim SRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      SRC.Connection = FramDB
      SRC.Transaction = FramTrans
      For Stk = 1 To NumStk
         For Age = MinAge To MaxAge
            If StockRecruit(Stk, Age, 1) <> 0 Then
               StockRecruit(Stk, Age, 2) = StockRecruit(Stk, Age, 1) * BaseCohortSize(Stk, Age)
               SRC.CommandText = "INSERT INTO StockRecruit (RunID,StockID,Age,RecruitScaleFactor,RecruitCohortSize) " & _
               "VALUES(" & NewRunID.ToString & "," & _
               Stk.ToString & "," & _
               Age.ToString & "," & _
               StockRecruit(Stk, Age, 1).ToString("######0.0000") & "," & _
               StockRecruit(Stk, Age, 2).ToString("######0.0000") & ")"
               SRC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- FISHERYSCALERS DataBase Table Save --------
      Dim FSC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      FSC.Connection = FramDB
      FSC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If FisheryFlag(Fish, TStep) <> 0 Or FisheryScaler(Fish, TStep) <> 0 Or FisheryQuota(Fish, TStep) <> 0 Or MSFFisheryScaler(Fish, TStep) <> 0 Or MSFFisheryQuota(Fish, TStep) <> 0 Then
               '- MarkSelectiveFlag currently not used ... placeholder after Quota
                    FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                    "VALUES(" & NewRunID.ToString & "," & _
                    Fish.ToString & "," & _
                    TStep.ToString & "," & _
                    FisheryFlag(Fish, TStep).ToString & "," & _
                    FisheryScaler(Fish, TStep).ToString("######0.0000") & "," & _
                    FisheryQuota(Fish, TStep).ToString("########0.0000") & "," & _
                    MSFFisheryScaler(Fish, TStep).ToString("######0.0000") & "," & _
                    MSFFisheryQuota(Fish, TStep).ToString("########0.0000") & "," & _
                    MarkSelectiveMortRate(Fish, TStep).ToString("######0.0000") & "," & _
                    MarkSelectiveMarkMisID(Fish, TStep).ToString("######0.0000") & "," & _
                    MarkSelectiveUnMarkMisID(Fish, TStep).ToString("######0.0000") & "," & _
                    MarkSelectiveIncRate(Fish, TStep).ToString("######0.0000") & ")"
               FSC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- NONRETENTION DataBase Table Save --------
      Dim NRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      NRC.Connection = FramDB
      NRC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If NonRetentionFlag(Fish, TStep) <> 0 Then
               NRC.CommandText = "INSERT INTO NonRetention (RunID,FisheryID,TimeStep,NonRetentionFlag,CNRInput1,CNRInput2,CNRInput3,CNRInput4) " & _
               "VALUES(" & NewRunID.ToString & "," & _
               Fish.ToString & "," & _
               TStep.ToString & "," & _
               NonRetentionFlag(Fish, TStep).ToString & "," & _
               NonRetentionInput(Fish, TStep, 1).ToString("######0.0000") & "," & _
               NonRetentionInput(Fish, TStep, 2).ToString("######0.0000") & "," & _
               NonRetentionInput(Fish, TStep, 3).ToString("######0.0000") & "," & _
               NonRetentionInput(Fish, TStep, 4).ToString("######0.0000") & ")"
               NRC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- SIZELIMIT DataBase Table Save --------
      If SpeciesName = "COHO" Then GoTo SkipCohoSL
      Dim SLC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      SLC.Connection = FramDB
      SLC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If MinSizeLimit(Fish, TStep) <> 0 Then
               SLC.CommandText = "INSERT INTO SizeLimits (RunID,FisheryID,TimeStep,MinimumSize) " & _
               "VALUES(" & NewRunID.ToString & "," & _
               Fish.ToString & "," & _
               TStep.ToString & "," & _
               MinSizeLimit(Fish, TStep).ToString("######0") & ")"
               SLC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()
SkipCohoSL:

      '- STOCKFISHERYRATESCALER DataBase Table Save --------
      Dim SFRS As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      SFRS.Connection = FramDB
      SFRS.Transaction = FramTrans
      For Stk = 1 To NumStk
         For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
               If StockFishRateScalers(Stk, Fish, TStep) <> 1 Then
                  SFRS.CommandText = "INSERT INTO StockFisheryRateScaler (RunID,StockID,FisheryID,TimeStep,StockFisheryRateScaler) " & _
                  "VALUES(" & NewRunID.ToString & "," & _
                  Stk.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  StockFishRateScalers(Stk, Fish, TStep).ToString("######0.0000") & ")"
                  SFRS.ExecuteNonQuery()
               End If
            Next
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- PSC ER Maximum DataBase Table Save --------
      If SpeciesName = "CHINOOK" Then GoTo SkipChinookMER
      Dim MEC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      MEC.Connection = FramDB
      MEC.Transaction = FramTrans
      For Stk = 1 To 17
         If PSCMaxER(Stk) = 0 Then PSCMaxER(Stk) = 0.5
         MEC.CommandText = "INSERT INTO PSCMaxER (RunID,PSCStockID,PSCMaxER) " & _
         "VALUES(" & NewRunID.ToString & "," & _
         Stk.ToString & "," & _
         PSCMaxER(Stk).ToString("0.0000") & ")"
         MEC.ExecuteNonQuery()
      Next
      FramTrans.Commit()
      FramDB.Close()
SkipChinookMER:

      '- Backwards FRAM Target Escapements and Flag
      Dim BFC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BFC.Connection = FramDB
      BFC.Transaction = FramTrans
      If SpeciesName = "COHO" Then
         For Stk = 1 To NumStk
            If BackwardsTarget(Stk) <> 0 Then
               BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
               "VALUES(" & NewRunID.ToString & "," & _
               Stk.ToString & "," & _
               BackwardsTarget(Stk).ToString("0.0") & ", 0, 0, " & _
               BackwardsFlag(Stk).ToString & ")"
               BFC.ExecuteNonQuery()
            End If
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         Dim SumChinTarget As Double
            'SumChinTarget = BackwardsChinook(Stk, 3) + BackwardsChinook(Stk, 4) + BackwardsChinook(Stk, 5)
            If NumStk = 38 Or NumStk = 76 Then
                NumChinTermRuns = 37
            ElseIf NumStk = 33 Or NumStk = 66 Then
                NumChinTermRuns = 32
            Else
                NumChinTermRuns = NumStk / 2 - 1
            End If
         For Stk = 1 To NumStk + NumChinTermRuns
                'If SumChinTarget <> 0 Then
                BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
                "VALUES(" & NewRunID.ToString & "," & _
                Stk.ToString & "," & _
                BackwardsChinook(Stk, 3).ToString("0.0") & "," & _
                BackwardsChinook(Stk, 4).ToString("0.0") & "," & _
                BackwardsChinook(Stk, 5).ToString("0.0") & "," & _
                BackwardsFlag(Stk).ToString & ")"
                BFC.ExecuteNonQuery()
                'End If
            Next
        End If
        FramTrans.Commit()
        FramDB.Close()

        '- Set Edit Change Variables to False
        ChangeAnyInput = False
        ChangeBackFram = False
        ChangeFishScalers = False
        ChangeNonRetention = False
        ChangePSCMaxER = False
        ChangeSizeLimit = False
        ChangeStockFishScaler = False
        ChangeStockRecruit = False

        '- Save Mortality, Cohort, and Escapement Arrays to Database (2/11/2011)

        '- MORTALITY DataBase Table Save --------

        Dim FIC As New OleDbCommand
        Dim RCount, TimeStep As Integer
        Dim MortSum As Double
        FramDB.Open()
        FramTrans = FramDB.BeginTransaction
        FIC.Connection = FramDB
        FIC.Transaction = FramTrans
        RCount = 0
        For Stk = 1 To NumStk
            For Age = 1 To MaxAge
                For Fish = 1 To NumFish
                    For TimeStep = 1 To NumSteps
                        MortSum = LandedCatch(Stk, Age, Fish, TimeStep) + NonRetention(Stk, Age, Fish, TimeStep) + Shakers(Stk, Age, Fish, TimeStep) + DropOff(Stk, Age, Fish, TimeStep) + MSFLandedCatch(Stk, Age, Fish, TimeStep) + MSFNonRetention(Stk, Age, Fish, TimeStep) + MSFShakers(Stk, Age, Fish, TimeStep) + MSFDropOff(Stk, Age, Fish, TimeStep)
                        If MortSum <> 0 Then
                            RCount += 1
                            FIC.CommandText = "INSERT INTO Mortality (PrimaryKey,RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                            "VALUES(" & RCount.ToString & "," & _
                            NewRunID.ToString & "," & _
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
                        End If
                    Next
                Next
            Next
        Next
        FramTrans.Commit()
        FramDB.Close()

        '- COHORT DataBase Table Save --------

        Dim CohTrans As OleDb.OleDbTransaction
        Dim FCC As New OleDbCommand
        FramDB.Open()
        CohTrans = FramDB.BeginTransaction
        FCC.Connection = FramDB
        FCC.Transaction = CohTrans
        For Stk = 1 To NumStk
            For Age = 1 To MaxAge
                For TimeStep = 1 To NumSteps
                    If Cohort(Stk, Age, 3, TimeStep) <> 0 Or Cohort(Stk, Age, 1, TimeStep) <> 0 Then
                        FCC.CommandText = "INSERT INTO Cohort (RunID,StockID,Age,TimeStep,Cohort,MatureCohort,StartCohort,WorkingCohort,MidCohort) " & _
                        "VALUES(" & NewRunID.ToString & "," & _
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

        '- ESCAPEMENT DataBase Table Save --------

        Dim ESCTrans As OleDb.OleDbTransaction
        Dim FEC As New OleDbCommand
        FramDB.Open()
        ESCTrans = FramDB.BeginTransaction
        FEC.Connection = FramDB
        FEC.Transaction = ESCTrans
        For Stk = 1 To NumStk
            For Age = 1 To MaxAge
                For TimeStep = 1 To NumSteps
                    If Escape(Stk, Age, TimeStep) <> 0 Then
                        FEC.CommandText = "INSERT INTO Escapement (RunID,StockID,Age,TimeStep,Escapement) " & _
                        "VALUES(" & NewRunID.ToString & "," & _
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

        '- Save Total FisheryMortality Table 

        Dim TFMTrans As OleDb.OleDbTransaction
        Dim TFM As New OleDbCommand
        FramDB.Open()
        TFMTrans = FramDB.BeginTransaction
        TFM.Connection = FramDB
        TFM.Transaction = TFMTrans
        Dim TotFM As Double
        For Fish = 1 To NumFish
            For TimeStep = 1 To NumSteps
                TotFM = TotalLandedCatch(Fish, TimeStep) + TotalLandedCatch(NumFish + Fish, TimeStep) + TotalEncounters(Fish, TimeStep) + TotalShakers(Fish, TimeStep) + TotalDropOff(Fish, TimeStep) + TotalNonRetention(Fish, TimeStep)
                If TotFM <> 0 Then
                    TFM.CommandText = "INSERT INTO FisheryMortality (RunID,FisheryID,TimeStep,TotalLandedCatch,TotalUnMarkedCatch,TotalNonRetention,TotalShakers,TotalDropOff,TotalEncounters) " & _
                    "VALUES(" & NewRunID.ToString & "," & _
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
        'Pete 12/13 -- Copy recordset code for SLRatio and RunEncounterRateAdjustment Tables
        '- SLRatio DataBase Table Save --------
        If ReadOldCmd = False Then
            If SpeciesName = "CHINOOK" Then
                Dim SLRatC As New OleDbCommand
                FramDB.Open()
                FramTrans = FramDB.BeginTransaction
                SLRatC.Connection = FramDB
                SLRatC.Transaction = FramTrans
                For Fish = 1 To NumFish
                    For Age = MinAge To MaxAge
                        For TimeStep = 1 To NumSteps

                            If UpdBy(Fish, Age, TimeStep) <> "not updated--ignore datetime" Then
                                'Uncomment the If end if content if cluttering with 1.00 is undesired; for now, testing, leave it in...
                                'If TargetRatio(Fish, Age, TimeStep) <> 1 Or TargetRatio(Fish, Age, TimeStep) <> 1 Then 
                                SLRatC.CommandText = "INSERT INTO SLRatio (RunID,FisheryID,Age,TimeStep,TargetRatio,RunEncounterRateAdjustment,UpdateWhen,UpdateBy) " & _
                                "VALUES(" & NewRunID.ToString & "," & _
                                Fish.ToString & "," & _
                                Age.ToString & "," & _
                                TimeStep.ToString & "," & _
                                TargetRatio(Fish, Age, TimeStep).ToString & "," & _
                                RunEncounterRateAdjustment(Fish, Age, TimeStep).ToString & "," & _
                                "'" & UpdWhen(Fish, Age, TimeStep).ToString & "'" & "," & _
                                "'" & UpdBy(Fish, Age, TimeStep).ToString & "'" & ")"
                                SLRatC.ExecuteNonQuery()
                                'End If
                            End If
                        Next
                    Next
                Next
                FramTrans.Commit()
                FramDB.Close()
            End If
        End If
        '===================================================================================



        '- Retrieve Current RunID and BaseID
        FVS_ModelRunSelection.GetRunVariables(BaseIDRead, RunIDRead)


   End Sub

   Sub DeleteRecordset()


      '============================================================================
      'Pete 12/13 Code for multi-run deletion
      'See also *** if/then statements for bypassing dialog boxes below...
      If multiRunDeleteMode = True Then
         RunIDDelete = multiRunPass
      End If
      '============================================================================

      '- Read RUN Selection Variables
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      Dim DeleteSpeciesName As String
      Dim result As Integer
      Dim RunIDNameDelete, RunIDTitleDelete, RunIDCommentsDelete As String
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM RunID WHERE RunID = " & CStr(RunIDDelete)
      FramDB.Open()
      drd1 = cmd1.ExecuteReader
      drd1.Read()
      'RunIDDelete = drd1.GetInt32(1)          '- Current Run ID (User Selection)
      DeleteSpeciesName = drd1.GetString(2)
      RunIDNameDelete = drd1.GetString(3)     '- Delete Run Name
      RunIDTitleDelete = drd1.GetString(4)    '- Delete Run Title
      RunIDCommentsDelete = drd1.GetString(6) '- Delete Run Comments
      cmd1.Dispose()
      drd1.Dispose()
      FramDB.Close()

      If multiRunDeleteMode = False Then '*** Pete 12/13 If/then allows for bypass during multi-delete mode
         result = MsgBox("Is this the Correct RunID/Recordset to DELETE???" & vbCrLf & _
             "RunID  = " & RunIDDelete.ToString & vbCrLf & _
             "Species= " & DeleteSpeciesName & vbCrLf & _
             "Name   = " & RunIDNameDelete & vbCrLf & _
             "Title  = " & RunIDTitleDelete, MsgBoxStyle.YesNo)
         If result = vbNo Then Exit Sub
      End If '*** Pete 12/13 If/then allows for bypass during multi-delete mode


      '- Delete All Records for this RunID

      '- RunID SELECT Statement
      Dim CmdStr As String
      CmdStr = "SELECT * FROM RunID WHERE RunID = " & RunIDDelete.ToString
      Dim RIcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim RunDA As New System.Data.OleDb.OleDbDataAdapter
      RunDA.SelectCommand = RIcm
      '- RunID DELETE Statement
      CmdStr = "DELETE * FROM RunID WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim RIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      RunDA.DeleteCommand = RIDcm
      '- Command Builder
      Dim RIcb As New OleDb.OleDbCommandBuilder
      RIcb = New OleDb.OleDbCommandBuilder(RunDA)
      FramDB.Open()
      RunDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- StockRecruit SELECT Statement
      CmdStr = "SELECT * FROM StockRecruit WHERE RunID = " & RunIDDelete.ToString
      Dim SRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SRDA As New System.Data.OleDb.OleDbDataAdapter
      SRDA.SelectCommand = SRcm
      '- StockRecruit DELETE Statement
      CmdStr = "DELETE * FROM StockRecruit WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim SRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SRDA.DeleteCommand = SRDcm
      '- Command Builder
      Dim SRcb As New OleDb.OleDbCommandBuilder
      SRcb = New OleDb.OleDbCommandBuilder(SRDA)
      FramDB.Open()
      SRDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- FisheryScalers SELECT Statement
      CmdStr = "SELECT * FROM FisheryScalers WHERE RunID = " & RunIDDelete.ToString
      Dim FScm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim FSDA As New System.Data.OleDb.OleDbDataAdapter
      FSDA.SelectCommand = FScm
      '- FisheryScalers DELETE Statement
      CmdStr = "DELETE * FROM FisheryScalers WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim FSDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      FSDA.DeleteCommand = FSDcm
      '- Command Builder
      Dim FScb As New OleDb.OleDbCommandBuilder
      FScb = New OleDb.OleDbCommandBuilder(FSDA)
      FramDB.Open()
      FSDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- NonRetention SELECT Statement
      CmdStr = "SELECT * FROM NonRetention WHERE RunID = " & RunIDDelete.ToString
      Dim NRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim NRDA As New System.Data.OleDb.OleDbDataAdapter
      NRDA.SelectCommand = NRcm
      '- NonRetention DELETE Statement
      CmdStr = "DELETE * FROM NonRetention WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim NRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      NRDA.DeleteCommand = NRDcm
      '- Command Builder
      Dim NRcb As New OleDb.OleDbCommandBuilder
      NRcb = New OleDb.OleDbCommandBuilder(NRDA)
      FramDB.Open()
      NRDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- SizeLimits SELECT Statement
      If DeleteSpeciesName = "COHO" Then GoTo SkipCohoDEL
      CmdStr = "SELECT * FROM SizeLimits WHERE RunID = " & RunIDDelete.ToString
      Dim SLcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SLDA As New System.Data.OleDb.OleDbDataAdapter
      SLDA.SelectCommand = SLcm
      '- SizeLimits DELETE Statement
      CmdStr = "DELETE * FROM SizeLimits WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim SLDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SLDA.DeleteCommand = SLDcm
      '- Command Builder
      Dim SLcb As New OleDb.OleDbCommandBuilder
      SLcb = New OleDb.OleDbCommandBuilder(SLDA)
      FramDB.Open()
      SLDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()
SkipCohoDEL:

      '- PSC Max ER SELECT Statement
      If DeleteSpeciesName = "CHINOOK" Then GoTo SkipChinookME
      CmdStr = "SELECT * FROM PSCMaxER WHERE RunID = " & RunIDDelete.ToString
      Dim MEcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim MEDA As New System.Data.OleDb.OleDbDataAdapter
      MEDA.SelectCommand = MEcm
      '- PSC Max ER DELETE Statement
      CmdStr = "DELETE * FROM PSCMaxER WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim MEDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      MEDA.DeleteCommand = MEDcm
      '- Command Builder
      Dim MEcb As New OleDb.OleDbCommandBuilder
      MEcb = New OleDb.OleDbCommandBuilder(MEDA)
      FramDB.Open()
      MEDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()
SkipChinookME:

      '- Mortality SELECT Statement
      CmdStr = "SELECT * FROM Mortality WHERE RunID = " & RunIDDelete.ToString
      Dim Mcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim MDA As New System.Data.OleDb.OleDbDataAdapter
      MDA.SelectCommand = Mcm
      '- Mortality DELETE Statement
      CmdStr = "DELETE * FROM Mortality WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim MDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      MDA.DeleteCommand = MDcm
      '- Command Builder
      Dim Mcb As New OleDb.OleDbCommandBuilder
      Mcb = New OleDb.OleDbCommandBuilder(MDA)
      FramDB.Open()
      MDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- Escapement SELECT Statement
      CmdStr = "SELECT * FROM Escapement WHERE RunID = " & RunIDDelete.ToString
      Dim Ecm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim EDA As New System.Data.OleDb.OleDbDataAdapter
      EDA.SelectCommand = Ecm
      '- Escapement DELETE Statement
      CmdStr = "DELETE * FROM Escapement WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim EDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      EDA.DeleteCommand = EDcm
      '- Command Builder
      Dim Ecb As New OleDb.OleDbCommandBuilder
      Ecb = New OleDb.OleDbCommandBuilder(EDA)
      FramDB.Open()
      EDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- Cohort SELECT Statement
      CmdStr = "SELECT * FROM Cohort WHERE RunID = " & RunIDDelete.ToString
      Dim Ccm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim CDA As New System.Data.OleDb.OleDbDataAdapter
      CDA.SelectCommand = Ccm
      '- Cohort DELETE Statement
      CmdStr = "DELETE * FROM Cohort WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim CDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      CDA.DeleteCommand = CDcm
      '- Command Builder
      Dim Ccb As New OleDb.OleDbCommandBuilder
      Ccb = New OleDb.OleDbCommandBuilder(CDA)
      FramDB.Open()
      CDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- BackwardsFRAM SELECT Statement
      CmdStr = "SELECT * FROM BackwardsFRAM WHERE RunID = " & RunIDDelete.ToString
      Dim BFcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim BFDA As New System.Data.OleDb.OleDbDataAdapter
      BFDA.SelectCommand = BFcm
      '- BackwardsFRAM DELETE Statement
      CmdStr = "DELETE * FROM BackwardsFRAM WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim BFDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      BFDA.DeleteCommand = BFDcm
      '- Command Builder
      Dim BFcb As New OleDb.OleDbCommandBuilder
      BFcb = New OleDb.OleDbCommandBuilder(BFDA)
      FramDB.Open()
      BFDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- FisheryMortality SELECT Statement
      CmdStr = "SELECT * FROM FisheryMortality WHERE RunID = " & RunIDDelete.ToString
      Dim FMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim FMDA As New System.Data.OleDb.OleDbDataAdapter
      FMDA.SelectCommand = FMcm
      '- FisheryMortality DELETE Statement
      CmdStr = "DELETE * FROM FisheryMortality WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim FMDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      FMDA.DeleteCommand = FMDcm
      '- Command Builder
      Dim FMcb As New OleDb.OleDbCommandBuilder
      FMcb = New OleDb.OleDbCommandBuilder(FMDA)
      FramDB.Open()
      FMDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- StockFisheryRateScaler SELECT Statement
      CmdStr = "SELECT * FROM StockFisheryRateScaler WHERE RunID = " & RunIDDelete.ToString
      Dim SRScm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SRSDA As New System.Data.OleDb.OleDbDataAdapter
      SRSDA.SelectCommand = SRScm
      '- StockFisheryRateScaler DELETE Statement
      CmdStr = "DELETE * FROM StockFisheryRateScaler WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim SRSDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SRSDA.DeleteCommand = SRSDcm
      '- Command Builder
      Dim SRScb As New OleDb.OleDbCommandBuilder
      SRScb = New OleDb.OleDbCommandBuilder(SRSDA)
      FramDB.Open()
      SRSDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '===========================================================================
      'Pete 12/13 Delete Records from SLRatio and RunEncounterRateAdjustment tables

      '- SLRatio SELECT Statement
      CmdStr = "SELECT * FROM SLRatio WHERE RunID = " & RunIDDelete.ToString
      Dim SLRatcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SLRatDA As New System.Data.OleDb.OleDbDataAdapter
      SLRatDA.SelectCommand = SLRatcm
      '- SLRatio DELETE Statement
      CmdStr = "DELETE * FROM SLRatio WHERE RunID = " & RunIDDelete.ToString & ";"
      Dim SLRatDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SLRatDA.DeleteCommand = SLRatDcm
      '- Command Builder
      Dim SLRatcb As New OleDb.OleDbCommandBuilder
      SLRatcb = New OleDb.OleDbCommandBuilder(SLRatDA)
      FramDB.Open()
      SLRatDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '===========================================================================


      RunIDDelete = 0

   End Sub

   Sub ReadOldBasePeriodOUTFile()

      '- Text File Reader
      Dim BaseReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(OldOUTFile)
      BaseReader.TextFieldType = FileIO.FieldType.Delimited
      BaseReader.SetDelimiters(",")

      Dim CurrentRow As String()
      Dim CurrentField, Scratch As String
      Dim LineNum, FieldNum, Result, NumRates As Integer
      Dim RecNum As Integer

      '- Read OUT Header Information ...
      '   NumStk, NumFish, NumSteps, MaxAge
      LineNum = 1
      While Not BaseReader.EndOfData
         Try
            CurrentRow = BaseReader.ReadFields()
            '- Fields are Line Number Specific
            Select Case LineNum
               Case 1
                  '- Number of Stocks
                  NumStk = CurrentRow(0)
               Case 2
                  '- Number of Fisheries
                  NumFish = CurrentRow(0)
               Case 3
                  '- Number of Time Steps
                  NumSteps = CurrentRow(0)
               Case 4
                  '- Maximum Age
                  MaxAge = CurrentRow(0)
               Case 5
                  '- Maximum Encounter Rate Age .. no longer used
                  Scratch = CurrentRow(0)
            End Select
         Catch ex As Exception
                MsgBox("The BASE-PERIOD FILE Selected has Format Problems" & vbCrLf & "in the HEADER INFORMATION section", MsgBoxStyle.OkOnly)
            Exit Sub
         End Try
         LineNum += 1
         '- Exit DO WHILE when Header Information has been Read
         If LineNum > 5 Then Exit While
      End While

      '- Query User about Species ... Not Included in Base-Period File Format
      If NumStk > 100 And NumFish > 100 Then
         Result = MsgBox("The Number of Stocks and Fisheries indicate that this" & vbCrLf & _
          "BASE PERIOD File is for COHO ... Is that Correct ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            SpeciesName = "COHO"
         Else
            SpeciesName = "CHINOOK"
         End If
      Else
         Result = MsgBox("The Number of Stocks and Fisheries indicate that this" & vbCrLf & _
          "BASE PERIOD File is for CHINOOK ... Is that Correct ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            SpeciesName = "CHINOOK"
         Else
            SpeciesName = "COHO"
         End If
      End If

      '- ReDim Base Arrays
      Call ReDimBaseArrays()

      '- Set NumAges and MinAge by Species plus COHO Default TermFlag and Maturation
      If SpeciesName = "COHO" Then
         NumAge = 1
         MinAge = 3
         '- Set Default TerminalFisheryFlag for COHO
         TStep = NumSteps
         For Fish = 1 To NumFish
            TerminalFisheryFlag(Fish, TStep) = 1
         Next
         '- Set Default Maturation Rates for COHO
         TStep = NumSteps
         Age = MaxAge
         For Stk = 1 To NumStk
            MaturationRate(Stk, Age, TStep) = 1
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         NumAge = 4
         MinAge = 2
      End If

      '- Read CHINOOK AEQ, Growth (VonBertLanf), and Shaker Flags
      If SpeciesName = "CHINOOK" Then
         '- AEQ Title Line
         CurrentRow = BaseReader.ReadFields()
         For Stk = 1 To NumStk
            For Age = MaxAge To MinAge Step -1
               For TStep = NumSteps To 1 Step -1
                  Try
                     CurrentRow = BaseReader.ReadFields()
                     AEQ(Stk, Age, TStep) = CDbl(CurrentRow(0))
                  Catch ex As Exception
                     MsgBox("The BASE_PERIOD FILE file Selected has Format Problems" & vbCrLf & "in the AEQ section for Stock=" & Stk.ToString, MsgBoxStyle.OkOnly)
                     Exit Sub
                  End Try
               Next
            Next
         Next
         '- Growth Parameter Title Line
         CurrentRow = BaseReader.ReadFields()
         Dim MatType As Integer
         For Stk = 1 To NumStk
            For MatType = 0 To 1
               Try
                  CurrentRow = BaseReader.ReadFields()
                  VonBertL(Stk, MatType) = CurrentRow(0)
                  CurrentRow = BaseReader.ReadFields()
                  VonBertT(Stk, MatType) = CurrentRow(0)
                  CurrentRow = BaseReader.ReadFields()
                  VonBertK(Stk, MatType) = CurrentRow(0)
                  For Age = MinAge To MaxAge
                     CurrentRow = BaseReader.ReadFields()
                     VonBertCV(Stk, Age, MatType) = CurrentRow(0)
                  Next
               Catch ex As Exception
                  MsgBox("The BASE_PERIOD FILE file Selected has Format Problems" & vbCrLf & "in the GROWTH section for Stock=" & Stk.ToString, MsgBoxStyle.OkOnly)
                  Exit Sub
               End Try
            Next
         Next
         '- CHINOOK TimeStep Midpoints (for PNV Calculations)
         For TStep = 1 To NumSteps
            CurrentRow = BaseReader.ReadFields()
            MidTimeStep(TStep) = CurrentRow(0)
         Next
         '- Shaker Inclusion Flag Title Line
         CurrentRow = BaseReader.ReadFields()
         'MidTimeStep(TStep) = CurrentRow(0)
         '- These Flags are no longer used but still exist in the BasePeriod File
         For Fish = 1 To NumFish
            CurrentRow = BaseReader.ReadFields()
         Next
      End If

      '---- Read Base Period COHORT SIZES
      LineNum = 1
      While Not BaseReader.EndOfData
         Try
            CurrentRow = BaseReader.ReadFields()
            Select Case LineNum
               Case 1
                  '- Cohort Title line
                  Scratch = CurrentRow(0)
               Case 2 To (((MaxAge - 1) * NumStk) + 1)
                  '- Cohort Sizes
                  RecNum = LineNum - 1
                  If SpeciesName = "COHO" Then
                     '- COHO has two lines - Ages 2&3
                     If (RecNum Mod 2) = 0 Then
                        Stk = RecNum / 2
                        Age = 3
                        BaseCohortSize(Stk, Age) = (CurrentRow(0))
                     End If
                  ElseIf SpeciesName = "CHINOOK" Then
                     '- CHINOOK has four lines - Ages 2 to 5
                     If (RecNum Mod 4) = 1 Then
                        Stk = (RecNum + 3) / 4
                        Age = 2
                     ElseIf (RecNum Mod 4) = 2 Then
                        Stk = (RecNum + 2) / 4
                        Age = 3
                     ElseIf (RecNum Mod 4) = 3 Then
                        Stk = (RecNum + 1) / 4
                        Age = 4
                     ElseIf (RecNum Mod 4) = 0 Then
                        Stk = RecNum / 4
                        Age = 5
                     End If
                     BaseCohortSize(Stk, Age) = (CurrentRow(0))
                  End If
            End Select
         Catch ex As Exception
                MsgBox("The BASE_PERIOD FILE Selected has Format Problems" & vbCrLf & "in the COHORT SIZES section", MsgBoxStyle.OkOnly)
            Exit Sub
         End Try
         LineNum += 1
         '- Exit DO WHILE when Stock-Recruit Information has been Read
         If LineNum >= (((MaxAge - 1) * NumStk) + 2) Then Exit While
      End While

      '- Read CHINOOK Fishery ModelStockProportions
      If SpeciesName = "CHINOOK" Then
         Try
            '- Title Line
            CurrentRow = BaseReader.ReadFields()
            For Fish = 1 To NumFish
               CurrentRow = BaseReader.ReadFields()
               ModelStockProportion(Fish) = CurrentRow(0)
            Next
         Catch ex As Exception
            MsgBox("The BASE_PERIOD FILE file Selected has Format Problems" & vbCrLf & "in the MODEL STOCK PROP. section", MsgBoxStyle.OkOnly)
            Exit Sub
         End Try
      End If

      '- Read Other Mortality (DropOff and DropOut Rates)
      LineNum = 1
      '- Title Line
      CurrentRow = BaseReader.ReadFields()
      While Not BaseReader.EndOfData
         Try
            CurrentRow = BaseReader.ReadFields()
            '- LineNum and Fish are the same here
            IncidentalRate(LineNum, 1) = CurrentRow(0)
            '- New Base Period has Incidental Rate for all Time-Steps
            For TStep = 2 To NumSteps
               IncidentalRate(LineNum, TStep) = IncidentalRate(LineNum, 1)
            Next
         Catch ex As Exception
            MsgBox("The BASE-PERIOD FILE file Selected has Format Problems" & vbCrLf & "in the INCIDENTAL RATE section", MsgBoxStyle.OkOnly)
            Exit Sub
         End Try
         LineNum += 1
         '- Exit DO WHILE when Incidental Information has been Read
         If LineNum > NumFish Then Exit While
      End While

      '- Read TIME-STEP Information ... 

      If SpeciesName = "COHO" Then
         '- Note: OUT Expl. Rate lines DO NOT have comma delimiters 
         BaseReader.SetDelimiters(",", " ")
         For TStep = 1 To NumSteps
            LineNum = 1
            While Not BaseReader.EndOfData
               Try
                  CurrentRow = BaseReader.ReadFields()
                  Select Case LineNum
                     Case 1
                        '- Natural-Mortality Title line
                        Scratch = CurrentRow(0)
                     Case 2
                        '- Natural-Mortality COHO Age 2 .. Not Used
                        Scratch = CurrentRow(0)
                     Case 3
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              NaturalMortality(3, TStep) = CurrentField
                              Exit For
                           End If
                        Next
                     Case 4
                        '- Exploitation-Rate Title line
                        Scratch = CurrentRow(0)
                     Case 5
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              NumRates = CurrentField
                              Exit For
                           End If
                        Next
                     Case 6
                        '- Exploitation-Rate Header line
                        Scratch = CurrentRow(0)
                     Case 7 To (NumRates + 7)
                        FieldNum = 1
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              Select Case FieldNum
                                 Case 1
                                    Stk = CurrentField
                                    FieldNum += 1
                                 Case 2
                                    Age = CurrentField
                                    FieldNum += 1
                                 Case 3
                                    Fish = CurrentField
                                    FieldNum += 1
                                 Case 4
                                    BaseExploitationRate(Stk, Age, Fish, TStep) = CurrentField
                                    Exit For
                              End Select
                           End If
                        Next
                  End Select
               Catch ex As Exception
                        MsgBox("The BASE-PERIOD FILE Selected has Format Problems" & vbCrLf & "in the EXPL. RATE section - TimeStep=" & TStep.ToString, MsgBoxStyle.OkOnly)
                  Exit Sub
               End Try
               LineNum += 1
               If LineNum = NumRates + 7 Then
                  Jim = 1
               End If
               '- Exit DO WHILE when Header Information has been Read
               If LineNum = NumRates + 7 Then Exit While
            End While
         Next

      ElseIf SpeciesName = "CHINOOK" Then

         '- Note: OUT Expl. Rate lines DO NOT have comma delimiters 
         BaseReader.SetDelimiters(",", " ")
         Dim NumMatRates As Integer
         For TStep = 1 To NumSteps
            LineNum = 1
            While Not BaseReader.EndOfData
               Try
                  CurrentRow = BaseReader.ReadFields()
                  Select Case LineNum
                     Case 1
                        '- Natural-Mortality Title line
                        Scratch = CurrentRow(0)
                     Case 2 To 5
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              NaturalMortality(LineNum, TStep) = CurrentField
                              Exit For
                           End If
                        Next
                     Case 6
                        '- Shaker-Rate Title line
                        Scratch = CurrentRow(0)
                     Case 7 To (NumFish + 6)
                        '- SubLegal Hooking Mortality Rates (ShakerMortRate)
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              ShakerMortRate(LineNum - 6, TStep) = CurrentField
                              Exit For
                           End If
                        Next
                     Case (NumFish + 7)
                        '- SubLegal Shaker Encounter Rate Adjustment Title line
                        Scratch = CurrentRow(0)
                     Case (NumFish + 8) To (NumFish * 2 + 7)
                        '- SubLegal Shaker Encounter Rate Adjustment
                        FieldNum = 1
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              EncounterRateAdjustment(FieldNum + 1, LineNum - (NumFish + 7), TStep) = CurrentField
                              FieldNum += 1
                              If FieldNum = 4 Then
                                 EncounterRateAdjustment(5, LineNum - (NumFish + 7), TStep) = 1
                                 Exit For
                              End If
                           End If
                        Next
                     Case (NumFish * 2 + 8)
                        '- Terminal Fishery Flag Title line
                        Scratch = CurrentRow(0)
                     Case (NumFish * 2 + 9) To (NumFish * 3 + 8)
                        '- Terminal Fishery Flags 
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              TerminalFisheryFlag(LineNum - (NumFish * 2 + 8), TStep) = CurrentField
                              Exit For
                           End If
                        Next
                     Case (NumFish * 3 + 9)
                        '- Maturity Rate Title Line
                        Scratch = CurrentRow(0)
                     Case (NumFish * 3 + 10)
                        '- Maturity Rate NumRates Line
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              NumMatRates = CurrentField
                              Exit For
                           End If
                        Next
                     Case (NumFish * 3 + 11)
                        '- Maturity Rate Headers Line
                        Scratch = CurrentRow(0)
                     Case (NumFish * 3 + 12) To (NumFish * 3 + 11 + NumMatRates)
                        '- Maturity Rates
                        FieldNum = 1
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              Select Case FieldNum
                                 Case 1
                                    Stk = CurrentField
                                 Case 2
                                    Age = CurrentField
                                 Case 3
                                    MaturationRate(Stk, Age, TStep) = CurrentField
                                    Exit For
                              End Select
                              FieldNum += 1
                           End If
                        Next

                     Case (NumFish * 3 + 12 + NumMatRates)
                        '- Exploitation Rate Title Line
                        Scratch = CurrentRow(0)
                     Case (NumFish * 3 + 13 + NumMatRates)
                        '- Exploitation Rate NumRates Line
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              NumRates = CurrentField
                              Exit For
                           End If
                        Next
                     Case (NumFish * 3 + 14 + NumMatRates)
                        '- Exploitation Rate Headers Line
                        Scratch = CurrentRow(0)
                     Case (NumFish * 3 + 15 + NumMatRates) To (NumFish * 3 + 14 + NumMatRates + NumRates)
                        '- Exploitation Rates
                        FieldNum = 1
                        For Each CurrentField In CurrentRow
                           If CurrentField <> "" Then
                              Select Case FieldNum
                                 Case 1
                                    Stk = CurrentField
                                 Case 2
                                    Age = CurrentField
                                 Case 3
                                    Fish = CurrentField
                                 Case 4
                                    BaseExploitationRate(Stk, Age, Fish, TStep) = CurrentField
                                 Case 5
                                    BaseSubLegalRate(Stk, Age, Fish, TStep) = CurrentField
                                    Exit For
                              End Select
                              FieldNum += 1
                           End If
                        Next
                  End Select
               Catch ex As Exception
                        MsgBox("The BASE-PERIOD FILE Selected has Format Problems" & vbCrLf & "in the EXPL. RATE section - TimeStep=" & TStep.ToString, MsgBoxStyle.OkOnly)
                  Exit Sub
               End Try
               LineNum += 1
               '- Exit DO WHILE when TimeStep Information has been Read
               If LineNum = (NumFish * 3 + 15 + NumMatRates + NumRates) Then Exit While
            End While
         Next

      End If

      '- Fill Database Tables with New RecordSet Values from Arrays

      '- BaseID DataBase Table New Record --------
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      Dim MaxOldID, NewBaseID, LoopLen, ChinookStockVersion As Integer
      Dim BaseName As String

      '- Get Current Max RunID Value, Add One for New Recordset RunID Value
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM BaseID ORDER BY BasePeriodID DESC"
      FramDB.Open()
      drd1 = cmd1.ExecuteReader
      drd1.Read()
      MaxOldID = drd1.GetInt32(1)
      cmd1.Dispose()
      drd1.Dispose()
      FramDB.Close()

      NewBaseID = MaxOldID + 1

      '- Number of CHINOOK Stocks varies by Base Period Type and Purpose
      If SpeciesName$ = "CHINOOK" Then
            Select Case NumStk
                Case 78
                    ChinookStockVersion = 5
                Case 76
                    ChinookStockVersion = 1
                Case 38
                    ChinookStockVersion = 2
                Case 66
                    ChinookStockVersion = 3
                Case 33
                    ChinookStockVersion = 4
            End Select
      ElseIf SpeciesName = "COHO" Then
         ChinookStockVersion = 1
      End If

      '- BaseID Database Table
      BaseName = My.Computer.FileSystem.GetFileInfo(OldOUTFile).Name
      LoopLen = InStr(BaseName.ToUpper, ".OUT")
      If LoopLen <> 0 Then
         BaseName = Mid(BaseName, 1, LoopLen - 1)
      End If
      Dim FramTrans As OleDb.OleDbTransaction
      Dim BIC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BIC.Connection = FramDB
      BIC.Transaction = FramTrans
      BIC.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
         "VALUES(" & NewBaseID.ToString & "," & _
         Chr(34) & BaseName.ToString & Chr(34) & "," & _
         Chr(34) & SpeciesName.ToString & Chr(34) & "," & _
         NumStk.ToString & "," & _
         NumFish.ToString & "," & _
         NumSteps.ToString & "," & _
         NumAge.ToString & "," & _
         MinAge.ToString & "," & _
         MaxAge.ToString & "," & _
         Chr(35) & Now().ToString & Chr(35) & "," & _
         Chr(34) & "From File = " & OldOUTFile & Chr(34) & "," & _
         ChinookStockVersion.ToString & ",1,1)"
      '- StockVersion, FisheryVersion, and TimeStepVersion will need to be updated
      '- when different Base-Period files are created ... Default values for now
      '- except for Chinook Stock Version
      BIC.ExecuteNonQuery()
      FramTrans.Commit()
      FramDB.Close()

      '- BaseCohort Size Database Table 
      Dim BCIC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BCIC.Connection = FramDB
      BCIC.Transaction = FramTrans
      For Stk = 1 To NumStk
         For Age = MinAge To MaxAge
            If BaseCohortSize(Stk, Age) <> 0 Then
               BCIC.CommandText = "INSERT INTO BaseCohort (BasePeriodID,StockID,Age,BaseCohortSize) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Stk.ToString & "," & _
                  Age.ToString & "," & _
                  BaseCohortSize(Stk, Age).ToString & ")"
               BCIC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- BaseExploitationRate Database Table 
      Dim BERC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BERC.Connection = FramDB
      BERC.Transaction = FramTrans
      For Stk = 1 To NumStk
         For Age = MinAge To MaxAge
            For Fish = 1 To NumFish
               For TStep = 1 To NumSteps
                  If BaseExploitationRate(Stk, Age, Fish, TStep) <> 0 Or BaseSubLegalRate(Stk, Age, Fish, TStep) <> 0 Then
                     BERC.CommandText = "INSERT INTO BaseExploitationRate (BasePeriodID,StockID,Age,FisheryID,TimeStep,ExploitationRate,SubLegalEncounterRate) " & _
                        "VALUES(" & NewBaseID.ToString & "," & _
                        Stk.ToString & "," & _
                        Age.ToString & "," & _
                        Fish.ToString & "," & _
                        TStep.ToString & "," & _
                        BaseExploitationRate(Stk, Age, Fish, TStep).ToString("0.0000000000") & "," & _
                        BaseSubLegalRate(Stk, Age, Fish, TStep).ToString("0.0000000000") & ")"
                     BERC.ExecuteNonQuery()
                  End If
               Next
            Next
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- IncidentalRate Database Table 
      Dim BIRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BIRC.Connection = FramDB
      BIRC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If IncidentalRate(Fish, TStep) <> 0 Then
               BIRC.CommandText = "INSERT INTO IncidentalRate (BasePeriodID,FisheryID,TimeStep,IncidentalRate) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  IncidentalRate(Fish, TStep).ToString("0.0000") & ")"
               BIRC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- ShakerMortRate Database Table 
      Dim SMRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      SMRC.Connection = FramDB
      SMRC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If ShakerMortRate(Fish, TStep) <> 0 Then
               SMRC.CommandText = "INSERT INTO ShakerMortRate (BasePeriodID,FisheryID,TimeStep,ShakerMortRate) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  ShakerMortRate(Fish, TStep).ToString("0.0000") & ")"
               SMRC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- NaturalMortality Rate Database Table 
      Dim NMRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      NMRC.Connection = FramDB
      NMRC.Transaction = FramTrans
      For TStep = 1 To NumSteps
         For Age = MinAge To MaxAge
            If NaturalMortality(Age, TStep) <> 0 Then
               NMRC.CommandText = "INSERT INTO NaturalMortality (BasePeriodID,Age,TimeStep,NaturalMortalityRate) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Age.ToString & "," & _
                  TStep.ToString & "," & _
                  NaturalMortality(Age, TStep).ToString("0.000000") & ")"
               NMRC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- TerminalFisheryFlag Database Table 
      Dim TFFC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      TFFC.Connection = FramDB
      TFFC.Transaction = FramTrans
      For Fish = 1 To NumFish
         For TStep = 1 To NumSteps
            If TerminalFisheryFlag(Fish, TStep) <> 0 Then
               TFFC.CommandText = "INSERT INTO TerminalFisheryFlag (BasePeriodID,FisheryID,TimeStep,TerminalFlag) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  TerminalFisheryFlag(Fish, TStep).ToString & ")"
               TFFC.ExecuteNonQuery()
            End If
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- Maturation Database Table 
      Dim BMRC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      BMRC.Connection = FramDB
      BMRC.Transaction = FramTrans
      For Stk = 1 To NumStk
         For Age = MinAge To MaxAge
            For TStep = 1 To NumSteps
               If MaturationRate(Stk, Age, TStep) <> 0 Then
                  BMRC.CommandText = "INSERT INTO MaturationRate (BasePeriodID,StockID,Age,TimeStep,MaturationRate) " & _
                     "VALUES(" & NewBaseID.ToString & "," & _
                     Stk.ToString & "," & _
                     Age.ToString & "," & _
                     TStep.ToString & "," & _
                     MaturationRate(Stk, Age, TStep).ToString("0.00000000") & ")"
                  BMRC.ExecuteNonQuery()
               End If
            Next
         Next
      Next
      FramTrans.Commit()
      FramDB.Close()

      '- ModelStockProportion Database Table ... CHINOOK Only
      If SpeciesName = "CHINOOK" Then
         Dim MSPC As New OleDbCommand
         FramDB.Open()
         FramTrans = FramDB.BeginTransaction
         MSPC.Connection = FramDB
         MSPC.Transaction = FramTrans
         For Fish = 1 To NumFish
            If ModelStockProportion(Fish) <> 0 Then
               MSPC.CommandText = "INSERT INTO FisheryModelStockProportion (BasePeriodID,FisheryID,ModelStockProportion) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Fish.ToString & "," & _
                  ModelStockProportion(Fish).ToString("0.00000000") & ")"
               MSPC.ExecuteNonQuery()
            End If
         Next
         FramTrans.Commit()
         FramDB.Close()
      End If

      '- EncounterRateAdjustment Database Table ... CHINOOK Only
      If SpeciesName = "CHINOOK" Then
         Dim EAPC As New OleDbCommand
         FramDB.Open()
         FramTrans = FramDB.BeginTransaction
         EAPC.Connection = FramDB
         EAPC.Transaction = FramTrans
         For Age = MinAge To MaxAge
            For Fish = 1 To NumFish
               For TStep = 1 To NumSteps
                  'If EncounterRateAdjustment(Age, Fish, TStep) <> 1 Then
                  EAPC.CommandText = "INSERT INTO EncounterRateAdjustment (BasePeriodID,Age,FisheryID,TimeStep,EncounterRateAdjustment) " & _
                     "VALUES(" & NewBaseID.ToString & "," & _
                     Age.ToString & "," & _
                     Fish.ToString & "," & _
                     TStep.ToString & "," & _
                     EncounterRateAdjustment(Age, Fish, TStep).ToString("###0.0000") & ")"
                  EAPC.ExecuteNonQuery()
                  'End If
               Next
            Next
         Next
         FramTrans.Commit()
         FramDB.Close()
      End If

      '- Growth Database Table ... CHINOOK Only
      If SpeciesName = "CHINOOK" Then
         Dim GPC As New OleDbCommand
         FramDB.Open()
         FramTrans = FramDB.BeginTransaction
         GPC.Connection = FramDB
         GPC.Transaction = FramTrans
         For Stk = 1 To NumStk
            If VonBertL(Stk, 0) <> 0 Then
               GPC.CommandText = "INSERT INTO Growth (BasePeriodID,StockID,LImmature,KImmature,TImmature,CV2Immature,CV3Immature,CV4Immature,CV5Immature,LMature,KMature,TMature,CV2Mature,CV3Mature,CV4Mature,CV5Mature) " & _
                  "VALUES(" & NewBaseID.ToString & "," & _
                  Stk.ToString & "," & _
                  VonBertL(Stk, 0).ToString("###0.000") & "," & _
                  VonBertK(Stk, 0).ToString("###0.000") & "," & _
                  VonBertT(Stk, 0).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 2, 0).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 3, 0).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 4, 0).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 5, 0).ToString("###0.000") & "," & _
                  VonBertL(Stk, 1).ToString("###0.000") & "," & _
                  VonBertK(Stk, 1).ToString("###0.000") & "," & _
                  VonBertT(Stk, 1).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 2, 1).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 3, 1).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 4, 1).ToString("###0.000") & "," & _
                  VonBertCV(Stk, 5, 1).ToString("###0.000") & ")"
               GPC.ExecuteNonQuery()
            End If
         Next
         FramTrans.Commit()
         FramDB.Close()
      End If

      If SpeciesName = "CHINOOK" Then
         '- AEQ Database Table 
         Dim ARC As New OleDbCommand
         FramDB.Open()
         FramTrans = FramDB.BeginTransaction
         ARC.Connection = FramDB
         ARC.Transaction = FramTrans
         For Stk = 1 To NumStk
            For Age = MinAge To MaxAge
               For TStep = 1 To NumSteps
                  If AEQ(Stk, Age, TStep) <> 0 Then
                     ARC.CommandText = "INSERT INTO AEQ (BasePeriodID,StockID,Age,TimeStep,AEQ) " & _
                        "VALUES(" & NewBaseID.ToString & "," & _
                        Stk.ToString & "," & _
                        Age.ToString & "," & _
                        TStep.ToString & "," & _
                        AEQ(Stk, Age, TStep).ToString("0.00000000") & ")"
                     ARC.ExecuteNonQuery()
                  End If
               Next
            Next
         Next
         FramTrans.Commit()
         FramDB.Close()
      End If

   End Sub

   Sub DeleteBasePeriodRecordset()

      Dim DeleteSpeciesName As String
      Dim AnyRunIDRecs, Result As Integer
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()

      '- Look for any RunID Records using the Selected BASE PERIOD Delete
      FramDB.Open()
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM RunID WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      drd1 = cmd1.ExecuteReader
      AnyRunIDRecs = 0
      Do While drd1.Read
         AnyRunIDRecs = 1
         Exit Do
      Loop
      If AnyRunIDRecs = 1 Then
         Result = MsgBox("There are Model Runs using this Base Period Recordset" & vbCrLf & "Do you want to Continue???", MsgBoxStyle.YesNo)
         If Result = vbNo Then
            FramDB.Close()
            Exit Sub
         End If
      End If
      drd1.Close()
      'BasePeriodName = drd1.GetString(2)
      'DeleteSpeciesName = drd1.GetString(3)
      'FramDB.Close()

      '- Get Delete BASE PERIOD Delete Selection SpeciesName
      'FramDB.Open()
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM BaseID WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      drd1 = cmd1.ExecuteReader
      drd1.Read()
      BasePeriodName = drd1.GetString(2)
      DeleteSpeciesName = drd1.GetString(3)
      FramDB.Close()

      '- BaseID SELECT Statement
      Dim CmdStr As String
      CmdStr = "SELECT * FROM BaseID WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim BIcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim BaseDA As New System.Data.OleDb.OleDbDataAdapter
      BaseDA.SelectCommand = BIcm
      '- BaseID DELETE Statement
      CmdStr = "DELETE * FROM BaseID WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim BIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      BaseDA.DeleteCommand = BIDcm
      '- Command Builder
      Dim BIcb As New OleDb.OleDbCommandBuilder
      BIcb = New OleDb.OleDbCommandBuilder(BaseDA)
      FramDB.Open()
      BaseDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- BaseCohort SELECT Statement
      CmdStr = "SELECT * FROM BaseCohort WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim BCcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim BCDA As New System.Data.OleDb.OleDbDataAdapter
      BCDA.SelectCommand = BCcm
      '- BaseCohort DELETE Statement
      CmdStr = "DELETE * FROM BaseCohort WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim BCDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      BCDA.DeleteCommand = BCDcm
      '- Command Builder
      Dim BCcb As New OleDb.OleDbCommandBuilder
      BCcb = New OleDb.OleDbCommandBuilder(BCDA)
      FramDB.Open()
      BCDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- BaseExploitationRate SELECT Statement
      CmdStr = "SELECT * FROM BaseExploitationRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim BEcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim BEDA As New System.Data.OleDb.OleDbDataAdapter
      BEDA.SelectCommand = BEcm
      '- BaseExploitationRate DELETE Statement
      CmdStr = "DELETE * FROM BaseExploitationRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim BEDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      BEDA.DeleteCommand = BEDcm
      '- Command Builder
      Dim BEcb As New OleDb.OleDbCommandBuilder
      BEcb = New OleDb.OleDbCommandBuilder(BEDA)
      FramDB.Open()
      BEDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- IncidentalRate SELECT Statement
      CmdStr = "SELECT * FROM IncidentalRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim IRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim IRDA As New System.Data.OleDb.OleDbDataAdapter
      IRDA.SelectCommand = IRcm
      '- IncidentalRate DELETE Statement
      CmdStr = "DELETE * FROM IncidentalRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim IRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      IRDA.DeleteCommand = IRDcm
      '- Command Builder
      Dim IRcb As New OleDb.OleDbCommandBuilder
      IRcb = New OleDb.OleDbCommandBuilder(IRDA)
      FramDB.Open()
      IRDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- ShakerMortRate SELECT Statement
      CmdStr = "SELECT * FROM ShakerMortRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim SMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SMDA As New System.Data.OleDb.OleDbDataAdapter
      SMDA.SelectCommand = SMcm
      '- ShakerMortRate DELETE Statement
      CmdStr = "DELETE * FROM ShakerMortRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim SMDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SMDA.DeleteCommand = SMDcm
      '- Command Builder
      Dim SMcb As New OleDb.OleDbCommandBuilder
      SMcb = New OleDb.OleDbCommandBuilder(SMDA)
      FramDB.Open()
      SMDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- MaturationRate SELECT Statement
      CmdStr = "SELECT * FROM MaturationRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim MRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim MRDA As New System.Data.OleDb.OleDbDataAdapter
      MRDA.SelectCommand = MRcm
      '- MaturationRate DELETE Statement
      CmdStr = "DELETE * FROM MaturationRate WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim MRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      MRDA.DeleteCommand = MRDcm
      '- Command Builder
      Dim MRcb As New OleDb.OleDbCommandBuilder
      MRcb = New OleDb.OleDbCommandBuilder(MRDA)
      FramDB.Open()
      MRDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- NaturalMortality SELECT Statement
      CmdStr = "SELECT * FROM NaturalMortality WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim NRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim NRDA As New System.Data.OleDb.OleDbDataAdapter
      NRDA.SelectCommand = NRcm
      '- NaturalMortality DELETE Statement
      CmdStr = "DELETE * FROM NaturalMortality WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim NRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      NRDA.DeleteCommand = NRDcm
      '- Command Builder
      Dim NRcb As New OleDb.OleDbCommandBuilder
      NRcb = New OleDb.OleDbCommandBuilder(NRDA)
      FramDB.Open()
      NRDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- TerminalFisheryFlag SELECT Statement
      CmdStr = "SELECT * FROM TerminalFisheryFlag WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim TFcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim TFDA As New System.Data.OleDb.OleDbDataAdapter
      TFDA.SelectCommand = TFcm
      '- TerminalFisheryFlag DELETE Statement
      CmdStr = "DELETE * FROM TerminalFisheryFlag WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim TFDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      TFDA.DeleteCommand = TFDcm
      '- Command Builder
      Dim TFcb As New OleDb.OleDbCommandBuilder
      TFcb = New OleDb.OleDbCommandBuilder(TFDA)
      FramDB.Open()
      TFDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      If DeleteSpeciesName = "COHO" Then GoTo SkipFMSP

      '- FisheryModelStockProportion SELECT Statement ... CHINOOK Only
      CmdStr = "SELECT * FROM FisheryModelStockProportion WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim SPcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim SPDA As New System.Data.OleDb.OleDbDataAdapter
      SPDA.SelectCommand = SPcm
      '- TerminalFisheryFlag DELETE Statement
      CmdStr = "DELETE * FROM FisheryModelStockProportion WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim SPDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      SPDA.DeleteCommand = SPDcm
      '- Command Builder
      Dim SPcb As New OleDb.OleDbCommandBuilder
      SPcb = New OleDb.OleDbCommandBuilder(SPDA)
      FramDB.Open()
      SPDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- EncounterRateAdjustment SELECT Statement ... CHINOOK Only
      CmdStr = "SELECT * FROM EncounterRateAdjustment WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim EAcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim EADA As New System.Data.OleDb.OleDbDataAdapter
      EADA.SelectCommand = EAcm
      '- EncounterRateAdjustment DELETE Statement
      CmdStr = "DELETE * FROM EncounterRateAdjustment WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim EADcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      EADA.DeleteCommand = EADcm
      '- Command Builder
      Dim EAcb As New OleDb.OleDbCommandBuilder
      EAcb = New OleDb.OleDbCommandBuilder(EADA)
      FramDB.Open()
      EADA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- Growth SELECT Statement ... CHINOOK Only
      CmdStr = "SELECT * FROM Growth WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim Gcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim GDA As New System.Data.OleDb.OleDbDataAdapter
      GDA.SelectCommand = Gcm
      '- EncounterRateAdjustment DELETE Statement
      CmdStr = "DELETE * FROM Growth WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim GDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      GDA.DeleteCommand = GDcm
      '- Command Builder
      Dim Gcb As New OleDb.OleDbCommandBuilder
      Gcb = New OleDb.OleDbCommandBuilder(GDA)
      FramDB.Open()
      GDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- AEQ SELECT Statement ... CHINOOK Only
      CmdStr = "SELECT * FROM AEQ WHERE BasePeriodID = " & BasePeriodIDSelect.ToString
      Dim Acm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim ADA As New System.Data.OleDb.OleDbDataAdapter
      ADA.SelectCommand = Acm
      '- EncounterRateAdjustment DELETE Statement
      CmdStr = "DELETE * FROM AEQ WHERE BasePeriodID = " & BasePeriodIDSelect.ToString & ";"
      Dim ADcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      ADA.DeleteCommand = ADcm
      '- Command Builder
      Dim Acb As New OleDb.OleDbCommandBuilder
      Acb = New OleDb.OleDbCommandBuilder(ADA)
      FramDB.Open()
      ADA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

SkipFMSP:

   End Sub

   Sub ReadTaaEtrsFile()

      '--- Read TAA and ETRS Instructions from Original Text File
      ReDim TaaEtrsNum(100)
      ReDim TaaEtrsStk(100, NumStk)
      ReDim TaaEtrsFish(100, NumFish)
      ReDim TaaEtrsTStep(100, 2)
      ReDim TaaEtrsType(100)
      ReDim TaaEtrsName(100)

      '- Text File Reader
      Dim TAAReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(FVSdatabasepath & "\TaaEtrsNum.txt")
      TAAReader.TextFieldType = FileIO.FieldType.Delimited
      TAAReader.SetDelimiters(",", " ")

      Dim CurrentRow As String()
      Dim CurrentField As String
      Dim StkLine, FishLine, CmdStr As String
      Dim FieldNum As Integer
      Dim NumTaaStks, NumTaaFish, TaaNum, TAA As Integer

      '- Read TAA/ETRS Information 
      '   Terminal Number, Stocks, Fisheries, Time Steps, Type, and Terminal Name
      While Not TAAReader.EndOfData
         CurrentRow = TAAReader.ReadFields()
         FieldNum = 0
         For Each CurrentField In CurrentRow
            If CurrentField = "" Then GoTo SkipTaaField
            Select Case FieldNum
               Case 0                                  '- Terminal Area Number
                  TaaNum = CInt(CurrentField)
                  TaaEtrsNum(TaaNum) = TaaNum
               Case 1                                  '- Number of TAA Stocks
                  NumTaaStks = CInt(CurrentField)
                  TaaEtrsStk(TaaNum, 0) = NumTaaStks
               Case 2 To NumTaaStks + 1                '- TAA Stock List
                  TaaEtrsStk(TaaNum, FieldNum - 1) = CInt(CurrentField)
               Case NumTaaStks + 2                     '- Number TAA Fisheries
                  NumTaaFish = CInt(CurrentField)
                  TaaEtrsFish(TaaNum, 0) = NumTaaFish
               Case (NumTaaStks + 3) To (NumTaaStks + NumTaaFish + 2) '- TAA Fishery List
                  If NumTaaFish = 0 Then GoTo NextTaaField
                  TaaEtrsFish(TaaNum, FieldNum - (NumTaaStks + 2)) = CInt(CurrentField)
               Case (NumTaaStks + NumTaaFish + 3) '- TAA Time Step 1
                  TaaEtrsTStep(TaaNum, 1) = CInt(CurrentField)
               Case (NumTaaStks + NumTaaFish + 4) '- TAA Time Step 2
                  TaaEtrsTStep(TaaNum, 2) = CInt(CurrentField)
               Case (NumTaaStks + NumTaaFish + 5) '- TAA Terminal Type (TAA or ETRS) 
                  TaaEtrsType(TaaNum) = CInt(CurrentField)
               Case (NumTaaStks + NumTaaFish + 6) '- TAA Terminal Name
                  TaaEtrsName(TaaNum) = CurrentField
            End Select
NextTaaField:
            FieldNum += 1
SkipTaaField:
         Next
      End While

      '- Delete Existing TAA Records in TAAETRSList Table

      CmdStr = "SELECT * FROM TAAETRSList"
      Dim Taacm As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim TaaDA As New System.Data.OleDb.OleDbDataAdapter
      TaaDA.SelectCommand = Taacm
      '- EncounterRateAdjustment DELETE Statement
      CmdStr = "DELETE * FROM TAAETRSList;"
      Dim TaaDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      TaaDA.DeleteCommand = TaaDcm
      '- Command Builder
      Dim Taacb As New OleDb.OleDbCommandBuilder
      Taacb = New OleDb.OleDbCommandBuilder(TaaDA)
      FramDB.Open()
      TaaDA.DeleteCommand.ExecuteScalar()
      FramDB.Close()

      '- Store TAA Arrays in TAAETRSList Table

      Dim FramTrans As OleDb.OleDbTransaction
      Dim TaaC As New OleDbCommand
      FramDB.Open()
      FramTrans = FramDB.BeginTransaction
      TaaC.Connection = FramDB
      TaaC.Transaction = FramTrans
      For TAA = 1 To 100
         'If TaaEtrsStk(TAA, 0) = 0 Or TaaEtrsFish(TAA, 0) = 0 Then GoTo SkipTaa
         If TaaEtrsNum(TAA) = 0 Then Exit For
         '- Stock List
         StkLine = ""
         For Stk = 1 To TaaEtrsStk(TAA, 0)
            StkLine &= TaaEtrsStk(TAA, Stk).ToString
            If Stk <> TaaEtrsStk(TAA, 0) Then StkLine &= ","
         Next
         '- Fishery List
         FishLine = ""
         For Fish = 1 To TaaEtrsFish(TAA, 0)
            FishLine &= TaaEtrsFish(TAA, Fish).ToString
            If Fish <> TaaEtrsFish(TAA, 0) Then FishLine &= ","
         Next
         If TaaEtrsFish(TAA, 0) = 0 Then FishLine = "0"

         TaaC.CommandText = "INSERT INTO TAAETRSList (TaaNum,NumTaaStks,TaaStkList,NumTaaFish,TaaFishList,TaaTimeStep1,TaaTimeStep2,TaaType,TaaName) " & _
            "VALUES(" & TAA.ToString & "," & _
            TaaEtrsStk(TAA, 0).ToString & "," & _
            Chr(34) & StkLine.ToString & Chr(34) & "," & _
            TaaEtrsFish(TAA, 0).ToString & "," & _
            Chr(34) & FishLine.ToString & Chr(34) & "," & _
            TaaEtrsTStep(TAA, 1).ToString & "," & _
            TaaEtrsTStep(TAA, 2).ToString & "," & _
            TaaEtrsType(TAA).ToString & "," & _
            Chr(34) & TaaEtrsName(TAA) & Chr(34) & ")"
         TaaC.ExecuteNonQuery()
SkipTaa:
      Next
      FramTrans.Commit()
      FramDB.Close()

   End Sub

   Sub SaveModelRunInputs()

      '- Save Input Variables in Database Tables

      Dim CmdStr As String

      '- UPDATE RunID InputModified Field

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
      FramDataSet.Tables("RunID").Rows(0)(8) = DateTime.Now
      RunIDDA.Update(FramDataSet, "RunID")
      RunIDDA = Nothing

      '- Backwards Fram TARGET ESCAPEMENTS 
      If ChangeBackFram = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM BackwardsFRAM WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID"
         Dim BFcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim BackDA As New System.Data.OleDb.OleDbDataAdapter
         BackDA.SelectCommand = BFcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM BackwardsFRAM WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim BFDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         BackDA.DeleteCommand = BFDcm
         '- Command Builder
         Dim BFcb As New OleDb.OleDbCommandBuilder
         BFcb = New OleDb.OleDbCommandBuilder(BackDA)
         '- Set Up DataBase Transaction
         Dim BFTrans As OleDb.OleDbTransaction
         Dim BFC As New OleDbCommand
         FramDB.Open()
         BackDA.DeleteCommand.ExecuteScalar()
         BFTrans = FramDB.BeginTransaction
         BFC.Connection = FramDB
         BFC.Transaction = BFTrans
         '- INSERT Records into DataBase Table
         If SpeciesName = "COHO" Then
            For Stk = 1 To NumStk
               If BackwardsTarget(Stk) <> 0 Then
                  BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Stk.ToString & "," & _
                  BackwardsTarget(Stk).ToString("0.0") & ", 0, 0, " & _
                  BackwardsFlag(Stk).ToString & ")"
                  BFC.ExecuteNonQuery()
               End If
            Next
         ElseIf SpeciesName = "CHINOOK" Then
            Dim SumChinTarget As Double
            For Stk = 1 To NumStk + NumChinTermRuns
               SumChinTarget = BackwardsChinook(Stk, 3) + BackwardsChinook(Stk, 4) + BackwardsChinook(Stk, 5)
               If SumChinTarget <> 0 Then
                  BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Stk.ToString & "," & _
                  BackwardsChinook(Stk, 3).ToString("0.0") & "," & _
                  BackwardsChinook(Stk, 4).ToString("0.0") & "," & _
                  BackwardsChinook(Stk, 5).ToString("0.0") & "," & _
                  BackwardsFlag(Stk).ToString & ")"
                  BFC.ExecuteNonQuery()
               End If
            Next
         End If
         BFTrans.Commit()
         FramDB.Close()
         BackDA = Nothing
      End If

      '- Fishery Scalers
      If ChangeFishScalers = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM FisheryScalers WHERE RunID = " & RunIDSelect.ToString & " ORDER BY FisheryID, TimeStep"
         Dim FScm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim ScalerDA As New System.Data.OleDb.OleDbDataAdapter
         ScalerDA.SelectCommand = FScm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM FisheryScalers WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim FSDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         ScalerDA.DeleteCommand = FSDcm
         '- Command Builder
         Dim FScb As New OleDb.OleDbCommandBuilder
         FScb = New OleDb.OleDbCommandBuilder(ScalerDA)
         'ScalerDA.Fill(FramDataSet, "FisheryScalers")
         '- Set Up DataBase Transaction
         Dim FSTrans As OleDb.OleDbTransaction
         Dim FSC As New OleDbCommand
         FramDB.Open()
         ScalerDA.DeleteCommand.ExecuteScalar()
         FSTrans = FramDB.BeginTransaction
         FSC.Connection = FramDB
         FSC.Transaction = FSTrans
         '- INSERT Records into DataBase Table
         For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextTStep2
               If FisheryFlag(Fish, TStep) = 0 And FisheryScaler(Fish, TStep) = 0 And FisheryQuota(Fish, TStep) = 0 And MSFFisheryScaler(Fish, TStep) = 0 And MSFFisheryQuota(Fish, TStep) = 0 Then GoTo NextTStep2
                    FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                    "VALUES(" & RunIDSelect.ToString & "," & _
                    Fish.ToString & "," & _
                    TStep.ToString & "," & _
                    FisheryFlag(Fish, TStep).ToString(" #0") & "," & _
                    FisheryScaler(Fish, TStep).ToString(" ####0.0000") & "," & _
                    FisheryQuota(Fish, TStep).ToString(" ######0.0000") & "," & _
                    MSFFisheryScaler(Fish, TStep).ToString(" ####0.0000") & "," & _
                    MSFFisheryQuota(Fish, TStep).ToString(" ######0.0000") & "," & _
                    MarkSelectiveMortRate(Fish, TStep).ToString(" #0.0000") & "," & _
                    MarkSelectiveMarkMisID(Fish, TStep).ToString(" #0.0000") & "," & _
                    MarkSelectiveUnMarkMisID(Fish, TStep).ToString(" #0.0000") & "," & _
                    MarkSelectiveIncRate(Fish, TStep).ToString(" #0.0000") & ") ;"
               FSC.ExecuteNonQuery()
NextTStep2:
            Next
         Next
         FSTrans.Commit()
         FramDB.Close()
         ScalerDA = Nothing
      End If

      '- Non Retention
      If ChangeNonRetention = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM NonRetention WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, TimeStep"
         Dim NRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim NonRetDA As New System.Data.OleDb.OleDbDataAdapter
         NonRetDA.SelectCommand = NRcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM NonRetention WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim NRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         NonRetDA.DeleteCommand = NRDcm
         '- Command Builder
         Dim NRcb As New OleDb.OleDbCommandBuilder
         NRcb = New OleDb.OleDbCommandBuilder(NonRetDA)
         'ScalerDA.Fill(FramDataSet, "NonRetention")
         '- Set Up DataBase Transaction
         Dim NRTrans As OleDb.OleDbTransaction
         Dim NRC As New OleDbCommand
         FramDB.Open()
         NonRetDA.DeleteCommand.ExecuteScalar()
         NRTrans = FramDB.BeginTransaction
         NRC.Connection = FramDB
         NRC.Transaction = NRTrans
         '- INSERT Records into DataBase Table
         For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextTStep4
               If NonRetentionFlag(Fish, TStep) = 0 And NonRetentionInput(Fish, TStep, 1) = 0 Then GoTo NextTStep4
               If SpeciesName = "COHO" Then
                  NRC.CommandText = "INSERT INTO NonRetention (RunID,FisheryID,TimeStep,NonRetentionFlag,CNRInput1) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  NonRetentionFlag(Fish, TStep).ToString(" #0") & "," & _
                  NonRetentionInput(Fish, TStep, 1).ToString(" ####0.0000") & ") ;"
               ElseIf SpeciesName = "CHINOOK" Then
                  NRC.CommandText = "INSERT INTO NonRetention (RunID,FisheryID,TimeStep,NonRetentionFlag,CNRInput1,CNRInput2,CNRInput3,CNRInput4) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  NonRetentionFlag(Fish, TStep).ToString(" #0") & "," & _
                  NonRetentionInput(Fish, TStep, 1).ToString(" ####0.0000") & "," & _
                  NonRetentionInput(Fish, TStep, 2).ToString(" ####0.0000") & "," & _
                  NonRetentionInput(Fish, TStep, 3).ToString(" ####0.0000") & "," & _
                  NonRetentionInput(Fish, TStep, 4).ToString(" ####0.0000") & ") ;"
               End If
               NRC.ExecuteNonQuery()
NextTStep4:
            Next
         Next
         NRTrans.Commit()
         FramDB.Close()
         NonRetDA = Nothing
      End If

      '- Stock/Fishery Rate Scalers
      If ChangeStockFishScaler = True Then
         '- DataApapter SELECT Statement
         'CmdStr = "SELECT * FROM StockFisheryRateScaler WHERE RunID = " & RunIDSelect.ToString & " AND FisheryID = " & FisheryEditSelection.ToString & " ORDER BY StockID, FisheryID, TimeStep"
         CmdStr = "SELECT * FROM StockFisheryRateScaler WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, FisheryID, TimeStep"
         Dim SFRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim StockFisheryDA As New System.Data.OleDb.OleDbDataAdapter
         StockFisheryDA.SelectCommand = SFRcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM StockFisheryRateScaler WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim SFRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         StockFisheryDA.DeleteCommand = SFRDcm
         '- Command Builder
         Dim SFRcb As New OleDb.OleDbCommandBuilder
         SFRcb = New OleDb.OleDbCommandBuilder(StockFisheryDA)
         'StockFisheryDA.Fill(FramDataSet, "NonRetention")
         '- Set Up DataBase Transaction
         Dim SFRTrans As OleDb.OleDbTransaction
         Dim SFRC As New OleDbCommand
         FramDB.Open()
         StockFisheryDA.DeleteCommand.ExecuteScalar()
         SFRTrans = FramDB.BeginTransaction
         SFRC.Connection = FramDB
         SFRC.Transaction = SFRTrans
         '- INSERT Records into DataBase Table
         For Stk = 1 To NumStk
            For Fish = 1 To NumFish
               For TStep = 1 To NumSteps
                  If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextTStep5
                  If StockFishRateScalers(Stk, Fish, TStep) = 1 Then GoTo NextTStep5
                  SFRC.CommandText = "INSERT INTO StockFisheryRateScaler (RunID,StockID,FisheryID,TimeStep,StockFisheryRateScaler) " & _
                  "VALUES(" & RunIDSelect.ToString & "," & _
                  Stk.ToString & "," & _
                  Fish.ToString & "," & _
                  TStep.ToString & "," & _
                  StockFishRateScalers(Stk, Fish, TStep).ToString(" ####0.0000") & ") ;"
                  SFRC.ExecuteNonQuery()
NextTStep5:
               Next
            Next
         Next
         SFRTrans.Commit()
         FramDB.Close()
         StockFisheryDA = Nothing
      End If

      '- PSCMaxER - Coho Only
      If ChangePSCMaxER = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM PSCMaxER WHERE RunID = " & RunIDSelect.ToString & " ORDER BY PSCStockID"
         Dim PSCcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim PSCDA As New System.Data.OleDb.OleDbDataAdapter
         PSCDA.SelectCommand = PSCcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM PSCMaxER WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim PSCDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         PSCDA.DeleteCommand = PSCDcm
         '- Command Builder
         Dim PSCcb As New OleDb.OleDbCommandBuilder
         PSCcb = New OleDb.OleDbCommandBuilder(PSCDA)
         'StockFisheryDA.Fill(FramDataSet, "NonRetention")
         '- Set Up DataBase Transaction
         Dim PSCTrans As OleDb.OleDbTransaction
         Dim PSCC As New OleDbCommand
         FramDB.Open()
         PSCDA.DeleteCommand.ExecuteScalar()
         PSCTrans = FramDB.BeginTransaction
         PSCC.Connection = FramDB
         PSCC.Transaction = PSCTrans
         '- INSERT Records into DataBase Table
         For Stk = 1 To 17
            PSCC.CommandText = "INSERT INTO PSCMaxER (RunID,PSCStockID,PSCMaxER) " & _
            "VALUES(" & RunIDSelect.ToString & "," & _
            Stk.ToString & "," & _
            PSCMaxER(Stk).ToString(" ####0.0000") & ") ;"
            PSCC.ExecuteNonQuery()
         Next
         PSCTrans.Commit()
         FramDB.Close()
         PSCDA = Nothing
      End If

      '- Size Limits - Chinook Only
      If ChangeSizeLimit = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM SizeLimits WHERE RunID = " & RunIDSelect.ToString & " ORDER BY FisheryID, TimeStep"
         Dim SLcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim SLDA As New System.Data.OleDb.OleDbDataAdapter
         SLDA.SelectCommand = SLcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM SizeLimits WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim SLDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         SLDA.DeleteCommand = SLDcm
         '- Command Builder
         Dim SLcb As New OleDb.OleDbCommandBuilder
         SLcb = New OleDb.OleDbCommandBuilder(SLDA)
         'StockFisheryDA.Fill(FramDataSet, "NonRetention")
         '- Set Up DataBase Transaction
         Dim SLTrans As OleDb.OleDbTransaction
         Dim SLC As New OleDbCommand
         FramDB.Open()
         SLDA.DeleteCommand.ExecuteScalar()
         SLTrans = FramDB.BeginTransaction
         SLC.Connection = FramDB
         SLC.Transaction = SLTrans
         '- INSERT Records into DataBase Table
         For Fish = 1 To NumFish
            For TStep = 1 To NumSteps
               SLC.CommandText = "INSERT INTO SizeLimits (RunID,FisheryID,TimeStep,MinimumSize,MaximumSize) " & _
               "VALUES(" & RunIDSelect.ToString & "," & _
               Fish.ToString & "," & _
               TStep.ToString & "," & _
               MinSizeLimit(Fish, TStep).ToString(" ####0") & "," & _
               MinSizeLimit(Fish, TStep).ToString(" ####0") & ") ;"
               SLC.ExecuteNonQuery()
            Next
         Next
         SLTrans.Commit()
         FramDB.Close()
         SLDA = Nothing
      End If

      '- Stock Recruits
      If ChangeStockRecruit = True Then
         '- DataApapter SELECT Statement
         CmdStr = "SELECT * FROM StockRecruit WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age"
         Dim SRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim SRDA As New System.Data.OleDb.OleDbDataAdapter
         SRDA.SelectCommand = SRcm
         '- DataApapter DELETE Statement
         CmdStr = "DELETE * FROM StockRecruit WHERE RunID = " & RunIDSelect.ToString & ";"
         Dim SRDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         SRDA.DeleteCommand = SRDcm
         '- Command Builder
         Dim SRcb As New OleDb.OleDbCommandBuilder
         SRcb = New OleDb.OleDbCommandBuilder(SRDA)
         'StockFisheryDA.Fill(FramDataSet, "NonRetention")
         '- Set Up DataBase Transaction
         Dim SRTrans As OleDb.OleDbTransaction
         Dim SRC As New OleDbCommand
         FramDB.Open()
         SRDA.DeleteCommand.ExecuteScalar()
         SRTrans = FramDB.BeginTransaction
         SRC.Connection = FramDB
         SRC.Transaction = SRTrans
         '- INSERT Records into DataBase Table
         For Stk = 1 To NumStk
            For Age = MinAge To MaxAge
               If StockRecruit(Stk, Age, 1) = 0 Then GoTo SkipSR
               SRC.CommandText = "INSERT INTO StockRecruit (RunID,StockID,Age,RecruitScaleFactor,RecruitCohortSize) " & _
               "VALUES(" & RunIDSelect.ToString & "," & _
               Stk.ToString & "," & _
               Age.ToString & "," & _
               StockRecruit(Stk, Age, 1).ToString(" ####0.0000") & "," & _
               StockRecruit(Stk, Age, 2).ToString(" ####0") & ") ;"
               SRC.ExecuteNonQuery()
SkipSR:
            Next
         Next
         SRTrans.Commit()
         FramDB.Close()
         SRDA = Nothing
      End If

      '- Set Edit Change Variables to False
      ChangeAnyInput = False
      ChangeBackFram = False
      ChangeFishScalers = False
      ChangeNonRetention = False
      ChangePSCMaxER = False
      ChangeSizeLimit = False
      ChangeStockFishScaler = False
      ChangeStockRecruit = False

   End Sub
    Sub TransferBasePeriodTables()
        Dim CmdStr As String
        Dim TransID, RecNum, NumRecs, TransferBaseID As Integer

        'Loop through User Selected RunID Transfers
        For TransID = 0 To NumTransferID - 1

            '- Transfer BaseID Record
            TransferBaseID = RunIDTransfer(TransID)
            CmdStr = "SELECT * FROM BaseID WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim BIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim BaseIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseIDDA.SelectCommand = BIDcm
            Dim BIDcb As New OleDb.OleDbCommandBuilder
            BIDcb = New OleDb.OleDbCommandBuilder(BaseIDDA)
            If TransferDataSet.Tables.Contains("BaseID") Then
                TransferDataSet.Tables("BaseID").Clear()
            End If
            BaseIDDA.Fill(TransferDataSet, "BaseID")
            Dim NumBID As Integer
            NumBID = TransferDataSet.Tables("BaseID").Rows.Count
            If NumBID <> 1 Then
                MsgBox("ERROR in BaseID Table of Database ... Duplicate Record", MsgBoxStyle.OkOnly)
            End If
            Dim BIDTrans As OleDb.OleDbTransaction
            Dim BID As New OleDbCommand
            TransDB.Open()
            BIDTrans = TransDB.BeginTransaction
            BID.Connection = TransDB
            BID.Transaction = BIDTrans
            RecNum = 0
            BID.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
               "VALUES(" & TransferDataSet.Tables("BaseID").Rows(RecNum)(1) & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(2) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(3) & Chr(34) & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(4).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(5).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(6).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(7).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(8).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(9).ToString & "," & _
               Chr(35) & TransferDataSet.Tables("BaseID").Rows(RecNum)(10) & Chr(35) & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(11) & Chr(34) & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(12).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(13).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(RecNum)(14).ToString & ")"
            BID.ExecuteNonQuery()
            BIDTrans.Commit()
            TransDB.Close()

           

            'Base Cohort
            CmdStr = "SELECT * FROM BaseCohort WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim BaseCohortcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim BaseCohortIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseCohortIDDA.SelectCommand = BaseCohortcm
            Dim BaseCohortcb As New OleDb.OleDbCommandBuilder
            BaseCohortcb = New OleDb.OleDbCommandBuilder(BaseCohortIDDA)
            If TransferDataSet.Tables.Contains("BaseCohort") Then
                TransferDataSet.Tables("BaseCohort").Clear()
            End If
            BaseCohortIDDA.Fill(TransferDataSet, "BaseCohort")
            Dim NumBaseCohort As Integer
            NumBaseCohort = TransferDataSet.Tables("BaseCohort").Rows.Count

            Dim BaseCohortTrans As OleDb.OleDbTransaction
            Dim BaseCohort As New OleDbCommand
            TransDB.Open()
            BaseCohortTrans = TransDB.BeginTransaction
            BaseCohort.Connection = TransDB
            BaseCohort.Transaction = BaseCohortTrans
            NumRecs = TransferDataSet.Tables("BaseCohort").Rows.Count
            For RecNum = 0 To NumRecs - 1
                BaseCohort.CommandText = "INSERT INTO BaseCohort (BasePeriodID,StockID,Age,BaseCohortSize) " & _
                   "VALUES(" & TransferDataSet.Tables("BaseCohort").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("BaseCohort").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("BaseCohort").Rows(RecNum)(2) & "," & _
                   TransferDataSet.Tables("BaseCohort").Rows(RecNum)(3) & ")"

                BaseCohort.ExecuteNonQuery()
            Next
            BaseCohortTrans.Commit()
            TransDB.Close()


            'AEQ
            If SpeciesName = "CHINOOK" Then
                CmdStr = "SELECT * FROM AEQ WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
                Dim AEQcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim AEQIDDA As New System.Data.OleDb.OleDbDataAdapter
                AEQIDDA.SelectCommand = AEQcm
                Dim AEQcb As New OleDb.OleDbCommandBuilder
                AEQcb = New OleDb.OleDbCommandBuilder(AEQIDDA)
                If TransferDataSet.Tables.Contains("AEQ") Then
                    TransferDataSet.Tables("AEQ").Clear()
                End If
                AEQIDDA.Fill(TransferDataSet, "AEQ")
                Dim NumAEQ As Integer
                NumAEQ = TransferDataSet.Tables("AEQ").Rows.Count

                Dim AEQTrans As OleDb.OleDbTransaction
                Dim AEQ As New OleDbCommand
                TransDB.Open()
                AEQTrans = TransDB.BeginTransaction
                AEQ.Connection = TransDB
                AEQ.Transaction = AEQTrans
                NumRecs = TransferDataSet.Tables("AEQ").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    AEQ.CommandText = "INSERT INTO AEQ (BasePeriodID,StockID,Age,TimeStep, AEQ) " & _
                       "VALUES(" & TransferDataSet.Tables("AEQ").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("AEQ").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("AEQ").Rows(RecNum)(2) & "," & _
                        TransferDataSet.Tables("AEQ").Rows(RecNum)(3) & "," & _
                       TransferDataSet.Tables("AEQ").Rows(RecNum)(4) & ")"
                    AEQ.ExecuteNonQuery()
                Next
                AEQTrans.Commit()
                TransDB.Close()
            End If

            'BaseExploitationRate
            CmdStr = "SELECT * FROM BaseExploitationRate WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim BaseExploitationRatecm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim BaseExploitationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseExploitationRateIDDA.SelectCommand = BaseExploitationRatecm
            Dim BaseExploitationRatecb As New OleDb.OleDbCommandBuilder
            BaseExploitationRatecb = New OleDb.OleDbCommandBuilder(BaseExploitationRateIDDA)
            If TransferDataSet.Tables.Contains("BaseExploitationRate") Then
                TransferDataSet.Tables("BaseExploitationRate").Clear()
            End If
            BaseExploitationRateIDDA.Fill(TransferDataSet, "BaseExploitationRate")
            Dim NumBaseExploitationRate As Integer
            NumBaseExploitationRate = TransferDataSet.Tables("BaseExploitationRate").Rows.Count

            Dim BaseExploitationRateTrans As OleDb.OleDbTransaction
            Dim BaseExploitationRate As New OleDbCommand
            TransDB.Open()
            BaseExploitationRateTrans = TransDB.BeginTransaction
            BaseExploitationRate.Connection = TransDB
            BaseExploitationRate.Transaction = BaseExploitationRateTrans
            NumRecs = TransferDataSet.Tables("BaseExploitationRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                BaseExploitationRate.CommandText = "INSERT INTO BaseExploitationRate (BasePeriodID,StockID,Age,FisheryID,TimeStep,ExploitationRate,SublegalEncounterRate) " & _
                   "VALUES(" & TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(3) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(4) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(5) & "," & _
                   TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(6) & ")"

                BaseExploitationRate.ExecuteNonQuery()
            Next
            BaseExploitationRateTrans.Commit()
            TransDB.Close()

            'ChinookBaseEncounterAdjustment
            If SpeciesName = "CHINOOK" Then
                CmdStr = "SELECT * FROM ChinookBaseEncounterAdjustment"
                Dim ChinookBaseEncounterAdjustmentcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim ChinookBaseEncounterAdjustmentIDDA As New System.Data.OleDb.OleDbDataAdapter
                ChinookBaseEncounterAdjustmentIDDA.SelectCommand = ChinookBaseEncounterAdjustmentcm
                Dim ChinookBaseEncounterAdjustmentcb As New OleDb.OleDbCommandBuilder
                ChinookBaseEncounterAdjustmentcb = New OleDb.OleDbCommandBuilder(ChinookBaseEncounterAdjustmentIDDA)
                If TransferDataSet.Tables.Contains("ChinookBaseEncounterAdjustment") Then
                    TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Clear()
                End If
                ChinookBaseEncounterAdjustmentIDDA.Fill(TransferDataSet, "ChinookBaseEncounterAdjustment")
                Dim NumChinookBaseEncounterAdjustment As Integer
                NumChinookBaseEncounterAdjustment = TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows.Count

                Dim ChinookBaseEncounterAdjustmentTrans As OleDb.OleDbTransaction
                Dim ChinookBaseEncounterAdjustment As New OleDbCommand
                TransDB.Open()
                ChinookBaseEncounterAdjustmentTrans = TransDB.BeginTransaction
                ChinookBaseEncounterAdjustment.Connection = TransDB
                ChinookBaseEncounterAdjustment.Transaction = ChinookBaseEncounterAdjustmentTrans
                NumRecs = TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    ChinookBaseEncounterAdjustment.CommandText = "INSERT INTO ChinookBaseEncounterAdjustment (FisheryID,Time1Adjustment,Time2Adjustment,Time3Adjustment,Time4Adjustment) " & _
                       "VALUES(" & TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(2) & "," & _
                        TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(3) & "," & _
                       TransferDataSet.Tables("ChinookBaseEncounterAdjustment").Rows(RecNum)(4) & ")"

                    ChinookBaseEncounterAdjustment.ExecuteNonQuery()
                Next
                ChinookBaseEncounterAdjustmentTrans.Commit()
                TransDB.Close()
            End If

            'ChinookBaseSizeLimit
            If SpeciesName = "CHINOOK" Then
                CmdStr = "SELECT * FROM ChinookBaseSizeLimit"
                Dim ChinookBaseSizeLimitcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim ChinookBaseSizeLimitIDDA As New System.Data.OleDb.OleDbDataAdapter
                ChinookBaseSizeLimitIDDA.SelectCommand = ChinookBaseSizeLimitcm
                Dim ChinookBaseSizeLimitcb As New OleDb.OleDbCommandBuilder
                ChinookBaseSizeLimitcb = New OleDb.OleDbCommandBuilder(ChinookBaseSizeLimitIDDA)
                If TransferDataSet.Tables.Contains("ChinookBaseSizeLimit") Then
                    TransferDataSet.Tables("ChinookBaseSizeLimit").Clear()
                End If
                ChinookBaseSizeLimitIDDA.Fill(TransferDataSet, "ChinookBaseSizeLimit")
                Dim NumChinookBaseSizeLimit As Integer
                NumChinookBaseSizeLimit = TransferDataSet.Tables("ChinookBaseSizeLimit").Rows.Count

                Dim ChinookBaseSizeLimitTrans As OleDb.OleDbTransaction
                Dim ChinookBaseSizeLimit As New OleDbCommand
                TransDB.Open()
                ChinookBaseSizeLimitTrans = TransDB.BeginTransaction
                ChinookBaseSizeLimit.Connection = TransDB
                ChinookBaseSizeLimit.Transaction = ChinookBaseSizeLimitTrans
                NumRecs = TransferDataSet.Tables("ChinookBaseSizeLimit").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    ChinookBaseSizeLimit.CommandText = "INSERT INTO ChinookBaseSizeLimit (FisheryID,Time1SizeLimit,Time2SizeLimit,Time3SizeLimit,Time4SizeLimit) " & _
                       "VALUES(" & TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(2) & "," & _
                        TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(3) & "," & _
                       TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(4) & ")"

                    ChinookBaseSizeLimit.ExecuteNonQuery()
                Next
                ChinookBaseSizeLimitTrans.Commit()
                TransDB.Close()

                'EncounterRateAdjustment
                CmdStr = "SELECT * FROM EncounterRateAdjustment WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
                Dim EncounterRateAdjustmentcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim EncounterRateAdjustmentIDDA As New System.Data.OleDb.OleDbDataAdapter
                EncounterRateAdjustmentIDDA.SelectCommand = EncounterRateAdjustmentcm
                Dim EncounterRateAdjustmentcb As New OleDb.OleDbCommandBuilder
                EncounterRateAdjustmentcb = New OleDb.OleDbCommandBuilder(EncounterRateAdjustmentIDDA)
                If TransferDataSet.Tables.Contains("EncounterRateAdjustment") Then
                    TransferDataSet.Tables("EncounterRateAdjustment").Clear()
                End If
                EncounterRateAdjustmentIDDA.Fill(TransferDataSet, "EncounterRateAdjustment")
                Dim NumEncounterRateAdjustment As Integer
                NumEncounterRateAdjustment = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count

                Dim EncounterRateAdjustmentTrans As OleDb.OleDbTransaction
                Dim EncounterRateAdjustment As New OleDbCommand
                TransDB.Open()
                EncounterRateAdjustmentTrans = TransDB.BeginTransaction
                EncounterRateAdjustment.Connection = TransDB
                EncounterRateAdjustment.Transaction = EncounterRateAdjustmentTrans
                NumRecs = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    EncounterRateAdjustment.CommandText = "INSERT INTO EncounterRateAdjustment (BasePeriodID,Age,FisheryID,TimeStep,EncounterRateAdjustment) " & _
                       "VALUES(" & TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(3) & "," & _
                       TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(4) & ")"

                    EncounterRateAdjustment.ExecuteNonQuery()
                Next
                EncounterRateAdjustmentTrans.Commit()
                TransDB.Close()



                'Fishery
                CmdStr = "SELECT * FROM Fishery"
                Dim Fisherycm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim FisheryIDDA As New System.Data.OleDb.OleDbDataAdapter
                FisheryIDDA.SelectCommand = Fisherycm
                Dim Fisherycb As New OleDb.OleDbCommandBuilder
                Fisherycb = New OleDb.OleDbCommandBuilder(FisheryIDDA)
                If TransferDataSet.Tables.Contains("Fishery") Then
                    TransferDataSet.Tables("Fishery").Clear()
                End If
                FisheryIDDA.Fill(TransferDataSet, "Fishery")
                Dim NumFishery As Integer
                NumFishery = TransferDataSet.Tables("Fishery").Rows.Count

                Dim FisheryTrans As OleDb.OleDbTransaction
                Dim Fishery As New OleDbCommand
                TransDB.Open()
                FisheryTrans = TransDB.BeginTransaction
                Fishery.Connection = TransDB
                Fishery.Transaction = FisheryTrans
                NumRecs = TransferDataSet.Tables("Fishery").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    Fishery.CommandText = "INSERT INTO Fishery (Species,VersionNumber,FisheryID,FisheryName,FisheryTitle) " & _
                       "VALUES(" & Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(0) & Chr(34) & "," & _
                        TransferDataSet.Tables("Fishery").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("Fishery").Rows(RecNum)(2) & "," & _
                        Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(3) & Chr(34) & "," & _
                       Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(4) & Chr(34) & ")"

                    Fishery.ExecuteNonQuery()
                Next
                FisheryTrans.Commit()
                TransDB.Close()

                'FisheryModelStockProportion
                CmdStr = "SELECT * FROM FisheryModelStockProportion WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
                Dim FisheryModelStockProportioncm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim FisheryModelStockProportionIDDA As New System.Data.OleDb.OleDbDataAdapter
                FisheryModelStockProportionIDDA.SelectCommand = FisheryModelStockProportioncm
                Dim FisheryModelStockProportioncb As New OleDb.OleDbCommandBuilder
                FisheryModelStockProportioncb = New OleDb.OleDbCommandBuilder(FisheryModelStockProportionIDDA)
                If TransferDataSet.Tables.Contains("FisheryModelStockProportion") Then
                    TransferDataSet.Tables("FisheryModelStockProportion").Clear()
                End If
                FisheryModelStockProportionIDDA.Fill(TransferDataSet, "FisheryModelStockProportion")
                Dim NumFisheryModelStockProportion As Integer
                NumFisheryModelStockProportion = TransferDataSet.Tables("FisheryModelStockProportion").Rows.Count

                Dim FisheryModelStockProportionTrans As OleDb.OleDbTransaction
                Dim FisheryModelStockProportion As New OleDbCommand
                TransDB.Open()
                FisheryModelStockProportionTrans = TransDB.BeginTransaction
                FisheryModelStockProportion.Connection = TransDB
                FisheryModelStockProportion.Transaction = FisheryModelStockProportionTrans
                NumRecs = TransferDataSet.Tables("FisheryModelStockProportion").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    FisheryModelStockProportion.CommandText = "INSERT INTO FisheryModelStockProportion (BasePeriodID,FisheryID,ModelStockProportion) " & _
                       "VALUES(" & TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(1) & "," & _
                       TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(2) & ")"

                    FisheryModelStockProportion.ExecuteNonQuery()
                Next
                FisheryModelStockProportionTrans.Commit()
                TransDB.Close()

                'Growth
                CmdStr = "SELECT * FROM Growth WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
                Dim Growthcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim GrowthIDDA As New System.Data.OleDb.OleDbDataAdapter
                GrowthIDDA.SelectCommand = Growthcm
                Dim Growthcb As New OleDb.OleDbCommandBuilder
                Growthcb = New OleDb.OleDbCommandBuilder(GrowthIDDA)
                If TransferDataSet.Tables.Contains("Growth") Then
                    TransferDataSet.Tables("Growth").Clear()
                End If
                GrowthIDDA.Fill(TransferDataSet, "Growth")
                Dim NumGrowth As Integer
                NumGrowth = TransferDataSet.Tables("Growth").Rows.Count

                Dim GrowthTrans As OleDb.OleDbTransaction
                Dim Growth As New OleDbCommand
                TransDB.Open()
                GrowthTrans = TransDB.BeginTransaction
                Growth.Connection = TransDB
                Growth.Transaction = GrowthTrans
                NumRecs = TransferDataSet.Tables("Growth").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    Growth.CommandText = "INSERT INTO Growth (BasePeriodID,StockID,LImmature,KImmature,TImmature,CV2Immature,CV3Immature,CV4Immature,CV5Immature,LMature,KMature,TMature,CV2Mature,CV3Mature,CV4Mature,CV5Mature) " & _
                       "VALUES(" & TransferDataSet.Tables("Growth").Rows(RecNum)(0) & "," & _
                        TransferDataSet.Tables("Growth").Rows(RecNum)(1) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(2) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(3) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(4) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(5) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(6) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(7) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(8) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(9) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(10) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(11) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(12) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(13) & "," & _
                     TransferDataSet.Tables("Growth").Rows(RecNum)(14) & "," & _
                       TransferDataSet.Tables("Growth").Rows(RecNum)(15) & ")"

                    Growth.ExecuteNonQuery()
                Next
                GrowthTrans.Commit()
                TransDB.Close()
            End If


            'IncidentalRate
            CmdStr = "SELECT * FROM IncidentalRate WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim IncidentalRatecm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim IncidentalRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            IncidentalRateIDDA.SelectCommand = IncidentalRatecm
            Dim IncidentalRatecb As New OleDb.OleDbCommandBuilder
            IncidentalRatecb = New OleDb.OleDbCommandBuilder(IncidentalRateIDDA)
            If TransferDataSet.Tables.Contains("IncidentalRate") Then
                TransferDataSet.Tables("IncidentalRate").Clear()
            End If
            IncidentalRateIDDA.Fill(TransferDataSet, "IncidentalRate")
            Dim NumIncidentalRate As Integer
            NumIncidentalRate = TransferDataSet.Tables("IncidentalRate").Rows.Count

            Dim IncidentalRateTrans As OleDb.OleDbTransaction
            Dim IncidentalRate As New OleDbCommand
            TransDB.Open()
            IncidentalRateTrans = TransDB.BeginTransaction
            IncidentalRate.Connection = TransDB
            IncidentalRate.Transaction = IncidentalRateTrans
            NumRecs = TransferDataSet.Tables("IncidentalRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                IncidentalRate.CommandText = "INSERT INTO IncidentalRate (BasePeriodID,FisheryID,TimeStep,IncidentalRate) " & _
                   "VALUES(" & TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(3) & ")"

                IncidentalRate.ExecuteNonQuery()
            Next
            IncidentalRateTrans.Commit()
            TransDB.Close()

            'MaturationRate

            CmdStr = "SELECT * FROM MaturationRate WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim MaturationRatecm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim MaturationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            MaturationRateIDDA.SelectCommand = MaturationRatecm
            Dim MaturationRatecb As New OleDb.OleDbCommandBuilder
            MaturationRatecb = New OleDb.OleDbCommandBuilder(MaturationRateIDDA)
            If TransferDataSet.Tables.Contains("MaturationRate") Then
                TransferDataSet.Tables("MaturationRate").Clear()
            End If
            MaturationRateIDDA.Fill(TransferDataSet, "MaturationRate")
            Dim NumMaturationRate As Integer
            NumMaturationRate = TransferDataSet.Tables("MaturationRate").Rows.Count

            Dim MaturationRateTrans As OleDb.OleDbTransaction
            Dim MaturationRate As New OleDbCommand
            TransDB.Open()
            MaturationRateTrans = TransDB.BeginTransaction
            MaturationRate.Connection = TransDB
            MaturationRate.Transaction = MaturationRateTrans
            NumRecs = TransferDataSet.Tables("MaturationRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                MaturationRate.CommandText = "INSERT INTO MaturationRate (BasePeriodID,StockID,Age,TimeStep,MaturationRate) " & _
                   "VALUES(" & TransferDataSet.Tables("MaturationRate").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(3) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(4) & ")"

                MaturationRate.ExecuteNonQuery()
            Next
            MaturationRateTrans.Commit()
            TransDB.Close()

            'NaturalMortality
            CmdStr = "SELECT * FROM NaturalMortality WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim NaturalMortalitycm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim NaturalMortalityIDDA As New System.Data.OleDb.OleDbDataAdapter
            NaturalMortalityIDDA.SelectCommand = NaturalMortalitycm
            Dim NaturalMortalitycb As New OleDb.OleDbCommandBuilder
            NaturalMortalitycb = New OleDb.OleDbCommandBuilder(NaturalMortalityIDDA)
            If TransferDataSet.Tables.Contains("NaturalMortality") Then
                TransferDataSet.Tables("NaturalMortality").Clear()
            End If
            NaturalMortalityIDDA.Fill(TransferDataSet, "NaturalMortality")
            Dim NumNaturalMortality As Integer
            NumNaturalMortality = TransferDataSet.Tables("NaturalMortality").Rows.Count

            Dim NaturalMortalityTrans As OleDb.OleDbTransaction
            Dim NaturalMortality As New OleDbCommand
            TransDB.Open()
            NaturalMortalityTrans = TransDB.BeginTransaction
            NaturalMortality.Connection = TransDB
            NaturalMortality.Transaction = NaturalMortalityTrans
            NumRecs = TransferDataSet.Tables("NaturalMortality").Rows.Count
            For RecNum = 0 To NumRecs - 1
                NaturalMortality.CommandText = "INSERT INTO NaturalMortality (BasePeriodID,Age,TimeStep,NaturalMortalityRate) " & _
                   "VALUES(" & TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(3) & ")"

                NaturalMortality.ExecuteNonQuery()
            Next
            NaturalMortalityTrans.Commit()
            TransDB.Close()

            'ShakerMortRate
            CmdStr = "SELECT * FROM ShakerMortRate WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim ShakerMortRatecm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim ShakerMortRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            ShakerMortRateIDDA.SelectCommand = ShakerMortRatecm
            Dim ShakerMortRatecb As New OleDb.OleDbCommandBuilder
            ShakerMortRatecb = New OleDb.OleDbCommandBuilder(ShakerMortRateIDDA)
            If TransferDataSet.Tables.Contains("ShakerMortRate") Then
                TransferDataSet.Tables("ShakerMortRate").Clear()
            End If
            ShakerMortRateIDDA.Fill(TransferDataSet, "ShakerMortRate")
            Dim NumShakerMortRate As Integer
            NumShakerMortRate = TransferDataSet.Tables("ShakerMortRate").Rows.Count

            Dim ShakerMortRateTrans As OleDb.OleDbTransaction
            Dim ShakerMortRate As New OleDbCommand
            TransDB.Open()
            ShakerMortRateTrans = TransDB.BeginTransaction
            ShakerMortRate.Connection = TransDB
            ShakerMortRate.Transaction = ShakerMortRateTrans
            NumRecs = TransferDataSet.Tables("ShakerMortRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                ShakerMortRate.CommandText = "INSERT INTO ShakerMortRate (BasePeriodID,FisheryID,TimeStep,ShakerMortRate) " & _
                   "VALUES(" & TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(3) & ")"

                ShakerMortRate.ExecuteNonQuery()
            Next
            ShakerMortRateTrans.Commit()
            TransDB.Close()

            'Stock
            If SpeciesName = "CHINOOK" Then
                CmdStr = "SELECT * FROM Stock"
                Dim Stockcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim StockIDDA As New System.Data.OleDb.OleDbDataAdapter
                StockIDDA.SelectCommand = Stockcm
                Dim Stockcb As New OleDb.OleDbCommandBuilder
                Stockcb = New OleDb.OleDbCommandBuilder(StockIDDA)
                If TransferDataSet.Tables.Contains("Stock") Then
                    TransferDataSet.Tables("Stock").Clear()
                End If
                StockIDDA.Fill(TransferDataSet, "Stock")
                Dim NumStock As Integer
                NumStock = TransferDataSet.Tables("Stock").Rows.Count

                Dim StockTrans As OleDb.OleDbTransaction
                Dim Stock As New OleDbCommand
                TransDB.Open()
                StockTrans = TransDB.BeginTransaction
                Stock.Connection = TransDB
                Stock.Transaction = StockTrans
                NumRecs = TransferDataSet.Tables("Stock").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    Stock.CommandText = "INSERT INTO Stock (Species,StockVersion,StockID,ProductionRegionNumber,ManagementUnitNumber,StockName,StockLongName) " & _
                       "VALUES(" & Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(0) & Chr(34) & "," & _
                        TransferDataSet.Tables("Stock").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("Stock").Rows(RecNum)(2) & "," & _
                        TransferDataSet.Tables("Stock").Rows(RecNum)(3) & "," & _
                        TransferDataSet.Tables("Stock").Rows(RecNum)(4) & "," & _
                        Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(5) & Chr(34) & "," & _
                       Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(6) & Chr(34) & ")"

                    Stock.ExecuteNonQuery()
                Next
                StockTrans.Commit()
                TransDB.Close()
            End If

            'TerminalFisheryFlag
            CmdStr = "SELECT * FROM TerminalFisheryFlag WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
            Dim TerminalFisheryFlagcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim TerminalFisheryFlagIDDA As New System.Data.OleDb.OleDbDataAdapter
            TerminalFisheryFlagIDDA.SelectCommand = TerminalFisheryFlagcm
            Dim TerminalFisheryFlagcb As New OleDb.OleDbCommandBuilder
            TerminalFisheryFlagcb = New OleDb.OleDbCommandBuilder(TerminalFisheryFlagIDDA)
            If TransferDataSet.Tables.Contains("TerminalFisheryFlag") Then
                TransferDataSet.Tables("TerminalFisheryFlag").Clear()
            End If
            TerminalFisheryFlagIDDA.Fill(TransferDataSet, "TerminalFisheryFlag")
            Dim NumTerminalFisheryFlag As Integer
            NumTerminalFisheryFlag = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count

            Dim TerminalFisheryFlagTrans As OleDb.OleDbTransaction
            Dim TerminalFisheryFlag As New OleDbCommand
            TransDB.Open()
            TerminalFisheryFlagTrans = TransDB.BeginTransaction
            TerminalFisheryFlag.Connection = TransDB
            TerminalFisheryFlag.Transaction = TerminalFisheryFlagTrans
            NumRecs = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count
            For RecNum = 0 To NumRecs - 1
                TerminalFisheryFlag.CommandText = "INSERT INTO TerminalFisheryFlag (BasePeriodID,FisheryID,TimeStep,TerminalFlag) " & _
                   "VALUES(" & TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(3) & ")"

                TerminalFisheryFlag.ExecuteNonQuery()
            Next
            TerminalFisheryFlagTrans.Commit()
            TransDB.Close()

            'TimeStep
            If SpeciesName = "CHINOOK" Then
                CmdStr = "SELECT * FROM TimeStep"
                Dim TimeStepcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim TimeStepIDDA As New System.Data.OleDb.OleDbDataAdapter
                TimeStepIDDA.SelectCommand = TimeStepcm
                Dim TimeStepcb As New OleDb.OleDbCommandBuilder
                TimeStepcb = New OleDb.OleDbCommandBuilder(TimeStepIDDA)
                If TransferDataSet.Tables.Contains("TimeStep") Then
                    TransferDataSet.Tables("TimeStep").Clear()
                End If
                TimeStepIDDA.Fill(TransferDataSet, "TimeStep")
                Dim NumTimeStep As Integer
                NumTimeStep = TransferDataSet.Tables("TimeStep").Rows.Count

                Dim TimeStepTrans As OleDb.OleDbTransaction
                Dim TimeStep As New OleDbCommand
                TransDB.Open()
                TimeStepTrans = TransDB.BeginTransaction
                TimeStep.Connection = TransDB
                TimeStep.Transaction = TimeStepTrans
                NumRecs = TransferDataSet.Tables("TimeStep").Rows.Count
                For RecNum = 0 To NumRecs - 1
                    TimeStep.CommandText = "INSERT INTO TimeStep (Species,VersionNumber,TimeStepID,TimeStepName,TimeStepTitle) " & _
                       "VALUES(" & Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(0) & Chr(34) & "," & _
                        TransferDataSet.Tables("TimeStep").Rows(RecNum)(1) & "," & _
                        TransferDataSet.Tables("TimeStep").Rows(RecNum)(2) & "," & _
                        Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(3) & Chr(34) & "," & _
                        Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(4) & Chr(34) & ")"

                    TimeStep.ExecuteNonQuery()
                Next
                TimeStepTrans.Commit()
                TransDB.Close()
            End If
        Next
    End Sub
   Sub TransferModelRunTables()

      Dim CmdStr As String
      Dim TransID, RecNum, NumRecs, TransferBaseID As Integer

      'Loop through User Selected RunID Transfers
      For TransID = 0 To NumTransferID - 1

         'RunIDTransfer(TransID)

         '- Transfer RunID Record

         CmdStr = "SELECT * FROM RunID WHERE RunID = " & RunIDTransfer(TransID).ToString & ";"
         Dim RIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim RunIDDA As New System.Data.OleDb.OleDbDataAdapter
         RunIDDA.SelectCommand = RIDcm
         Dim RIDcb As New OleDb.OleDbCommandBuilder
         RIDcb = New OleDb.OleDbCommandBuilder(RunIDDA)
         If TransferDataSet.Tables.Contains("RunID") Then
            TransferDataSet.Tables("RunID").Clear()
         End If
         RunIDDA.Fill(TransferDataSet, "RunID")
         Dim NumRID As Integer
         NumRID = TransferDataSet.Tables("RunID").Rows.Count
         If NumRID <> 1 Then
            MsgBox("ERROR in RunID Table of Database ... Duplicate Record", MsgBoxStyle.OkOnly)
         End If
         SelectSpeciesName = TransferDataSet.Tables("RunID").Rows(0)(2)
         Dim RIDTrans As OleDb.OleDbTransaction
         Dim RID As New OleDbCommand
         TransDB.Open()
         RIDTrans = TransDB.BeginTransaction
         RID.Connection = TransDB
         RID.Transaction = RIDTrans
            RecNum = 0
         TransferBaseID = TransferDataSet.Tables("RunID").Rows(RecNum)(5)
            RID.CommandText = "INSERT INTO RunID (RunID,SpeciesName,RunName,RunTitle,BasePeriodID,RunComments,CreationDate,ModifyInputDate,RunTimeDate,RunYear) " & _
            "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(RecNum)(2) & Chr(34) & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(RecNum)(3) & Chr(34) & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(RecNum)(4) & Chr(34) & "," & _
            TransferDataSet.Tables("RunID").Rows(RecNum)(5).ToString & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(RecNum)(6) & Chr(34) & "," & _
            Chr(35) & TransferDataSet.Tables("RunID").Rows(RecNum)(7) & Chr(35) & "," & _
            Chr(35) & TransferDataSet.Tables("RunID").Rows(RecNum)(8) & Chr(35) & "," & _
            Chr(35) & TransferDataSet.Tables("RunID").Rows(RecNum)(9) & Chr(35) & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(RecNum)(10) & Chr(34) & ")"
         RID.ExecuteNonQuery()
         RIDTrans.Commit()
         TransDB.Close()

         '- Transfer BaseID Record

         CmdStr = "SELECT * FROM BaseID WHERE BasePeriodID = " & TransferBaseID.ToString & ";"
         Dim BIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
         Dim BaseIDDA As New System.Data.OleDb.OleDbDataAdapter
         BaseIDDA.SelectCommand = BIDcm
         Dim BIDcb As New OleDb.OleDbCommandBuilder
         BIDcb = New OleDb.OleDbCommandBuilder(BaseIDDA)
         If TransferDataSet.Tables.Contains("BaseID") Then
            TransferDataSet.Tables("BaseID").Clear()
         End If
         BaseIDDA.Fill(TransferDataSet, "BaseID")
         Dim NumBID As Integer
         NumBID = TransferDataSet.Tables("BaseID").Rows.Count
         If NumBID <> 1 Then
            MsgBox("ERROR in BaseID Table of Database ... Duplicate Record", MsgBoxStyle.OkOnly)
         End If
         Dim BIDTrans As OleDb.OleDbTransaction
         Dim BID As New OleDbCommand
         TransDB.Open()
         BIDTrans = TransDB.BeginTransaction
         BID.Connection = TransDB
         BID.Transaction = BIDTrans
         RecNum = 0
         BID.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
            "VALUES(" & TransferDataSet.Tables("BaseID").Rows(RecNum)(1) & "," & _
            Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(2) & Chr(34) & "," & _
            Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(3) & Chr(34) & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(4).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(5).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(6).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(7).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(8).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(9).ToString & "," & _
            Chr(35) & TransferDataSet.Tables("BaseID").Rows(RecNum)(10) & Chr(35) & "," & _
            Chr(34) & TransferDataSet.Tables("BaseID").Rows(RecNum)(11) & Chr(34) & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(12).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(13).ToString & "," & _
            TransferDataSet.Tables("BaseID").Rows(RecNum)(14).ToString & ")"
         BID.ExecuteNonQuery()
         BIDTrans.Commit()
         TransDB.Close()

         '- Transfer Backwards FRAM Table

         CmdStr = "SELECT * FROM BackwardsFRAM WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY StockID;"
         Dim BFcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim BFDA As New System.Data.OleDb.OleDbDataAdapter
            Dim i As Integer
         BFDA.SelectCommand = BFcm
         Dim BFcb As New OleDb.OleDbCommandBuilder
         BFcb = New OleDb.OleDbCommandBuilder(BFDA)
         If TransferDataSet.Tables.Contains("BackwardsFRAM") Then
            TransferDataSet.Tables("BackwardsFRAM").Clear()
         End If
         BFDA.Fill(TransferDataSet, "BackwardsFRAM")
         NumRecs = TransferDataSet.Tables("BackwardsFRAM").Rows.Count
         If NumRecs = 0 Then
            GoTo SkipBF
         End If
         Dim BFTrans As OleDb.OleDbTransaction
         Dim BFC As New OleDbCommand
         TransDB.Open()
         BFTrans = TransDB.BeginTransaction
         BFC.Connection = TransDB
            BFC.Transaction = BFTrans

            i = FramDataSet.Tables("BackwardsFRAM").Columns.IndexOf("Comment")
            For RecNum = 0 To NumRecs - 1
                If i <> -1 Then
                    BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag,Comment) " & _
                    "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(1).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(2).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(3).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(4).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(5).ToString & "," & _
                    Chr(34) & TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(6) & Chr(34) & ")"
                    Try
                        BFC.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox("Please select TransferFile version 4 or higher")
                        GoTo ExitTransfer
                    End Try
                Else
                    BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
                    "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(1).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(2).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(3).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(4).ToString & "," & _
                    TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(5).ToString & ")"
                    BFC.ExecuteNonQuery()
                End If
            Next
            BFTrans.Commit()
            TransDB.Close()
SkipBF:

            '- Transfer FisheryScalers Table

            CmdStr = "SELECT * FROM FisheryScalers WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY FisheryID, TimeStep;"
            Dim FScm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim FSDA As New System.Data.OleDb.OleDbDataAdapter
            FSDA.SelectCommand = FScm
            Dim FScb As New OleDb.OleDbCommandBuilder
            FScb = New OleDb.OleDbCommandBuilder(FSDA)
            If TransferDataSet.Tables.Contains("FisheryScalers") Then
                TransferDataSet.Tables("FisheryScalers").Clear()
            End If
            FSDA.Fill(TransferDataSet, "FisheryScalers")
            NumRecs = TransferDataSet.Tables("FisheryScalers").Rows.Count
            '- First Check if this Transfer Database is from "Old" format
            Dim column As DataColumn
            For Each column In TransferDataSet.Tables("FisheryScalers").Columns
                If (column.ColumnName) = "MSFFisheryScaleFactor" Then GoTo FoundNewColumn
            Next
            MsgBox("ERROR - You have an Old Format NewModelRunTransfer.Mdb Database" & vbCrLf & _
                     "You need to get the New Format Database to do Model Run Transfers", MsgBoxStyle.OkOnly)
            Exit Sub
FoundNewColumn:
            If NumRecs = 0 Then
                MsgBox("Error in FisheryScalers Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipFS
            End If
            i = FramDataSet.Tables("FisheryScalers").Columns.IndexOf("Comment")
            Dim FSTrans As OleDb.OleDbTransaction
            Dim FSC As New OleDbCommand
            TransDB.Open()
            FSTrans = TransDB.BeginTransaction
            FSC.Connection = TransDB
            FSC.Transaction = FSTrans

            'MessageBox2: MsgBox("Please select TransferFile version 4 or higher")

            For RecNum = 0 To NumRecs - 1
                '- MarkSelectiveFlag currently not used ... placeholder after Quota
                If i <> -1 Then
                    FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate,Comment) " & _
                     "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(7).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(12).ToString & "," & _
                     Chr(34) & TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(13) & Chr(34) & ")"
                    Try
                        FSC.ExecuteNonQuery()
                    Catch ex As Exception
                        MsgBox("Please select TransferFile version 4 or higher")
                        GoTo ExitTransfer
                    End Try
                Else
                    FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                     "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(7).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & "," & _
                     TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(12).ToString & ")"
                    FSC.ExecuteNonQuery()
                End If
            Next
            FSTrans.Commit()
            TransDB.Close()
SkipFS:

            '- Transfer NonRetention Table

            CmdStr = "SELECT * FROM NonRetention WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY FisheryID, TimeStep;"
            Dim NRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim NRDA As New System.Data.OleDb.OleDbDataAdapter
            NRDA.SelectCommand = NRcm
            Dim NRcb As New OleDb.OleDbCommandBuilder
            NRcb = New OleDb.OleDbCommandBuilder(NRDA)
            If TransferDataSet.Tables.Contains("NonRetention") Then
                TransferDataSet.Tables("NonRetention").Clear()
            End If
            NRDA.Fill(TransferDataSet, "NonRetention")
            NumRecs = TransferDataSet.Tables("NonRetention").Rows.Count
            If NumRecs = 0 Then
                MsgBox("Error in NonRetention Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipNR
            End If
            Dim NRTrans As OleDb.OleDbTransaction
            Dim NRC As New OleDbCommand
            TransDB.Open()
            NRTrans = TransDB.BeginTransaction
            NRC.Connection = TransDB
            NRC.Transaction = NRTrans
            For RecNum = 0 To NumRecs - 1
                NRC.CommandText = "INSERT INTO NonRetention (RunID,FisheryID,TimeStep,NonRetentionFlag,CNRInput1,CNRInput2,CNRInput3,CNRInput4) " & _
                   "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(2).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(3).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(4).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(5).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(6).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(7).ToString & "," & _
                   TransferDataSet.Tables("NonRetention").Rows(RecNum)(8).ToString & ")"
                NRC.ExecuteNonQuery()
            Next
            NRTrans.Commit()
            TransDB.Close()
SkipNR:

            '- Transfer Stock/Fishery Rate Scalers

            CmdStr = "SELECT * FROM StockFisheryRateScaler WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY StockID, FisheryID, TimeStep"
            Dim SFRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim SFDA As New System.Data.OleDb.OleDbDataAdapter
            SFDA.SelectCommand = SFRcm
            Dim SFRcb As New OleDb.OleDbCommandBuilder
            SFRcb = New OleDb.OleDbCommandBuilder(SFDA)
            If TransferDataSet.Tables.Contains("StockFisheryRateScaler") Then
                TransferDataSet.Tables("StockFisheryRateScaler").Clear()
            End If
            SFDA.Fill(TransferDataSet, "StockFisheryRateScaler")
            NumRecs = TransferDataSet.Tables("StockFisheryRateScaler").Rows.Count
            Dim SFRTrans As OleDb.OleDbTransaction
            Dim SFRC As New OleDbCommand
            TransDB.Open()
            SFRTrans = TransDB.BeginTransaction
            SFRC.Connection = TransDB
            SFRC.Transaction = SFRTrans
            For RecNum = 0 To NumRecs - 1
                SFRC.CommandText = "INSERT INTO StockFisheryRateScaler (RunID,StockID,FisheryID,TimeStep,StockFisheryRateScaler) " & _
                 "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                 TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(1).ToString & "," & _
                 TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(2).ToString & "," & _
                 TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(3).ToString & "," & _
                 TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(4).ToString & ")"
                SFRC.ExecuteNonQuery()
            Next
            SFRTrans.Commit()
            TransDB.Close()
            SFDA = Nothing
SkipSFR:

            '- Transfer PSCMaxER - Coho Only

            If SelectSpeciesName = "CHINOOK" Then GoTo SkipPSCER
            CmdStr = "SELECT * FROM PSCMaxER WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY PSCStockID"
            Dim PSCcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim PSCDA As New System.Data.OleDb.OleDbDataAdapter
            PSCDA.SelectCommand = PSCcm
            Dim PSCcb As New OleDb.OleDbCommandBuilder
            PSCcb = New OleDb.OleDbCommandBuilder(PSCDA)
            If TransferDataSet.Tables.Contains("PSCMaxER") Then
                TransferDataSet.Tables("PSCMaxER").Clear()
            End If
            PSCDA.Fill(TransferDataSet, "PSCMaxER")
            NumRecs = TransferDataSet.Tables("PSCMaxER").Rows.Count
            If NumRecs = 0 Then
                MsgBox("Error in PSCMaxER Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipPSCER
            End If
            Dim PSCTrans As OleDb.OleDbTransaction
            Dim PSCC As New OleDbCommand
            TransDB.Open()
            PSCTrans = TransDB.BeginTransaction
            PSCC.Connection = TransDB
            PSCC.Transaction = PSCTrans
            For RecNum = 0 To NumRecs - 1
                PSCC.CommandText = "INSERT INTO PSCMaxER (RunID,PSCStockID,PSCMaxER) " & _
                 "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                 TransferDataSet.Tables("PSCMaxER").Rows(RecNum)(1).ToString & "," & _
                 TransferDataSet.Tables("PSCMaxER").Rows(RecNum)(2).ToString & ")"
                PSCC.ExecuteNonQuery()
            Next
            PSCTrans.Commit()
            TransDB.Close()
            PSCDA = Nothing
SkipPSCER:

            '- Size Limits - Chinook Only

            If SelectSpeciesName = "COHO" Then GoTo SkipSL
            CmdStr = "SELECT * FROM SizeLimits WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY FisheryID, TimeStep"
            Dim SLcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim SLDA As New System.Data.OleDb.OleDbDataAdapter
            SLDA.SelectCommand = SLcm
            Dim SLcb As New OleDb.OleDbCommandBuilder
            SLcb = New OleDb.OleDbCommandBuilder(SLDA)
            If TransferDataSet.Tables.Contains("SizeLimits") Then
                TransferDataSet.Tables("SizeLimits").Clear()
            End If
            SLDA.Fill(TransferDataSet, "SizeLimits")
            NumRecs = TransferDataSet.Tables("SizeLimits").Rows.Count
            If NumRecs = 0 Then
                MsgBox("Error in SizeLimits Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipSL
            End If
            Dim SLTrans As OleDb.OleDbTransaction
            Dim SLC As New OleDbCommand
            TransDB.Open()
            SLTrans = TransDB.BeginTransaction
            SLC.Connection = TransDB
            SLC.Transaction = SLTrans
            For RecNum = 0 To NumRecs - 1
                SLC.CommandText = "INSERT INTO SizeLimits (RunID,FisheryID,TimeStep,MinimumSize,MaximumSize) " & _
                 "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                 TransferDataSet.Tables("SizeLimits").Rows(RecNum)(2).ToString & "," & _
                 TransferDataSet.Tables("SizeLimits").Rows(RecNum)(3).ToString & "," & _
                 TransferDataSet.Tables("SizeLimits").Rows(RecNum)(4).ToString & "," & _
                 TransferDataSet.Tables("SizeLimits").Rows(RecNum)(5).ToString & ")"
                SLC.ExecuteNonQuery()
            Next
            SLTrans.Commit()
            TransDB.Close()
            SLDA = Nothing
SkipSL:

            '- Transfer Stock Recruits

            CmdStr = "SELECT * FROM StockRecruit WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY StockID, Age"
            Dim SRcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim SRDA As New System.Data.OleDb.OleDbDataAdapter
            SRDA.SelectCommand = SRcm
            Dim SRcb As New OleDb.OleDbCommandBuilder
            SRcb = New OleDb.OleDbCommandBuilder(SRDA)
            If TransferDataSet.Tables.Contains("StockRecruit") Then
                TransferDataSet.Tables("StockRecruit").Clear()
            End If
            SRDA.Fill(TransferDataSet, "StockRecruit")
            NumRecs = TransferDataSet.Tables("StockRecruit").Rows.Count
            If NumRecs = 0 Then
                MsgBox("Error in StockRecruit Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipSR
            End If
            Dim SRTrans As OleDb.OleDbTransaction
            Dim SRC As New OleDbCommand
            TransDB.Open()
            SRTrans = TransDB.BeginTransaction
            SRC.Connection = TransDB
            SRC.Transaction = SRTrans
            For RecNum = 0 To NumRecs - 1
                SRC.CommandText = "INSERT INTO StockRecruit (RunID,StockID,Age,RecruitScaleFactor,RecruitCohortSize) " & _
                 "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                 TransferDataSet.Tables("StockRecruit").Rows(RecNum)(2).ToString & "," & _
                 TransferDataSet.Tables("StockRecruit").Rows(RecNum)(3).ToString & "," & _
                 TransferDataSet.Tables("StockRecruit").Rows(RecNum)(4).ToString & "," & _
                 TransferDataSet.Tables("StockRecruit").Rows(RecNum)(5).ToString & ")"
                SRC.ExecuteNonQuery()
            Next
            SRTrans.Commit()
            TransDB.Close()
            SRDA = Nothing
SkipSR:

            '- Transfer Cohort Run Sizes

            CmdStr = "SELECT * FROM Cohort WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY StockID, Age, TimeStep"
            Dim COHcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim COHDA As New System.Data.OleDb.OleDbDataAdapter
            COHDA.SelectCommand = COHcm
            Dim COHcb As New OleDb.OleDbCommandBuilder
            COHcb = New OleDb.OleDbCommandBuilder(COHDA)
            If TransferDataSet.Tables.Contains("Cohort") Then
                TransferDataSet.Tables("Cohort").Clear()
            End If
            COHDA.Fill(TransferDataSet, "Cohort")
            NumRecs = TransferDataSet.Tables("Cohort").Rows.Count
            Dim COHTrans As OleDb.OleDbTransaction
            Dim COHC As New OleDbCommand
            TransDB.Open()
            COHTrans = TransDB.BeginTransaction
            COHC.Connection = TransDB
            COHC.Transaction = COHTrans
            For RecNum = 0 To NumRecs - 1
                COHC.CommandText = "INSERT INTO Cohort (RunID,StockID,Age,TimeStep,Cohort,MatureCohort,StartCohort,WorkingCohort,MidCohort) " & _
                "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(2).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(3).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(4).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(5).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(6).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(7).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(8).ToString & "," & _
                TransferDataSet.Tables("Cohort").Rows(RecNum)(9).ToString & ")"
                COHC.ExecuteNonQuery()
            Next
            COHTrans.Commit()
            TransDB.Close()
            COHDA = Nothing

            '- Transfer Escapement

            CmdStr = "SELECT * FROM Escapement WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY StockID, Age, TimeStep"
            Dim ESCcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim ESCDA As New System.Data.OleDb.OleDbDataAdapter
            ESCDA.SelectCommand = ESCcm
            Dim ESCcb As New OleDb.OleDbCommandBuilder
            ESCcb = New OleDb.OleDbCommandBuilder(ESCDA)
            If TransferDataSet.Tables.Contains("Escapement") Then
                TransferDataSet.Tables("Escapement").Clear()
            End If
            ESCDA.Fill(TransferDataSet, "Escapement")
            NumRecs = TransferDataSet.Tables("Escapement").Rows.Count
            Dim ESCTrans As OleDb.OleDbTransaction
            Dim ESCC As New OleDbCommand
            TransDB.Open()
            ESCTrans = TransDB.BeginTransaction
            ESCC.Connection = TransDB
            ESCC.Transaction = ESCTrans
            For RecNum = 0 To NumRecs - 1
                ESCC.CommandText = "INSERT INTO Escapement (RunID,StockID,Age,TimeStep,Escapement) " & _
                "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                TransferDataSet.Tables("Escapement").Rows(RecNum)(2).ToString & "," & _
                TransferDataSet.Tables("Escapement").Rows(RecNum)(3).ToString & "," & _
                TransferDataSet.Tables("Escapement").Rows(RecNum)(4).ToString & "," & _
                TransferDataSet.Tables("Escapement").Rows(RecNum)(5).ToString & ")"
                ESCC.ExecuteNonQuery()
            Next
            ESCTrans.Commit()
            TransDB.Close()
            ESCDA = Nothing

            '- Transfer FisheryMortality

            CmdStr = "SELECT * FROM FisheryMortality WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY FisheryID, TimeStep"
            Dim FMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim FMDA As New System.Data.OleDb.OleDbDataAdapter
            FMDA.SelectCommand = FMcm
            Dim FMcb As New OleDb.OleDbCommandBuilder
            FMcb = New OleDb.OleDbCommandBuilder(FMDA)
            If TransferDataSet.Tables.Contains("FisheryMortality") Then
                TransferDataSet.Tables("FisheryMortality").Clear()
            End If
            FMDA.Fill(TransferDataSet, "FisheryMortality")
            NumRecs = TransferDataSet.Tables("FisheryMortality").Rows.Count
            Dim FMTrans As OleDb.OleDbTransaction
            Dim FMC As New OleDbCommand
            TransDB.Open()
            FMTrans = TransDB.BeginTransaction
            FMC.Connection = TransDB
            FMC.Transaction = FMTrans
            For RecNum = 0 To NumRecs - 1
                FMC.CommandText = "INSERT INTO FisheryMortality (RunID,FisheryID,TimeStep,TotalLandedCatch,TotalUnMarkedCatch,TotalNonRetention,TotalShakers,TotalDropOff,TotalEncounters) " & _
                   "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(1).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(2).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(3).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(4).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(5).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(6).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(7).ToString & "," & _
                   TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(8).ToString & ")"
                FMC.ExecuteNonQuery()
            Next
            FMTrans.Commit()
            TransDB.Close()
            FMDA = Nothing

            '- Transfer All Mortality Records

            CmdStr = "SELECT * FROM Mortality WHERE RunID = " & RunIDTransfer(TransID).ToString & " ORDER BY FisheryID, TimeStep"
            Dim MRTcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim MRTDA As New System.Data.OleDb.OleDbDataAdapter
            MRTDA.SelectCommand = MRTcm
            Dim MRTcb As New OleDb.OleDbCommandBuilder
            MRTcb = New OleDb.OleDbCommandBuilder(MRTDA)
            If TransferDataSet.Tables.Contains("Mortality") Then
                TransferDataSet.Tables("Mortality").Clear()
            End If
            MRTDA.Fill(TransferDataSet, "Mortality")
            NumRecs = TransferDataSet.Tables("Mortality").Rows.Count
            Dim MRTTrans As OleDb.OleDbTransaction
            Dim MRTC As New OleDbCommand
            TransDB.Open()
            MRTTrans = TransDB.BeginTransaction
            MRTC.Connection = TransDB
            MRTC.Transaction = MRTTrans
            For RecNum = 0 To NumRecs - 1
                MRTC.CommandText = "INSERT INTO Mortality (RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                   "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(2).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(3).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(4).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(5).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(6).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(7).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(8).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(9).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(10).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(11).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(12).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(13).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(14).ToString & "," & _
                   TransferDataSet.Tables("Mortality").Rows(RecNum)(15).ToString & ")"
                MRTC.ExecuteNonQuery()
            Next
            MRTTrans.Commit()
            TransDB.Close()
            MRTDA = Nothing

            '==============================================================================================
            '- (Pete 12/13) Inject transfer database with Target Sublegal:Legal Ratio (SLRatio) 
            '- and run-specific sublegal encounter rate adjustment (RunEncounterRateAdjustment) content

            '- Transfer Sublegal Ratios
            CmdStr = "SELECT * FROM SLRatio WHERE RunID = " & RunIDTransfer(TransID).ToString
            Dim SLRatcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim SLRatDA As New System.Data.OleDb.OleDbDataAdapter
            SLRatDA.SelectCommand = SLRatcm
            Dim SLRatcb As New OleDb.OleDbCommandBuilder
            SLRatcb = New OleDb.OleDbCommandBuilder(SLRatDA)
            If TransferDataSet.Tables.Contains("SLRatio") Then
                TransferDataSet.Tables("SLRatio").Clear()
            End If
            SLRatDA.Fill(TransferDataSet, "SLRatio")
            NumRecs = TransferDataSet.Tables("SLRatio").Rows.Count
            If NumRecs = 0 Then
                'MsgBox("Error in StockRecruit Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                GoTo SkipSLRat
            End If
            Dim SLRatTrans As OleDb.OleDbTransaction
            Dim SLRatC As New OleDbCommand
            TransDB.Open()
            SLRatTrans = TransDB.BeginTransaction
            SLRatC.Connection = TransDB
            SLRatC.Transaction = SLRatTrans
            For RecNum = 0 To NumRecs - 1
                SLRatC.CommandText = "INSERT INTO SLRatio (RunID,FisheryID,Age,TimeStep,TargetRatio,RunEncounterRateAdjustment, UpdateWhen, UpdateBy) " & _
                    "VALUES(" & RunIDTransfer(TransID).ToString & "," & _
                    TransferDataSet.Tables("SLRatio").Rows(RecNum)(1).ToString & "," & _
                    TransferDataSet.Tables("SLRatio").Rows(RecNum)(2).ToString & "," & _
                    TransferDataSet.Tables("SLRatio").Rows(RecNum)(3).ToString & "," & _
                    TransferDataSet.Tables("SLRatio").Rows(RecNum)(4).ToString & "," & _
                    TransferDataSet.Tables("SLRatio").Rows(RecNum)(5).ToString & "," & _
                    "'" & TransferDataSet.Tables("SLRatio").Rows(RecNum)(6).ToString & "'" & "," & _
                    "'" & TransferDataSet.Tables("SLRatio").Rows(RecNum)(7).ToString & "'" & ")"
                SLRatC.ExecuteNonQuery()
            Next
            SLRatTrans.Commit()
            TransDB.Close()
            SLRatDA = Nothing
SkipSLRat:

            '==============================================================================================



        Next
ExitTransfer:
        Exit Sub

    End Sub

    Sub GetTransferBasePeriodTables()
        '- This SubRoutine is the opposite of the TransferBaseRunTables and reads in a new base period


        Dim CmdStr As String
        Dim RecNum, NumRecs, TransID, NewRunID, OldRunID As Integer


        CmdStr = "SELECT * FROM BaseID;"
        Dim BIDcm As New OleDb.OleDbCommand(CmdStr, TransBP)
        Dim BaseIDDA As New System.Data.OleDb.OleDbDataAdapter
        BaseIDDA.SelectCommand = BIDcm
        Dim BIDcb As New OleDb.OleDbCommandBuilder
        BIDcb = New OleDb.OleDbCommandBuilder(BaseIDDA)
        If TransferDataSet.Tables.Contains("BaseID") Then
            TransferDataSet.Tables("BaseID").Clear()
        End If
        BaseIDDA.Fill(TransferDataSet, "BaseID")
        BaseIDDA = Nothing



        Dim TransBaseID As Integer

        'Dim drd2 As OleDb.OleDbDataReader
        'Dim cmd2 As New OleDb.OleDbCommand()
        'cmd2.Connection = FramDB
        'FramDB.Open()

        Dim row As Integer
        Dim NewBasePeriodID As Integer
        'for each baser period in transfer database test if base period number already exists in database
        For i = 0 To TransferDataSet.Tables("BaseID").Rows.Count - 1

            Dim drd2 As OleDb.OleDbDataReader
            Dim cmd2 As New OleDb.OleDbCommand()
            cmd2.Connection = FramDB
            FramDB.Open()






            TransBaseID = TransferDataSet.Tables("BaseID").Rows(i)("BasePeriodID")
            cmd2.CommandText = "SELECT * FROM BaseID WHERE BasePeriodID = " & TransBaseID & " ORDER BY BasePeriodID;"
            drd2 = cmd2.ExecuteReader

            If drd2.Read() <> False Then
                drd2.Close()
                ' the transfer base ID already exists in the database
                cmd2.CommandText = "Select BasePeriodID FROM BaseID ORDER BY BasePeriodID;"
                drd2 = cmd2.ExecuteReader

                Do While drd2.Read()
                    'Console.WriteLine(drd2.GetInt32(0))
                    NewBasePeriodID = drd2.GetValue(0) + 1 ' equals higher BPID in database plus 1

                Loop
                MsgBox("TransferBasePeriodID = '" & TransBaseID & "' already exists in FramVS Database. The new Transfer Database will receive Base Period ID " & NewBasePeriodID & ".")
            Else
                NewBasePeriodID = TransBaseID
            End If



            'Add BaseID record from TransferDatabase to recipient database, replace BASEID in table 

            Dim BaseIDTrans As OleDb.OleDbTransaction
            Dim BaseID As New OleDbCommand

            BaseIDTrans = FramDB.BeginTransaction
            BaseID.Connection = FramDB
            BaseID.Transaction = BaseIDTrans
            RecNum = 0

            BaseID.CommandText = "INSERT INTO BaseID (BasePeriodID,BasePeriodName,SpeciesName,NumStocks,NumFisheries,NumTimeSteps,NumAges,MinAge,MaxAge,DateCreated,BaseComments,StockVersion,FisheryVersion,TimeStepVersion) " & _
               "VALUES(" & NewBasePeriodID & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(i)(2) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(i)(3) & Chr(34) & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(4).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(5).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(6).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(7).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(8).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(9).ToString & "," & _
               Chr(35) & TransferDataSet.Tables("BaseID").Rows(i)(10) & Chr(35) & "," & _
               Chr(34) & TransferDataSet.Tables("BaseID").Rows(i)(11) & Chr(34) & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(12).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(13).ToString & "," & _
               TransferDataSet.Tables("BaseID").Rows(i)(14).ToString & ")"
            BaseID.ExecuteNonQuery()
            BaseIDTrans.Commit()
            FramDB.Close()


            'populate remaining transfer datasets



            'BaseCohort()
            FramDB.Open()
            CmdStr = "SELECT * FROM BaseCohort WHERE BasePeriodID = " & TransBaseID & ";"
            Dim BaseCohortcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim BaseCohortIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseCohortIDDA.SelectCommand = BaseCohortcm
            Dim BaseCohortcb As New OleDb.OleDbCommandBuilder
            BaseCohortcb = New OleDb.OleDbCommandBuilder(BaseCohortIDDA)
            If TransferDataSet.Tables.Contains("BaseCohort") Then
                TransferDataSet.Tables("BaseCohort").Clear()
            End If
            BaseCohortIDDA.Fill(TransferDataSet, "BaseCohort")
            BaseCohortIDDA = Nothing

            Dim NumBaseCohort As Integer
            NumBaseCohort = TransferDataSet.Tables("BaseCohort").Rows.Count
            Dim BaseCohortTrans As OleDb.OleDbTransaction
            Dim BaseCohort As New OleDbCommand

            BaseCohortTrans = FramDB.BeginTransaction
            BaseCohort.Connection = FramDB
            BaseCohort.Transaction = BaseCohortTrans
            NumRecs = TransferDataSet.Tables("BaseCohort").Rows.Count
            For RecNum = 0 To NumRecs - 1

                BaseCohort.CommandText = "INSERT INTO BaseCohort (BasePeriodID,StockID,Age,BaseCohortSize) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("BaseCohort").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("BaseCohort").Rows(RecNum)(2) & "," & _
                   TransferDataSet.Tables("BaseCohort").Rows(RecNum)(3) & ")"

                BaseCohort.ExecuteNonQuery()
            Next
            BaseCohortTrans.Commit()
            FramDB.Close()

            'BaseExploitationRate
            FramDB.Open()
            CmdStr = "SELECT * FROM BaseExploitationRate WHERE BasePeriodID = " & TransBaseID & ";"
            Dim BaseExploitationRatecm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim BaseExploitationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseExploitationRateIDDA.SelectCommand = BaseExploitationRatecm
            Dim BaseExploitationRatecb As New OleDb.OleDbCommandBuilder
            BaseExploitationRatecb = New OleDb.OleDbCommandBuilder(BaseExploitationRateIDDA)
            If TransferDataSet.Tables.Contains("BaseExploitationRate") Then
                TransferDataSet.Tables("BaseExploitationRate").Clear()
            End If
            BaseExploitationRateIDDA.Fill(TransferDataSet, "BaseExploitationRate")
            BaseExploitationRateIDDA = Nothing

            Dim NumBaseExploitationRate As Integer
            NumBaseExploitationRate = TransferDataSet.Tables("BaseExploitationRate").Rows.Count
            Dim BaseExploitationRateTrans As OleDb.OleDbTransaction
            Dim BaseExploitationRate As New OleDbCommand

            BaseExploitationRateTrans = FramDB.BeginTransaction
            BaseExploitationRate.Connection = FramDB
            BaseExploitationRate.Transaction = BaseExploitationRateTrans
            NumRecs = TransferDataSet.Tables("BaseExploitationRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                BaseExploitationRate.CommandText = "INSERT INTO BaseExploitationRate (BasePeriodID,StockID,Age,FisheryID,TimeStep,ExploitationRate,SublegalEncounterRate) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(3) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(4) & "," & _
                    TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(5) & "," & _
                   TransferDataSet.Tables("BaseExploitationRate").Rows(RecNum)(6) & ")"

                BaseExploitationRate.ExecuteNonQuery()
            Next
            BaseExploitationRateTrans.Commit()
            FramDB.Close()

            'EncounterRateAdjustment()
            FramDB.Open()
            CmdStr = "SELECT * FROM EncounterRateAdjustment WHERE BasePeriodID = " & TransBaseID & ";"
            Dim EncounterRateAdjustmentcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim EncounterRateAdjustmentIDDA As New System.Data.OleDb.OleDbDataAdapter
            EncounterRateAdjustmentIDDA.SelectCommand = EncounterRateAdjustmentcm
            Dim EncounterRateAdjustmentcb As New OleDb.OleDbCommandBuilder
            EncounterRateAdjustmentcb = New OleDb.OleDbCommandBuilder(EncounterRateAdjustmentIDDA)
            If TransferDataSet.Tables.Contains("EncounterRateAdjustment") Then
                TransferDataSet.Tables("EncounterRateAdjustment").Clear()
            End If
            EncounterRateAdjustmentIDDA.Fill(TransferDataSet, "EncounterRateAdjustment")
            EncounterRateAdjustmentIDDA = Nothing

            Dim NumEncounterRateAdjustment As Integer
            NumEncounterRateAdjustment = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count
            Dim EncounterRateAdjustmentTrans As OleDb.OleDbTransaction
            Dim EncounterRateAdjustment As New OleDbCommand

            EncounterRateAdjustmentTrans = FramDB.BeginTransaction
            EncounterRateAdjustment.Connection = FramDB
            EncounterRateAdjustment.Transaction = EncounterRateAdjustmentTrans
            NumRecs = TransferDataSet.Tables("EncounterRateAdjustment").Rows.Count
            For RecNum = 0 To NumRecs - 1
                EncounterRateAdjustment.CommandText = "INSERT INTO EncounterRateAdjustment (BasePeriodID,Age,FisheryID,TimeStep,EncounterRateAdjustment) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(3) & "," & _
                   TransferDataSet.Tables("EncounterRateAdjustment").Rows(RecNum)(4) & ")"

                EncounterRateAdjustment.ExecuteNonQuery()
            Next
            EncounterRateAdjustmentTrans.Commit()
            FramDB.Close()

            'FisheryModelStockProportion      
            FramDB.Open()
            CmdStr = "SELECT * FROM FisheryModelStockProportion WHERE BasePeriodID = " & TransBaseID & ";"
            Dim FisheryModelStockProportioncm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim FisheryModelStockProportionIDDA As New System.Data.OleDb.OleDbDataAdapter
            FisheryModelStockProportionIDDA.SelectCommand = FisheryModelStockProportioncm
            Dim FisheryModelStockProportioncb As New OleDb.OleDbCommandBuilder
            FisheryModelStockProportioncb = New OleDb.OleDbCommandBuilder(FisheryModelStockProportionIDDA)
            If TransferDataSet.Tables.Contains("FisheryModelStockProportion") Then
                TransferDataSet.Tables("FisheryModelStockProportion").Clear()
            End If
            FisheryModelStockProportionIDDA.Fill(TransferDataSet, "FisheryModelStockProportion")
            FisheryModelStockProportionIDDA = Nothing

            Dim NumFisheryModelStockProportion As Integer
            NumFisheryModelStockProportion = TransferDataSet.Tables("FisheryModelStockProportion").Rows.Count
            Dim FisheryModelStockProportionTrans As OleDb.OleDbTransaction
            Dim FisheryModelStockProportion As New OleDbCommand

            FisheryModelStockProportionTrans = FramDB.BeginTransaction
            FisheryModelStockProportion.Connection = FramDB
            FisheryModelStockProportion.Transaction = FisheryModelStockProportionTrans
            NumRecs = TransferDataSet.Tables("FisheryModelStockProportion").Rows.Count
            For RecNum = 0 To NumRecs - 1
                FisheryModelStockProportion.CommandText = "INSERT INTO FisheryModelStockProportion (BasePeriodID,FisheryID,ModelStockProportion) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(1) & "," & _
                   TransferDataSet.Tables("FisheryModelStockProportion").Rows(RecNum)(2) & ")"
                FisheryModelStockProportion.ExecuteNonQuery()
            Next
            FisheryModelStockProportionTrans.Commit()
            FramDB.Close()

            'AEQ      
            FramDB.Open()
            CmdStr = "SELECT * FROM AEQ WHERE BasePeriodID = " & TransBaseID & ";"
            Dim AEQcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim AEQIDDA As New System.Data.OleDb.OleDbDataAdapter
            AEQIDDA.SelectCommand = AEQcm
            Dim AEQcb As New OleDb.OleDbCommandBuilder
            AEQcb = New OleDb.OleDbCommandBuilder(AEQIDDA)
            If TransferDataSet.Tables.Contains("AEQ") Then
                TransferDataSet.Tables("AEQ").Clear()
            End If
            AEQIDDA.Fill(TransferDataSet, "AEQ")
            AEQIDDA = Nothing

            Dim NumAEQ As Integer
            NumAEQ = TransferDataSet.Tables("AEQ").Rows.Count
            Dim AEQTrans As OleDb.OleDbTransaction
            Dim AEQ As New OleDbCommand

            AEQTrans = FramDB.BeginTransaction
            AEQ.Connection = FramDB
            AEQ.Transaction = AEQTrans
            NumRecs = TransferDataSet.Tables("AEQ").Rows.Count
            For RecNum = 0 To NumRecs - 1
                AEQ.CommandText = "INSERT INTO AEQ (BasePeriodID,StockID,Age,TimeStep,AEQ) " & _
                  "VALUES(" & NewBasePeriodID & "," & _
                   TransferDataSet.Tables("AEQ").Rows(RecNum)(1) & "," & _
                   TransferDataSet.Tables("AEQ").Rows(RecNum)(2) & "," & _
                  TransferDataSet.Tables("AEQ").Rows(RecNum)(3) & "," & _
                  TransferDataSet.Tables("AEQ").Rows(RecNum)(4) & ")"
                AEQ.ExecuteNonQuery()
            Next
            AEQTrans.Commit()
            FramDB.Close()

            'Growth      
            FramDB.Open()
            CmdStr = "SELECT * FROM Growth WHERE BasePeriodID = " & TransBaseID & ";"
            Dim Growthcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim GrowthIDDA As New System.Data.OleDb.OleDbDataAdapter
            GrowthIDDA.SelectCommand = Growthcm
            Dim Growthcb As New OleDb.OleDbCommandBuilder
            Growthcb = New OleDb.OleDbCommandBuilder(GrowthIDDA)
            If TransferDataSet.Tables.Contains("Growth") Then
                TransferDataSet.Tables("Growth").Clear()
            End If
            GrowthIDDA.Fill(TransferDataSet, "Growth")
            GrowthIDDA = Nothing

            Dim NumGrowth As Integer
            NumGrowth = TransferDataSet.Tables("Growth").Rows.Count
            Dim GrowthTrans As OleDb.OleDbTransaction
            Dim Growth As New OleDbCommand

            GrowthTrans = FramDB.BeginTransaction
            Growth.Connection = FramDB
            Growth.Transaction = GrowthTrans
            NumRecs = TransferDataSet.Tables("Growth").Rows.Count
            For RecNum = 0 To NumRecs - 1
                Growth.CommandText = "INSERT INTO Growth (BasePeriodID,StockID,LImmature,KImmature,TImmature,CV2Immature,CV3Immature,CV4Immature,CV5Immature,LMature,KMature,TMature,CV2Mature,CV3Mature,CV4Mature,CV5Mature) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("Growth").Rows(RecNum)(1) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(2) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(3) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(4) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(5) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(6) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(7) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(8) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(9) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(10) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(11) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(12) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(13) & "," & _
                 TransferDataSet.Tables("Growth").Rows(RecNum)(14) & "," & _
                   TransferDataSet.Tables("Growth").Rows(RecNum)(15) & ")"
                Growth.ExecuteNonQuery()
            Next
            GrowthTrans.Commit()
            FramDB.Close()

            'IncidentalRate()
            FramDB.Open()
            CmdStr = "SELECT * FROM IncidentalRate WHERE BasePeriodID = " & TransBaseID & ";"
            Dim IncidentalRatecm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim IncidentalRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            IncidentalRateIDDA.SelectCommand = IncidentalRatecm
            Dim IncidentalRatecb As New OleDb.OleDbCommandBuilder
            IncidentalRatecb = New OleDb.OleDbCommandBuilder(IncidentalRateIDDA)
            If TransferDataSet.Tables.Contains("IncidentalRate") Then
                TransferDataSet.Tables("IncidentalRate").Clear()
            End If
            IncidentalRateIDDA.Fill(TransferDataSet, "IncidentalRate")
            IncidentalRateIDDA = Nothing

            Dim NumIncidentalRate As Integer
            NumIncidentalRate = TransferDataSet.Tables("IncidentalRate").Rows.Count
            Dim IncidentalRateTrans As OleDb.OleDbTransaction
            Dim IncidentalRate As New OleDbCommand

            IncidentalRateTrans = FramDB.BeginTransaction
            IncidentalRate.Connection = FramDB
            IncidentalRate.Transaction = IncidentalRateTrans
            NumRecs = TransferDataSet.Tables("IncidentalRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                IncidentalRate.CommandText = "INSERT INTO IncidentalRate (BasePeriodID,FisheryID,TimeStep,IncidentalRate) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("IncidentalRate").Rows(RecNum)(3) & ")"

                IncidentalRate.ExecuteNonQuery()
            Next
            IncidentalRateTrans.Commit()
            FramDB.Close()

            'MaturationRate      
            FramDB.Open()
            CmdStr = "SELECT * FROM MaturationRate WHERE BasePeriodID = " & TransBaseID & ";"
            Dim MaturationRatecm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim MaturationRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            MaturationRateIDDA.SelectCommand = MaturationRatecm
            Dim MaturationRatecb As New OleDb.OleDbCommandBuilder
            MaturationRatecb = New OleDb.OleDbCommandBuilder(MaturationRateIDDA)
            If TransferDataSet.Tables.Contains("MaturationRate") Then
                TransferDataSet.Tables("MaturationRate").Clear()
            End If
            MaturationRateIDDA.Fill(TransferDataSet, "MaturationRate")
            MaturationRateIDDA = Nothing

            Dim NumMaturationRate As Integer
            NumMaturationRate = TransferDataSet.Tables("MaturationRate").Rows.Count
            Dim MaturationRateTrans As OleDb.OleDbTransaction
            Dim MaturationRate As New OleDbCommand

            MaturationRateTrans = FramDB.BeginTransaction
            MaturationRate.Connection = FramDB
            MaturationRate.Transaction = MaturationRateTrans
            NumRecs = TransferDataSet.Tables("MaturationRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                MaturationRate.CommandText = "INSERT INTO MaturationRate (BasePeriodID,StockID,Age,TimeStep,MaturationRate) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(3) & "," & _
                    TransferDataSet.Tables("MaturationRate").Rows(RecNum)(4) & ")"

                MaturationRate.ExecuteNonQuery()
            Next
            MaturationRateTrans.Commit()
            FramDB.Close()

            'NaturalMortality      
            FramDB.Open()
            CmdStr = "SELECT * FROM NaturalMortality WHERE BasePeriodID = " & TransBaseID & ";"
            Dim NaturalMortalitycm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim NaturalMortalityIDDA As New System.Data.OleDb.OleDbDataAdapter
            NaturalMortalityIDDA.SelectCommand = NaturalMortalitycm
            Dim NaturalMortalitycb As New OleDb.OleDbCommandBuilder
            NaturalMortalitycb = New OleDb.OleDbCommandBuilder(NaturalMortalityIDDA)
            If TransferDataSet.Tables.Contains("NaturalMortality") Then
                TransferDataSet.Tables("NaturalMortality").Clear()
            End If
            NaturalMortalityIDDA.Fill(TransferDataSet, "NaturalMortality")
            NaturalMortalityIDDA = Nothing

            Dim NumNaturalMortality As Integer
            NumNaturalMortality = TransferDataSet.Tables("NaturalMortality").Rows.Count
            Dim NaturalMortalityTrans As OleDb.OleDbTransaction
            Dim NaturalMortality As New OleDbCommand

            NaturalMortalityTrans = FramDB.BeginTransaction
            NaturalMortality.Connection = FramDB
            NaturalMortality.Transaction = NaturalMortalityTrans
            NumRecs = TransferDataSet.Tables("NaturalMortality").Rows.Count
            For RecNum = 0 To NumRecs - 1
                NaturalMortality.CommandText = "INSERT INTO NaturalMortality (BasePeriodID,Age,TimeStep,NaturalMortalityRate) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("NaturalMortality").Rows(RecNum)(3) & ")"

                NaturalMortality.ExecuteNonQuery()
            Next
            NaturalMortalityTrans.Commit()
            FramDB.Close()

            'ShakerMortRate      
            FramDB.Open()
            CmdStr = "SELECT * FROM ShakerMortRate WHERE BasePeriodID = " & TransBaseID & ";"
            Dim ShakerMortRatecm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim ShakerMortRateIDDA As New System.Data.OleDb.OleDbDataAdapter
            ShakerMortRateIDDA.SelectCommand = ShakerMortRatecm
            Dim ShakerMortRatecb As New OleDb.OleDbCommandBuilder
            ShakerMortRatecb = New OleDb.OleDbCommandBuilder(ShakerMortRateIDDA)
            If TransferDataSet.Tables.Contains("ShakerMortRate") Then
                TransferDataSet.Tables("ShakerMortRate").Clear()
            End If
            ShakerMortRateIDDA.Fill(TransferDataSet, "ShakerMortRate")
            ShakerMortRateIDDA = Nothing

            Dim NumShakerMortRate As Integer
            NumShakerMortRate = TransferDataSet.Tables("ShakerMortRate").Rows.Count
            Dim ShakerMortRateTrans As OleDb.OleDbTransaction
            Dim ShakerMortRate As New OleDbCommand

            ShakerMortRateTrans = FramDB.BeginTransaction
            ShakerMortRate.Connection = FramDB
            ShakerMortRate.Transaction = ShakerMortRateTrans
            NumRecs = TransferDataSet.Tables("ShakerMortRate").Rows.Count
            For RecNum = 0 To NumRecs - 1
                ShakerMortRate.CommandText = "INSERT INTO ShakerMortRate (BasePeriodID,FisheryID,TimeStep,ShakerMortRate) " & _
                  "VALUES(" & NewBasePeriodID & "," & _
                   TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(1) & "," & _
                   TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(2) & "," & _
                   TransferDataSet.Tables("ShakerMortRate").Rows(RecNum)(3) & ")"
                ShakerMortRate.ExecuteNonQuery()
            Next
            ShakerMortRateTrans.Commit()
            FramDB.Close()

            'TerminalFisheryFlag      
            FramDB.Open()
            CmdStr = "SELECT * FROM TerminalFisheryFlag WHERE BasePeriodID = " & TransBaseID & ";"
            Dim TerminalFisheryFlagcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim TerminalFisheryFlagIDDA As New System.Data.OleDb.OleDbDataAdapter
            TerminalFisheryFlagIDDA.SelectCommand = TerminalFisheryFlagcm
            Dim TerminalFisheryFlagcb As New OleDb.OleDbCommandBuilder
            TerminalFisheryFlagcb = New OleDb.OleDbCommandBuilder(TerminalFisheryFlagIDDA)
            If TransferDataSet.Tables.Contains("TerminalFisheryFlag") Then
                TransferDataSet.Tables("TerminalFisheryFlag").Clear()
            End If
            TerminalFisheryFlagIDDA.Fill(TransferDataSet, "TerminalFisheryFlag")
            TerminalFisheryFlagIDDA = Nothing

            Dim NumTerminalFisheryFlag As Integer
            NumTerminalFisheryFlag = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count
            Dim TerminalFisheryFlagTrans As OleDb.OleDbTransaction
            Dim TerminalFisheryFlag As New OleDbCommand

            TerminalFisheryFlagTrans = FramDB.BeginTransaction
            TerminalFisheryFlag.Connection = FramDB
            TerminalFisheryFlag.Transaction = TerminalFisheryFlagTrans
            NumRecs = TransferDataSet.Tables("TerminalFisheryFlag").Rows.Count
            For RecNum = 0 To NumRecs - 1
                TerminalFisheryFlag.CommandText = "INSERT INTO TerminalFisheryFlag (BasePeriodID,FisheryID,TimeStep,TerminalFlag) " & _
                   "VALUES(" & NewBasePeriodID & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("TerminalFisheryFlag").Rows(RecNum)(3) & ")"
                TerminalFisheryFlag.ExecuteNonQuery()
            Next
            TerminalFisheryFlagTrans.Commit()
            FramDB.Close()

            drd2.Close()
            cmd2.Dispose()
            drd2.Dispose()
        Next i
        '_____________________________________________________________________________
        'check whether to import new calibration tables without base period IDs. Will replace existing tables

        Dim cmd3 As New OleDb.OleDbCommand()
        cmd3.Connection = FramDB


        'import base period size limits
        If ImportBP = True Then
            'delete existing BaseSizeLimits in FRAM database
            'Dim CmdStr2 As String
            CmdStr = "SELECT * FROM ChinookBaseSizeLimit;"
            Dim ChinookBaseSizeLimit2cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim ChinookBaseSizeLimit2DA As New System.Data.OleDb.OleDbDataAdapter
            ChinookBaseSizeLimit2DA.SelectCommand = ChinookBaseSizeLimit2cm
            '- DELETE Statement
            CmdStr = "DELETE * FROM ChinookBaseSizeLimit;"
            Dim ChinookBaseSizeLimit3cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim ChinookBaseSizeLimit3da As New System.Data.OleDb.OleDbDataAdapter
            ChinookBaseSizeLimit3DA.DeleteCommand = ChinookBaseSizeLimit3cm
            '- Command Builder
            Dim ChinookBaseSizeLimit3cb As New OleDb.OleDbCommandBuilder
            ChinookBaseSizeLimit3cb = New OleDb.OleDbCommandBuilder(ChinookBaseSizeLimit3DA)
            FramDB.Open()
            ChinookBaseSizeLimit3DA.DeleteCommand.ExecuteScalar()

            'fill FRAM with new Base Period Size Limits
            CmdStr = "SELECT * FROM ChinookBaseSizeLimit;"
            Dim BaseSizeLimcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim BaseSizeLimIDDA As New System.Data.OleDb.OleDbDataAdapter
            BaseSizeLimIDDA.SelectCommand = BaseSizeLimcm
            Dim BaseSizeLimcb As New OleDb.OleDbCommandBuilder
            BaseSizeLimcb = New OleDb.OleDbCommandBuilder(BaseSizeLimIDDA)
            If TransferDataSet.Tables.Contains("ChinookBaseSizeLimit") Then
                TransferDataSet.Tables("ChinookBaseSizeLimit").Clear()
            End If
            BaseSizeLimIDDA.Fill(TransferDataSet, "ChinookBaseSizeLimit")
            BaseSizeLimIDDA = Nothing


            Dim BaseSizeLimitTrans As OleDb.OleDbTransaction
            Dim BaseSizeLimit As New OleDbCommand

            BaseSizeLimitTrans = FramDB.BeginTransaction
            BaseSizeLimit.Connection = FramDB
            BaseSizeLimit.Transaction = BaseSizeLimitTrans
            NumRecs = TransferDataSet.Tables("ChinookBaseSizeLimit").Rows.Count

            For RecNum = 0 To NumRecs - 1
                BaseSizeLimit.CommandText = "INSERT INTO ChinookBaseSizeLimit (FisheryID,Time1SizeLimit,Time2SizeLimit,Time3SizeLimit,Time4SizeLimit) " & _
                  "VALUES(" & TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(0) & "," & _
                    TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(1) & "," & _
                    TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(2) & "," & _
                    TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(3) & "," & _
                    TransferDataSet.Tables("ChinookBaseSizeLimit").Rows(RecNum)(4) & ")"
                BaseSizeLimit.ExecuteNonQuery()
            Next
            BaseSizeLimitTrans.Commit()
            FramDB.Close()
        End If

        'import stocks table
        If ImportStock = True Then
            'delete existing Stock Table in FRAM database
            'Dim CmdStr2 As String
            CmdStr = "SELECT * FROM Stock;"
            Dim Stock2cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim Stock2DA As New System.Data.OleDb.OleDbDataAdapter
            Stock2DA.SelectCommand = Stock2cm
            '- DELETE Statement
            CmdStr = "DELETE * FROM Stock;"
            Dim Stock3cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim Stock3da As New System.Data.OleDb.OleDbDataAdapter
            Stock3da.DeleteCommand = Stock3cm
            '- Command Builder
            Dim Stock3cb As New OleDb.OleDbCommandBuilder
            Stock3cb = New OleDb.OleDbCommandBuilder(Stock3da)
            FramDB.Open()
            Stock3da.DeleteCommand.ExecuteScalar()

            'fill FRAM with new Stocks
            CmdStr = "SELECT * FROM Stock;"
            Dim Stockcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim StockIDDA As New System.Data.OleDb.OleDbDataAdapter
            StockIDDA.SelectCommand = Stockcm
            Dim Stockcb As New OleDb.OleDbCommandBuilder
            Stockcb = New OleDb.OleDbCommandBuilder(StockIDDA)
            If TransferDataSet.Tables.Contains("Stock") Then
                TransferDataSet.Tables("Stock").Clear()
            End If
            StockIDDA.Fill(TransferDataSet, "Stock")
            StockIDDA = Nothing


            Dim StockTrans As OleDb.OleDbTransaction
            Dim Stock As New OleDbCommand

            StockTrans = FramDB.BeginTransaction
            Stock.Connection = FramDB
            Stock.Transaction = StockTrans
            NumRecs = TransferDataSet.Tables("Stock").Rows.Count

            For RecNum = 0 To NumRecs - 1
                Stock.CommandText = "INSERT INTO Stock (Species,StockVersion,StockID,ProductionRegionNumber,ManagementUnitNumber,StockName,StockLongName) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(2) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(3) & "," & _
                TransferDataSet.Tables("Stock").Rows(RecNum)(4) & "," & _
                Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(5) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("Stock").Rows(RecNum)(6) & Chr(34) & ")"
                Stock.ExecuteNonQuery()
            Next
            StockTrans.Commit()
            FramDB.Close()
        End If

        'import Fisherys table
        If ImportFish = True Then
            'delete existing Fishery Table in FRAM database
            'Dim CmdStr2 As String
            CmdStr = "SELECT * FROM Fishery;"
            Dim Fishery2cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim Fishery2DA As New System.Data.OleDb.OleDbDataAdapter
            Fishery2DA.SelectCommand = Fishery2cm
            '- DELETE Statement
            CmdStr = "DELETE * FROM Fishery;"
            Dim Fishery3cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim Fishery3da As New System.Data.OleDb.OleDbDataAdapter
            Fishery3da.DeleteCommand = Fishery3cm
            '- Command Builder
            Dim Fishery3cb As New OleDb.OleDbCommandBuilder
            Fishery3cb = New OleDb.OleDbCommandBuilder(Fishery3da)
            FramDB.Open()
            Fishery3da.DeleteCommand.ExecuteScalar()

            'fill FRAM with new Fisherys
            CmdStr = "SELECT * FROM Fishery;"
            Dim Fisherycm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim FisheryIDDA As New System.Data.OleDb.OleDbDataAdapter
            FisheryIDDA.SelectCommand = Fisherycm
            Dim Fisherycb As New OleDb.OleDbCommandBuilder
            Fisherycb = New OleDb.OleDbCommandBuilder(FisheryIDDA)
            If TransferDataSet.Tables.Contains("Fishery") Then
                TransferDataSet.Tables("Fishery").Clear()
            End If
            FisheryIDDA.Fill(TransferDataSet, "Fishery")
            FisheryIDDA = Nothing


            Dim FisheryTrans As OleDb.OleDbTransaction
            Dim Fishery As New OleDbCommand

            FisheryTrans = FramDB.BeginTransaction
            Fishery.Connection = FramDB
            Fishery.Transaction = FisheryTrans
            NumRecs = TransferDataSet.Tables("Fishery").Rows.Count

            For RecNum = 0 To NumRecs - 1
                Fishery.CommandText = "INSERT INTO Fishery (Species,VersionNumber,FisheryID,FisheryName,FisheryTitle) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("Fishery").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("Fishery").Rows(RecNum)(2) & "," & _
                Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(3) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("Fishery").Rows(RecNum)(4) & Chr(34) & ")"
                If TransferDataSet.Tables("Fishery").Rows(RecNum)(2) <> 74 Then
                    Fishery.ExecuteNonQuery()
                End If
            Next
            FisheryTrans.Commit()
            FramDB.Close()
        End If

        'import TimeStep table
        If ImportTS = True Then
            'delete existing TimeStep Table in FRAM database
            'Dim CmdStr2 As String
            CmdStr = "SELECT * FROM TimeStep;"
            Dim TimeStep2cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim TimeStep2DA As New System.Data.OleDb.OleDbDataAdapter
            TimeStep2DA.SelectCommand = TimeStep2cm
            '- DELETE Statement
            CmdStr = "DELETE * FROM TimeStep;"
            Dim TimeStep3cm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim TimeStep3da As New System.Data.OleDb.OleDbDataAdapter
            TimeStep3da.DeleteCommand = TimeStep3cm
            '- Command Builder
            Dim TimeStep3cb As New OleDb.OleDbCommandBuilder
            TimeStep3cb = New OleDb.OleDbCommandBuilder(TimeStep3da)
            FramDB.Open()
            TimeStep3da.DeleteCommand.ExecuteScalar()

            'fill FRAM with new TimeSteps
            CmdStr = "SELECT * FROM TimeStep;"
            Dim TimeStepcm As New OleDb.OleDbCommand(CmdStr, TransBP)
            Dim TimeStepIDDA As New System.Data.OleDb.OleDbDataAdapter
            TimeStepIDDA.SelectCommand = TimeStepcm
            Dim TimeStepcb As New OleDb.OleDbCommandBuilder
            TimeStepcb = New OleDb.OleDbCommandBuilder(TimeStepIDDA)
            If TransferDataSet.Tables.Contains("TimeStep") Then
                TransferDataSet.Tables("TimeStep").Clear()
            End If
            TimeStepIDDA.Fill(TransferDataSet, "TimeStep")
            TimeStepIDDA = Nothing


            Dim TimeStepTrans As OleDb.OleDbTransaction
            Dim TimeStep As New OleDbCommand

            TimeStepTrans = FramDB.BeginTransaction
            TimeStep.Connection = FramDB
            TimeStep.Transaction = TimeStepTrans
            NumRecs = TransferDataSet.Tables("TimeStep").Rows.Count

            For RecNum = 0 To NumRecs - 1
                TimeStep.CommandText = "INSERT INTO TimeStep (Species,VersionNumber,TimeStepID,TimeStepName,TimeStepTitle) " & _
               "VALUES(" & Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(0) & Chr(34) & "," & _
                TransferDataSet.Tables("TimeStep").Rows(RecNum)(1) & "," & _
                TransferDataSet.Tables("TimeStep").Rows(RecNum)(2) & "," & _
                Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(3) & Chr(34) & "," & _
                Chr(34) & TransferDataSet.Tables("TimeStep").Rows(RecNum)(4) & Chr(34) & ")"
                TimeStep.ExecuteNonQuery()
            Next
            TimeStepTrans.Commit()
            FramDB.Close()
        End If

        cmd3.Dispose()
        'FVS_FramUtils.Show()
    End Sub
    Sub GetTransferModelRunTables()

        '- This SubRoutine is the opposite of the TransferModelRunTables

        Dim CmdStr As String
        Dim RecNum, NumRecs, TransID, NewRunID, OldRunID As Integer

        '- First put All Records from Transfer Database into TransferDataSet (Temp Memory)

        '- Transfer RunID Records
        CmdStr = "SELECT * FROM RunID;"
        Dim RIDcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim RunIDDA As New System.Data.OleDb.OleDbDataAdapter
        RunIDDA.SelectCommand = RIDcm
        Dim RIDcb As New OleDb.OleDbCommandBuilder
        RIDcb = New OleDb.OleDbCommandBuilder(RunIDDA)
        If TransferDataSet.Tables.Contains("RunID") Then
            TransferDataSet.Tables("RunID").Clear()
        End If
        RunIDDA.Fill(TransferDataSet, "RunID")
        RunIDDA = Nothing
        Dim NumRID As Integer
        NumRID = TransferDataSet.Tables("RunID").Rows.Count
        If NumRID = 0 Then
            MsgBox("ERROR in RunID Table of Transfer Database ... ", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        '- BaseID Records Associated with Transfer RunID's
        CmdStr = "SELECT * FROM BaseID;"
        Dim BIDcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim BaseIDDA As New System.Data.OleDb.OleDbDataAdapter
        BaseIDDA.SelectCommand = BIDcm
        Dim BIDcb As New OleDb.OleDbCommandBuilder
        BIDcb = New OleDb.OleDbCommandBuilder(BaseIDDA)
        If TransferDataSet.Tables.Contains("BaseID") Then
            TransferDataSet.Tables("BaseID").Clear()
        End If
        BaseIDDA.Fill(TransferDataSet, "BaseID")
        BaseIDDA = Nothing
        '- Transfer Backwards FRAM Table
        CmdStr = "SELECT * FROM BackwardsFRAM ORDER BY StockID;"
        Dim BFcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim BFDA As New System.Data.OleDb.OleDbDataAdapter
        BFDA.SelectCommand = BFcm
        Dim BFcb As New OleDb.OleDbCommandBuilder
        BFcb = New OleDb.OleDbCommandBuilder(BFDA)
        If TransferDataSet.Tables.Contains("BackwardsFRAM") Then
            TransferDataSet.Tables("BackwardsFRAM").Clear()
        End If
        BFDA.Fill(TransferDataSet, "BackwardsFRAM")
        BFDA = Nothing
        '- Transfer FisheryScalers Table
        CmdStr = "SELECT * FROM FisheryScalers ORDER BY FisheryID, TimeStep;"
        Dim FScm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim FSDA As New System.Data.OleDb.OleDbDataAdapter
        FSDA.SelectCommand = FScm
        Dim FScb As New OleDb.OleDbCommandBuilder
        FScb = New OleDb.OleDbCommandBuilder(FSDA)
        If TransferDataSet.Tables.Contains("FisheryScalers") Then
            TransferDataSet.Tables("FisheryScalers").Clear()
        End If
        FSDA.Fill(TransferDataSet, "FisheryScalers")
        FSDA = Nothing
        '- Transfer NonRetention Table
        CmdStr = "SELECT * FROM NonRetention ORDER BY FisheryID, TimeStep;"
        Dim NRcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim NRDA As New System.Data.OleDb.OleDbDataAdapter
        NRDA.SelectCommand = NRcm
        Dim NRcb As New OleDb.OleDbCommandBuilder
        NRcb = New OleDb.OleDbCommandBuilder(NRDA)
        If TransferDataSet.Tables.Contains("NonRetention") Then
            TransferDataSet.Tables("NonRetention").Clear()
        End If
        NRDA.Fill(TransferDataSet, "NonRetention")
        NRDA = Nothing
        '- Transfer Stock/Fishery Rate Scalers
        CmdStr = "SELECT * FROM StockFisheryRateScaler ORDER BY StockID, FisheryID, TimeStep"
        Dim SFRcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim SFDA As New System.Data.OleDb.OleDbDataAdapter
        SFDA.SelectCommand = SFRcm
        Dim SFRcb As New OleDb.OleDbCommandBuilder
        SFRcb = New OleDb.OleDbCommandBuilder(SFDA)
        If TransferDataSet.Tables.Contains("StockFisheryRateScaler") Then
            TransferDataSet.Tables("StockFisheryRateScaler").Clear()
        End If
        SFDA.Fill(TransferDataSet, "StockFisheryRateScaler")
        SFDA = Nothing
        '- Transfer PSCMaxER - Coho Only
        CmdStr = "SELECT * FROM PSCMaxER ORDER BY PSCStockID"
        Dim PSCcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim PSCDA As New System.Data.OleDb.OleDbDataAdapter
        PSCDA.SelectCommand = PSCcm
        Dim PSCcb As New OleDb.OleDbCommandBuilder
        PSCcb = New OleDb.OleDbCommandBuilder(PSCDA)
        If TransferDataSet.Tables.Contains("PSCMaxER") Then
            TransferDataSet.Tables("PSCMaxER").Clear()
        End If
        PSCDA.Fill(TransferDataSet, "PSCMaxER")
        PSCDA = Nothing
        '- Size Limits - Chinook Only
        CmdStr = "SELECT * FROM SizeLimits ORDER BY FisheryID, TimeStep"
        Dim SLcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim SLDA As New System.Data.OleDb.OleDbDataAdapter
        SLDA.SelectCommand = SLcm
        Dim SLcb As New OleDb.OleDbCommandBuilder
        SLcb = New OleDb.OleDbCommandBuilder(SLDA)
        If TransferDataSet.Tables.Contains("SizeLimits") Then
            TransferDataSet.Tables("SizeLimits").Clear()
        End If
        SLDA.Fill(TransferDataSet, "SizeLimits")
        SLDA = Nothing
        '- Transfer Stock Recruits
        CmdStr = "SELECT * FROM StockRecruit ORDER BY StockID, Age"
        Dim SRcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim SRDA As New System.Data.OleDb.OleDbDataAdapter
        SRDA.SelectCommand = SRcm
        Dim SRcb As New OleDb.OleDbCommandBuilder
        SRcb = New OleDb.OleDbCommandBuilder(SRDA)
        If TransferDataSet.Tables.Contains("StockRecruit") Then
            TransferDataSet.Tables("StockRecruit").Clear()
        End If
        SRDA.Fill(TransferDataSet, "StockRecruit")
        SRDA = Nothing
        '- Transfer Cohort Run Sizes
        CmdStr = "SELECT * FROM Cohort ORDER BY StockID, Age, TimeStep"
        Dim COHcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim COHDA As New System.Data.OleDb.OleDbDataAdapter
        COHDA.SelectCommand = COHcm
        Dim COHcb As New OleDb.OleDbCommandBuilder
        COHcb = New OleDb.OleDbCommandBuilder(COHDA)
        If TransferDataSet.Tables.Contains("Cohort") Then
            TransferDataSet.Tables("Cohort").Clear()
        End If
        COHDA.Fill(TransferDataSet, "Cohort")
        COHDA = Nothing
        '- Transfer Escapement
        CmdStr = "SELECT * FROM Escapement ORDER BY StockID, Age, TimeStep"
        Dim ESCcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim ESCDA As New System.Data.OleDb.OleDbDataAdapter
        ESCDA.SelectCommand = ESCcm
        Dim ESCcb As New OleDb.OleDbCommandBuilder
        ESCcb = New OleDb.OleDbCommandBuilder(ESCDA)
        If TransferDataSet.Tables.Contains("Escapement") Then
            TransferDataSet.Tables("Escapement").Clear()
        End If
        ESCDA.Fill(TransferDataSet, "Escapement")
        ESCDA = Nothing
        '- Transfer FisheryMortality
        CmdStr = "SELECT * FROM FisheryMortality ORDER BY FisheryID, TimeStep"
        Dim FMcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim FMDA As New System.Data.OleDb.OleDbDataAdapter
        FMDA.SelectCommand = FMcm
        Dim FMcb As New OleDb.OleDbCommandBuilder
        FMcb = New OleDb.OleDbCommandBuilder(FMDA)
        If TransferDataSet.Tables.Contains("FisheryMortality") Then
            TransferDataSet.Tables("FisheryMortality").Clear()
        End If
        FMDA.Fill(TransferDataSet, "FisheryMortality")
        FMDA = Nothing
        '- Transfer All Mortality Records
        CmdStr = "SELECT * FROM Mortality ORDER BY FisheryID, TimeStep"
        Dim MRTcm As New OleDb.OleDbCommand(CmdStr, TransDB)
        Dim MRTDA As New System.Data.OleDb.OleDbDataAdapter
        MRTDA.SelectCommand = MRTcm
        Dim MRTcb As New OleDb.OleDbCommandBuilder
        MRTcb = New OleDb.OleDbCommandBuilder(MRTDA)
        If TransferDataSet.Tables.Contains("Mortality") Then
            TransferDataSet.Tables("Mortality").Clear()
        End If
        MRTDA.Fill(TransferDataSet, "Mortality")
        MRTDA = Nothing


        '==============================================================================================
        '- (Pete 12/13) Fill temp data sets with Target Sublegal:Legal Ratio (SLRatio) 
        '- and run-specific sublegal encounter rate adjustment (RunEncounterRateAdjustment) content

        '- Code that checks for the existence of the Target Sublegal:Legal Ratio 
        '- works to make things functional retroactively
        Dim sql As String       'SQL Query text string
        Dim oledbAdapter As OleDb.OleDbDataAdapter

        'First check the FRAM database for the SLRatio and RunEncounterRateAdjustment tables
        TransDB.Open()
        Dim restrictions1(3) As String
        Dim DoesTableExist1 As Boolean
        restrictions1(2) = "SLRatio"
        Dim dbTbl As DataTable = TransDB.GetSchema("Tables", restrictions1)
        If dbTbl.Rows.Count = 0 Then
            'Table does not exist
            DoesTableExist1 = False
        Else
            'Table exists
            DoesTableExist1 = True
        End If
        dbTbl.Dispose()
        TransDB.Close()

        'If SLRatio exists in this transfer get the content.
        If DoesTableExist1 = True Then

            '- Transfer Sublegal Ratios
            CmdStr = "SELECT * FROM SLRatio"
            Dim SLRatcm As New OleDb.OleDbCommand(CmdStr, TransDB)
            Dim SLRatDA As New System.Data.OleDb.OleDbDataAdapter
            SLRatDA.SelectCommand = SLRatcm
            Dim SLRatcb As New OleDb.OleDbCommandBuilder
            SLRatcb = New OleDb.OleDbCommandBuilder(SLRatDA)
            If TransferDataSet.Tables.Contains("SLRatio") Then
                TransferDataSet.Tables("SLRatio").Clear()
            End If
            SLRatDA.Fill(TransferDataSet, "SLRatio")
            SLRatDA = Nothing
        End If

        '==============================================================================================


        '- Get Current Max RunID Value, Add One for Transfer Recordset RunID Value
        Dim drd1 As OleDb.OleDbDataReader
        Dim cmd1 As New OleDb.OleDbCommand()
        Dim MaxOldID As Integer
        cmd1.Connection = FramDB
        cmd1.CommandText = "SELECT * FROM RunID ORDER BY RunID DESC"
        FramDB.Open()
        drd1 = cmd1.ExecuteReader
        drd1.Read()
        MaxOldID = drd1.GetInt32(1)
        cmd1.Dispose()
        drd1.Dispose()
        FramDB.Close()
        'RunIDTransfer = MaxOldID + 1

        '- Loop Through Transfer RunID and Add Records into FramDB with New RunID Numbers

        For TransID = 1 To NumRID
            NewRunID = MaxOldID + TransID
            OldRunID = TransferDataSet.Tables("RunID").Rows(TransID - 1)(1)
            SelectSpeciesName = TransferDataSet.Tables("RunID").Rows(TransID - 1)(2)

            '- Find BaseID Record that matches Transfer RunID
            Dim TransBaseIDName As String
            Dim TransBaseID As Integer
            Dim drd2 As OleDb.OleDbDataReader
            Dim cmd2 As New OleDb.OleDbCommand()
            TransBaseIDName = TransferDataSet.Tables("BaseID").Rows(TransID - 1)(2)
            cmd2.Connection = FramDB
            cmd2.CommandText = "SELECT * FROM BaseID WHERE BasePeriodName = " & Chr(34) & TransBaseIDName & Chr(34) & " ORDER BY BasePeriodID;"
            FramDB.Open()
            drd2 = cmd2.ExecuteReader
            'If drd2.Read() = False Then
            '    MsgBox("Can't find Matching BasePeriodName = '" & TransBaseIDName & "' in FramVS Database" & vbCrLf & _
            '    "Please Read Corresponding Base file for this Model Run Transfer", MsgBoxStyle.OkOnly)
            'Else
            '    TransBaseID = drd2.GetInt32(1)
            'End If
            '***********************************
            'code dealing with the problem of associating a transfer run to the first BaseID containing the specified base period name
            'the database does not require unique base period names so several different base periods while containing unique IDs can have
            'dublicate names. 
            If drd2.Read() = False Then
                MsgBox("Can't find Matching BasePeriodName = '" & TransBaseIDName & "' in FramVS Database" & vbCrLf & _
                "Please Read Corresponding Base file for this Model Run Transfer", MsgBoxStyle.OkOnly)
            Else
                '    TransBaseID = drd2.GetInt32(1)
                Dim N As Integer = 1
                Dim multiBaseID As String

                While drd2.Read()
                    N = N + 1
                End While
                If N > 1 Then
                    drd2.Dispose()
                    drd2 = cmd2.ExecuteReader
                    drd2.Read()
                    multiBaseID = drd2.GetInt32(1)
                    For I = 1 To N - 1
                        drd2.Read()
                        TransBaseID = drd2.GetInt32(1)
                        multiBaseID = multiBaseID & ", " & TransBaseID
                    Next I
                    TransBaseID = InputBox("Your database contains multiple base periods with the name " & BasePeriodName & ". Please enter one of the following base period IDs that you wish to associate with your run: " & multiBaseID)
                    Do While multiBaseID.Contains(TransBaseID) <> True
                        TransBaseID = InputBox("The number you entered is not a valid base period ID.Please enter one of the following base period IDs: " & multiBaseID)
                    Loop

                Else

                    drd2.Dispose()
                    drd2 = cmd2.ExecuteReader
                    drd2.Read()
                    TransBaseID = drd2.GetInt32(1)
                End If
            End If
            '*****************************************
            cmd2.Dispose()
            drd2.Dispose()
            FramDB.Close()
            Dim RIDTrans As OleDb.OleDbTransaction
            Dim RID As New OleDbCommand
            FramDB.Open()
            RIDTrans = FramDB.BeginTransaction
            RID.Connection = FramDB
            RID.Transaction = RIDTrans
            RecNum = 0
            RID.CommandText = "INSERT INTO RunID (RunID,SpeciesName,RunName,RunTitle,BasePeriodID,RunComments,CreationDate,ModifyInputDate,RunTimeDate,RunYear) " & _
               "VALUES(" & NewRunID.ToString & "," & _
               Chr(34) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(2) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(3) & Chr(34) & "," & _
               Chr(34) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(4) & Chr(34) & "," & _
               TransBaseID.ToString & "," & _
               Chr(34) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(6) & Chr(34) & "," & _
               Chr(35) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(7) & Chr(35) & "," & _
               Chr(35) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(8) & Chr(35) & "," & _
            Chr(35) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(9) & Chr(35) & "," & _
            Chr(34) & TransferDataSet.Tables("RunID").Rows(TransID - 1)(10) & Chr(34) & ")"
            RID.ExecuteNonQuery()
            RIDTrans.Commit()
            FramDB.Close()

            '- Transfer Backwards FRAM Table
            Dim j, k As Integer
            NumRecs = TransferDataSet.Tables("BackwardsFRAM").Rows.Count
            If NumRecs = 0 Then
                GoTo SkipBF
            End If

            'check if comment column exists in TransferDB
            j = TransferDataSet.Tables("BackwardsFRAM").Columns.IndexOf("Comment")
            k = FramDataSet.Tables("BackwardsFRAM").Columns.IndexOf("Comment")
            If k = -1 Then 'add column if it doesn't exist in main FRAM database
                FramDB.Open()
                Dim BKFRAMTable As String = "BackwardsFRAM"
                CmdStr = "SELECT * FROM [" & BKFRAMTable & "];"
                Dim BKFRAMcm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim BKFRAMDA As New System.Data.OleDb.OleDbDataAdapter
                BKFRAMDA.SelectCommand = BKFRAMcm
                Dim BKFRAMcb As New OleDb.OleDbCommandBuilder
                BKFRAMcb = New OleDb.OleDbCommandBuilder(BKFRAMDA)
                BKFRAMDA.Fill(FramDataSet, "BackwardsFRAM")

                BKFRAMcm.CommandText = "ALTER TABLE " & BKFRAMTable & " ADD " & "Comment" & " " & "String"
                BKFRAMcm.ExecuteNonQuery()   'executes the SQL code in cmd without querry
                FramDB.Close()
            End If


            Dim BFTrans As OleDb.OleDbTransaction
            Dim BFC As New OleDbCommand
            FramDB.Open()
            BFTrans = FramDB.BeginTransaction
            BFC.Connection = FramDB
            BFC.Transaction = BFTrans
            For RecNum = 0 To NumRecs - 1
                If j = -1 Then 'Comment column does not exist
                    '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                    If OldRunID = TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(0) Then
                        BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag) " & _
                        "VALUES(" & NewRunID.ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(1).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(2).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(3).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(4).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(5).ToString & ")"
                        BFC.ExecuteNonQuery()
                    End If
                Else 'comment column exists in TransferDB
                    
                    If OldRunID = TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(0) Then
                        BFC.CommandText = "INSERT INTO BackwardsFRAM (RunID,StockID,TargetEscAge3,TargetEscAge4,TargetEscAge5,TargetFlag,Comment) " & _
                        "VALUES(" & NewRunID.ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(1).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(2).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(3).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(4).ToString & "," & _
                        TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(5).ToString & "," & _
                        Chr(34) & TransferDataSet.Tables("BackwardsFRAM").Rows(RecNum)(6).ToString & Chr(34) & ")"
                        BFC.ExecuteNonQuery()
                    End If

                End If
            Next
            BFTrans.Commit()
            FramDB.Close()
SkipBF:

            '- Transfer FisheryScalers Table
            NumRecs = TransferDataSet.Tables("FisheryScalers").Rows.Count

            'check if comment column exists in TransferDB
            j = TransferDataSet.Tables("FisheryScalers").Columns.IndexOf("Comment")
            k = FramDataSet.Tables("FisheryScalers").Columns.IndexOf("Comment")
            If k = -1 Then 'add column if it doesn't exist in main FRAM database
                FramDB.Open()
                Dim FisheryScalersTable As String = "FisheryScalers"
                CmdStr = "SELECT * FROM [" & FisheryScalersTable & "];"
                Dim FisheryScalerscm As New OleDb.OleDbCommand(CmdStr, FramDB)
                Dim FisheryScalersDA As New System.Data.OleDb.OleDbDataAdapter
                FisheryScalersDA.SelectCommand = FisheryScalerscm
                Dim FisheryScalerscb As New OleDb.OleDbCommandBuilder
                FisheryScalerscb = New OleDb.OleDbCommandBuilder(FisheryScalersDA)
                FisheryScalersDA.Fill(FramDataSet, "FisheryScalers")

                FisheryScalerscm.CommandText = "ALTER TABLE " & FisheryScalersTable & " ADD " & "Comment" & " " & "String"
                FisheryScalerscm.ExecuteNonQuery()   'executes the SQL code in cmd without querry
                FramDB.Close()
            End If


            Dim FSTrans As OleDb.OleDbTransaction
            Dim FSC As New OleDbCommand

            '- First Check if this Transfer Database is from "Old" format
            Dim column As DataColumn
            For Each column In TransferDataSet.Tables("FisheryScalers").Columns
                If (column.ColumnName) = "MSFFisheryScaleFactor" Then GoTo FoundNewColumn
            Next

            '- "Old" format
            'MsgBox("Wrong Format for Database Tabel 'FisheryScalers' !!!!" & vbCrLf & "You have the WRONG Type database (ie Old Version VS)" & vbCrLf & _
            '       "Please Choose Another Database to use" & vbCrLf & "with this Version of FramVS (Multiple MSF)", MsgBoxStyle.OkOnly)
            FramDB.Open()
            FSTrans = FramDB.BeginTransaction
            FSC.Connection = FramDB
            FSC.Transaction = FSTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(1) Then
                    'FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MarkSelectiveFlag,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                    '- Put MSFScaler & Quota in correct field depending on FisheryFlag
                    If TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4) > 6 Then
                        FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                           "VALUES(" & NewRunID.ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & ", 0, 0," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & ")"
                    Else
                        FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                           "VALUES(" & NewRunID.ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & ", 0, 0," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                           TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & ")"
                    End If
                    FSC.ExecuteNonQuery()
                End If
            Next
            FSTrans.Commit()
            FramDB.Close()
            GoTo SkipFS

FoundNewColumn:
            'NumRecs = TransferDataSet.Tables("FisheryScalers").Rows.Count
            'Dim FSTrans As OleDb.OleDbTransaction
            'Dim FSC As New OleDbCommand

            '- "New" format
            FramDB.Open()
            FSTrans = FramDB.BeginTransaction
            FSC.Connection = FramDB
            FSC.Transaction = FSTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(1) Then

                    If j <> -1 Then
                        FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate,Comment) " & _
                                                 "VALUES(" & NewRunID.ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(7).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & "," & _
                        TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(12).ToString & "," & _
                        Chr(34) & TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(13).ToString & Chr(34) & ")"
                        FSC.ExecuteNonQuery()
                    Else
                        FSC.CommandText = "INSERT INTO FisheryScalers (RunID,FisheryID,TimeStep,FisheryFlag,FisheryScaleFactor,Quota,MSFFisheryScaleFactor,MSFQuota,MarkReleaseRate,MarkMisIDRate,UnMarkMisIDRate,MarkIncidentalRate) " & _
                         "VALUES(" & NewRunID.ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(2).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(3).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(4).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(5).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(6).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(7).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(8).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(9).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(10).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(11).ToString & "," & _
                         TransferDataSet.Tables("FisheryScalers").Rows(RecNum)(12).ToString & ")"
                        FSC.ExecuteNonQuery()
                    End If
                End If
            Next
            FSTrans.Commit()
            FramDB.Close()
SkipFS:

            '- Transfer NonRetention Table
            NumRecs = TransferDataSet.Tables("NonRetention").Rows.Count
            Dim NRTrans As OleDb.OleDbTransaction
            Dim NRC As New OleDbCommand
            FramDB.Open()
            NRTrans = FramDB.BeginTransaction
            NRC.Connection = FramDB
            NRC.Transaction = NRTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("NonRetention").Rows(RecNum)(1) Then
                    NRC.CommandText = "INSERT INTO NonRetention (RunID,FisheryID,TimeStep,NonRetentionFlag,CNRInput1,CNRInput2,CNRInput3,CNRInput4) " & _
                      "VALUES(" & NewRunID.ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(2).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(3).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(4).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(5).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(6).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(7).ToString & "," & _
                      TransferDataSet.Tables("NonRetention").Rows(RecNum)(8).ToString & ")"
                    NRC.ExecuteNonQuery()
                End If
            Next
            NRTrans.Commit()
            FramDB.Close()
SkipNR:

            '- Transfer Stock/Fishery Rate Scalers
            NumRecs = TransferDataSet.Tables("StockFisheryRateScaler").Rows.Count
            Dim SFRTrans As OleDb.OleDbTransaction
            Dim SFRC As New OleDbCommand
            FramDB.Open()
            SFRTrans = FramDB.BeginTransaction
            SFRC.Connection = FramDB
            SFRC.Transaction = SFRTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(0) Then
                    SFRC.CommandText = "INSERT INTO StockFisheryRateScaler (RunID,StockID,FisheryID,TimeStep,StockFisheryRateScaler) " & _
                     "VALUES(" & NewRunID.ToString & "," & _
                     TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(1).ToString & "," & _
                     TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(2).ToString & "," & _
                     TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(3).ToString & "," & _
                     TransferDataSet.Tables("StockFisheryRateScaler").Rows(RecNum)(4).ToString & ")"
                    SFRC.ExecuteNonQuery()
                End If
            Next
            SFRTrans.Commit()
            FramDB.Close()
            SFDA = Nothing
SkipSFR:

            '- Transfer PSCMaxER - Coho Only
            If SelectSpeciesName = "CHINOOK" Then GoTo SkipPSCER
            NumRecs = TransferDataSet.Tables("PSCMaxER").Rows.Count
            Dim PSCTrans As OleDb.OleDbTransaction
            Dim PSCC As New OleDbCommand
            FramDB.Open()
            PSCTrans = FramDB.BeginTransaction
            PSCC.Connection = FramDB
            PSCC.Transaction = PSCTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("PSCMaxER").Rows(RecNum)(0) Then
                    PSCC.CommandText = "INSERT INTO PSCMaxER (RunID,PSCStockID,PSCMaxER) " & _
                     "VALUES(" & NewRunID.ToString & "," & _
                     TransferDataSet.Tables("PSCMaxER").Rows(RecNum)(1).ToString & "," & _
                     TransferDataSet.Tables("PSCMaxER").Rows(RecNum)(2).ToString & ")"
                    PSCC.ExecuteNonQuery()
                End If
            Next
            PSCTrans.Commit()
            FramDB.Close()
SkipPSCER:

            '- Size Limits - Chinook Only
            If SelectSpeciesName = "COHO" Then GoTo SkipSL
            NumRecs = TransferDataSet.Tables("SizeLimits").Rows.Count
            Dim SLTrans As OleDb.OleDbTransaction
            Dim SLC As New OleDbCommand
            FramDB.Open()
            SLTrans = FramDB.BeginTransaction
            SLC.Connection = FramDB
            SLC.Transaction = SLTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("SizeLimits").Rows(RecNum)(1) Then
                    SLC.CommandText = "INSERT INTO SizeLimits (RunID,FisheryID,TimeStep,MinimumSize,MaximumSize) " & _
                     "VALUES(" & NewRunID.ToString & "," & _
                     TransferDataSet.Tables("SizeLimits").Rows(RecNum)(2).ToString & "," & _
                     TransferDataSet.Tables("SizeLimits").Rows(RecNum)(3).ToString & "," & _
                     TransferDataSet.Tables("SizeLimits").Rows(RecNum)(4).ToString & "," & _
                     TransferDataSet.Tables("SizeLimits").Rows(RecNum)(5).ToString & ")"
                    SLC.ExecuteNonQuery()
                End If
            Next
            SLTrans.Commit()
            FramDB.Close()
SkipSL:

            '- Transfer Stock Recruits
            NumRecs = TransferDataSet.Tables("StockRecruit").Rows.Count
            Dim SRTrans As OleDb.OleDbTransaction
            Dim SRC As New OleDbCommand
            FramDB.Open()
            SRTrans = FramDB.BeginTransaction
            SRC.Connection = FramDB
            SRC.Transaction = SRTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("StockRecruit").Rows(RecNum)(1) Then
                    SRC.CommandText = "INSERT INTO StockRecruit (RunID,StockID,Age,RecruitScaleFactor,RecruitCohortSize) " & _
                     "VALUES(" & NewRunID.ToString & "," & _
                     TransferDataSet.Tables("StockRecruit").Rows(RecNum)(2).ToString & "," & _
                     TransferDataSet.Tables("StockRecruit").Rows(RecNum)(3).ToString & "," & _
                     TransferDataSet.Tables("StockRecruit").Rows(RecNum)(4).ToString & "," & _
                     TransferDataSet.Tables("StockRecruit").Rows(RecNum)(5).ToString & ")"
                    SRC.ExecuteNonQuery()
                End If
            Next
            SRTrans.Commit()
            FramDB.Close()
SkipSR:

            '- Transfer Cohort Run Sizes
            NumRecs = TransferDataSet.Tables("Cohort").Rows.Count
            Dim COHTrans As OleDb.OleDbTransaction
            Dim COHC As New OleDbCommand
            FramDB.Open()
            COHTrans = FramDB.BeginTransaction
            COHC.Connection = FramDB
            COHC.Transaction = COHTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("Cohort").Rows(RecNum)(1) Then
                    COHC.CommandText = "INSERT INTO Cohort (RunID,StockID,Age,TimeStep,Cohort,MatureCohort,StartCohort,WorkingCohort,MidCohort) " & _
                       "VALUES(" & NewRunID.ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(2).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(3).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(4).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(5).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(6).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(7).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(8).ToString & "," & _
                       TransferDataSet.Tables("Cohort").Rows(RecNum)(9).ToString & ")"
                    COHC.ExecuteNonQuery()
                End If
            Next
            COHTrans.Commit()
            FramDB.Close()

            '- Transfer Escapement
            NumRecs = TransferDataSet.Tables("Escapement").Rows.Count
            Dim ESCTrans As OleDb.OleDbTransaction
            Dim ESCC As New OleDbCommand
            FramDB.Open()
            ESCTrans = FramDB.BeginTransaction
            ESCC.Connection = FramDB
            ESCC.Transaction = ESCTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("Escapement").Rows(RecNum)(1) Then
                    ESCC.CommandText = "INSERT INTO Escapement (RunID,StockID,Age,TimeStep,Escapement) " & _
                       "VALUES(" & NewRunID.ToString & "," & _
                       TransferDataSet.Tables("Escapement").Rows(RecNum)(2).ToString & "," & _
                       TransferDataSet.Tables("Escapement").Rows(RecNum)(3).ToString & "," & _
                       TransferDataSet.Tables("Escapement").Rows(RecNum)(4).ToString & "," & _
                       TransferDataSet.Tables("Escapement").Rows(RecNum)(5).ToString & ")"
                    ESCC.ExecuteNonQuery()
                End If
            Next
            ESCTrans.Commit()
            FramDB.Close()

            '- Transfer FisheryMortality
            NumRecs = TransferDataSet.Tables("FisheryMortality").Rows.Count
            Dim FMTrans As OleDb.OleDbTransaction
            Dim FMC As New OleDbCommand

            '- First Check if this Transfer Database is from "Old" format
            For Each column In TransferDataSet.Tables("FisheryMortality").Columns
                If (column.ColumnName) = "TotalLegalShakers" Then
                    '- "Old" format
                    FramDB.Open()
                    FMTrans = FramDB.BeginTransaction
                    FMC.Connection = FramDB
                    FMC.Transaction = FMTrans
                    For RecNum = 0 To NumRecs - 1
                        '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                        If OldRunID = TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(0) Then
                            FMC.CommandText = "INSERT INTO FisheryMortality (RunID,FisheryID,TimeStep,TotalLandedCatch,TotalUnMarkedCatch,TotalNonRetention,TotalShakers,TotalDropOff,TotalEncounters) " & _
                               "VALUES(" & NewRunID.ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(1).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(2).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(3).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(4).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(5).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(7).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(8).ToString & "," & _
                               TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(9).ToString & ")"
                            FMC.ExecuteNonQuery()
                        End If
                    Next
                    GoTo CommitTFM
                End If
            Next

            '- "New" format
            FramDB.Open()
            FMTrans = FramDB.BeginTransaction
            FMC.Connection = FramDB
            FMC.Transaction = FMTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(0) Then
                    FMC.CommandText = "INSERT INTO FisheryMortality (RunID,FisheryID,TimeStep,TotalLandedCatch,TotalUnMarkedCatch,TotalNonRetention,TotalShakers,TotalDropOff,TotalEncounters) " & _
                       "VALUES(" & NewRunID.ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(1).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(2).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(3).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(4).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(5).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(6).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(7).ToString & "," & _
                       TransferDataSet.Tables("FisheryMortality").Rows(RecNum)(8).ToString & ")"
                    FMC.ExecuteNonQuery()
                End If
            Next
CommitTFM:
            FMTrans.Commit()
            FramDB.Close()

            '- Transfer All Mortality Records
            NumRecs = TransferDataSet.Tables("Mortality").Rows.Count
            Dim MRTTrans As OleDb.OleDbTransaction
            Dim MRTC As New OleDbCommand

            '- First Check if this Transfer Database is from "Old" format
            For Each column In TransferDataSet.Tables("Mortality").Columns
                If (column.ColumnName) = "LegalShaker" Then
                    '- "Old" format
                    FramDB.Open()
                    MRTTrans = FramDB.BeginTransaction
                    MRTC.Connection = FramDB
                    MRTC.Transaction = MRTTrans
                    For RecNum = 0 To NumRecs - 1
                        '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                        If OldRunID = TransferDataSet.Tables("Mortality").Rows(RecNum)(1) Then
                            '- Check if "Old" LegalShaker is Non-Zero (ie MSF) and put other values into "New" fields
                            If TransferDataSet.Tables("Mortality").Rows(RecNum)(9) > 0 Then
                                MRTC.CommandText = "INSERT INTO Mortality (PrimaryKey,RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                                   "VALUES(" & (RecNum + 1).ToString & "," & NewRunID.ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(2).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(3).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(4).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(5).ToString & ",0,0,0,0,0," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(6).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(7).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(8).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(10).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(11).ToString & ")"
                            Else
                                MRTC.CommandText = "INSERT INTO Mortality (PrimaryKey,RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                                   "VALUES(" & (RecNum + 1).ToString & "," & NewRunID.ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(2).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(3).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(4).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(5).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(6).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(7).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(8).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(10).ToString & "," & _
                                   TransferDataSet.Tables("Mortality").Rows(RecNum)(11).ToString & ",0,0,0,0,0)"
                            End If
                            MRTC.ExecuteNonQuery()
                        End If
                    Next
                    GoTo CommitMorts
                End If
            Next

            '- "New" format
            FramDB.Open()
            MRTTrans = FramDB.BeginTransaction
            MRTC.Connection = FramDB
            MRTC.Transaction = MRTTrans
            For RecNum = 0 To NumRecs - 1
                '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                If OldRunID = TransferDataSet.Tables("Mortality").Rows(RecNum)(1) Then
                    MRTC.CommandText = "INSERT INTO Mortality (PrimaryKey,RunID,StockID,Age,FisheryID,TimeStep,LandedCatch,NonRetention,Shaker,DropOff,Encounter,MSFLandedCatch,MSFNonRetention,MSFShaker,MSFDropOff,MSFEncounter) " & _
                       "VALUES(" & (RecNum + 1).ToString & "," & NewRunID.ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(2).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(3).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(4).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(5).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(6).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(7).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(8).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(9).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(10).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(11).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(12).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(13).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(14).ToString & "," & _
                       TransferDataSet.Tables("Mortality").Rows(RecNum)(15).ToString & ")"
                    MRTC.ExecuteNonQuery()
                End If
            Next
CommitMorts:
            MRTTrans.Commit()
            FramDB.Close()


            '==============================================================================================
            '- (Pete 12/13) Code that creates transfers the Target Sublegal:Legal Ratio (SLRatio) 
            '- and run-specific sublegal encounter rate adjustment (RunEncounterRateAdjustment) table
            '- content associated withe runs in question.

            If DoesTableExist1 = True Then
                '- Transfer SLRatio
                NumRecs = TransferDataSet.Tables("SLRatio").Rows.Count
                'If NumRecs = 0 Then
                '   'MsgBox("Error in StockRecruit Table Transfer .. No Records", MsgBoxStyle.OkOnly)
                '   GoTo SkipSLRat
                'End If
                Dim SLRatTrans As OleDb.OleDbTransaction
                Dim SLRatC As New OleDbCommand
                FramDB.Open()
                SLRatTrans = FramDB.BeginTransaction
                SLRatC.Connection = FramDB
                SLRatC.Transaction = SLRatTrans
                For RecNum = 0 To NumRecs - 1
                    '- Check to see if record matches OldRunID being Tranferred in this RunID Loop
                    If OldRunID = TransferDataSet.Tables("SLRatio").Rows(RecNum)(0) Then
                        SLRatC.CommandText = "INSERT INTO SLRatio (RunID,FisheryID,Age,TimeStep,TargetRatio,RunEncounterRateAdjustment, UpdateWhen, UpdateBy) " & _
                         "VALUES(" & NewRunID.ToString & "," & _
                         TransferDataSet.Tables("SLRatio").Rows(RecNum)(1).ToString & "," & _
                         TransferDataSet.Tables("SLRatio").Rows(RecNum)(2).ToString & "," & _
                         TransferDataSet.Tables("SLRatio").Rows(RecNum)(3).ToString & "," & _
                         TransferDataSet.Tables("SLRatio").Rows(RecNum)(4).ToString & "," & _
                         TransferDataSet.Tables("SLRatio").Rows(RecNum)(5).ToString & "," & _
                         "'" & TransferDataSet.Tables("SLRatio").Rows(RecNum)(6).ToString & "'" & "," & _
                         "'" & TransferDataSet.Tables("SLRatio").Rows(RecNum)(7).ToString & "'" & ")"
                        SLRatC.ExecuteNonQuery()
                    End If
                Next
                SLRatTrans.Commit()
                FramDB.Close()
SkipSLRat:
            End If
            '==============================================================================================


            'End of RunID Transfer Loop
        Next

    End Sub

End Module
