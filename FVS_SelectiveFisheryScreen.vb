Imports Microsoft.Office.Interop
Imports System.IO.File

Public Class FVS_SelectiveFisheryScreen
   Public ComboFishIndex(,), NumMSF, NumContribStk As Integer

   Private Sub FVS_SelectiveFisheryScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      Dim ComboLine As String
      ReDim ComboFishIndex(NumFish, 2)

      'FormHeight = 926
      FormHeight = 946
      FormWidth = 1234
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
         If FVS_SelectiveFisheryScreen_ReSize = False Then
            Resize_Form(Me)
            FVS_SelectiveFisheryScreen_ReSize = True
         End If
      End If

      NumMSF = 0
      MSFComboBox.SelectedIndex = -1
      MSFComboBox.Items.Clear()
      MSFSelectedLabel.Text = "MSF-Selection"
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) > 6 Then
               ComboLine = FisheryName(Fish) & " - " & TimeStepTitle(TStep)
               MSFComboBox.Items.Add(ComboLine)
               NumMSF += 1
               ComboFishIndex(NumMSF, 1) = Fish
               ComboFishIndex(NumMSF, 2) = TStep
            End If
         Next
      Next

      MSFGrid.Columns.Clear()

   End Sub

   Private Sub MSFExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSFExitButton.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
   End Sub

   Private Sub MSFComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MSFComboBox.SelectedIndexChanged
      Dim SelectedFishery, SelectedTimeStep, NamePos As Integer
      Dim Tempval As Double
      Dim StkTotal As Double
      Dim TotUEnc, TotUCat, TotUNon, TotUShk, TotUSub As Double
      Dim TotMEnc, TotMCat, TotMNon, TotMShk, TotMSub As Double

      If MSFComboBox.SelectedIndex = -1 Then Exit Sub
      SelectedFishery = ComboFishIndex(MSFComboBox.SelectedIndex + 1, 1)
      SelectedTimeStep = ComboFishIndex(MSFComboBox.SelectedIndex + 1, 2)
      NumContribStk = 0
      Fish = SelectedFishery
      TStep = SelectedTimeStep
      For Stk As Integer = 1 To NumStk
         For Age As Integer = MinAge To MaxAge
            Tempval = MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
            If Tempval > 0 Then
               NumContribStk += 1
            End If
         Next
      Next
      MSFSelectedLabel.Text = FisheryTitle(SelectedFishery) & "-" & TimeStepTitle(SelectedTimeStep)
      MSFGrid.Columns.Clear()
      MSFGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      MSFGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         MSFGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         MSFGrid.Columns.Add("Name", "StockName")
         MSFGrid.Columns(0).Width = 250 / FormWidthScaler
         MSFGrid.Columns(0).ReadOnly = True
         MSFGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         MSFGrid.Columns.Add("Age", "Age")
         MSFGrid.Columns(1).Width = 50 / FormWidthScaler
         MSFGrid.Columns(1).DefaultCellStyle.Format = ("0")
         MSFGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T1", "UnM Hand")
         MSFGrid.Columns(2).Width = 80 / FormWidthScaler
         MSFGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T2", "UnM Cat")
         MSFGrid.Columns(3).Width = 80 / FormWidthScaler
         MSFGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T3", "UnM NonR")
         MSFGrid.Columns(4).Width = 80 / FormWidthScaler
         MSFGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T4", "UnM Drop")
         MSFGrid.Columns(5).Width = 80 / FormWidthScaler
         MSFGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T5", "UnM SbLg")
         MSFGrid.Columns(6).Width = 80 / FormWidthScaler
         MSFGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T6", "Mrk Hand")
         MSFGrid.Columns(7).Width = 80 / FormWidthScaler
         MSFGrid.Columns(7).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T7", "Mrk Cat")
         MSFGrid.Columns(8).Width = 80 / FormWidthScaler
         MSFGrid.Columns(8).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T8", "Mrk NonR")
         MSFGrid.Columns(9).Width = 80 / FormWidthScaler
         MSFGrid.Columns(9).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T9", "Mrk Drop")
         MSFGrid.Columns(10).Width = 80 / FormWidthScaler
         MSFGrid.Columns(10).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T10", "Mrk SbLg")
         MSFGrid.Columns(11).Width = 80 / FormWidthScaler
         MSFGrid.Columns(11).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         '- Old Style Report
         'MSFGrid.Columns.Add("T1", "UnMrk Handled")
         'MSFGrid.Columns(2).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T2", "UnMrk Catch")
         'MSFGrid.Columns(3).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T3", "UnMrk NonRet")
         'MSFGrid.Columns(4).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T4", "UnMrk DropOff")
         'MSFGrid.Columns(5).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T5", "UnMrk SubLeg")
         'MSFGrid.Columns(6).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T6", "Markd Handled")
         'MSFGrid.Columns(7).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(7).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T7", "Markd Catch")
         'MSFGrid.Columns(8).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(8).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T8", "Markd NonRet")
         'MSFGrid.Columns(9).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(9).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T9", "Markd DropOff")
         'MSFGrid.Columns(10).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(10).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         'MSFGrid.Columns.Add("T10", "Markd SubLeg")
         'MSFGrid.Columns(11).Width = 90 / FormWidthScaler
         'MSFGrid.Columns(11).DefaultCellStyle.Format = ("########0")
         'MSFGrid.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.RowCount = NumContribStk + 1
      ElseIf SpeciesName = "COHO" Then
         MSFGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         MSFGrid.Columns.Add("Name", "StockName")
         MSFGrid.Columns(0).Width = 250 / FormWidthScaler
         MSFGrid.Columns(0).ReadOnly = True
         MSFGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         MSFGrid.Columns.Add("T1", "UnMrk Handled")
         MSFGrid.Columns(1).Width = 100 / FormWidthScaler
         MSFGrid.Columns(1).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T2", "UnMrk Catch")
         MSFGrid.Columns(2).Width = 100 / FormWidthScaler
         MSFGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T3", "UnMrk NonRet")
         MSFGrid.Columns(3).Width = 100 / FormWidthScaler
         MSFGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T4", "UnMrk DropOff")
         MSFGrid.Columns(4).Width = 100 / FormWidthScaler
         MSFGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T5", "Markd Handled")
         MSFGrid.Columns(5).Width = 100 / FormWidthScaler
         MSFGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T6", "Markd Catch")
         MSFGrid.Columns(6).Width = 100 / FormWidthScaler
         MSFGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T7", "Markd NonRet")
         MSFGrid.Columns(7).Width = 100 / FormWidthScaler
         MSFGrid.Columns(7).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.Columns.Add("T8", "Markd DropOff")
         MSFGrid.Columns(8).Width = 100 / FormWidthScaler
         MSFGrid.Columns(8).DefaultCellStyle.Format = ("########0")
         MSFGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MSFGrid.RowCount = NumContribStk + 1
      End If

      '- Get Data for each Stock/Age
      Fish = SelectedFishery
      TStep = SelectedTimeStep
      NumContribStk = 0
      For Stk As Integer = 1 To NumStk Step 2
         For Age As Integer = MinAge To MaxAge
            StkTotal = 0
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
               NumContribStk += 1
               If SpeciesName = "CHINOOK" Then
                  '- CHINOOK Stock Title has UnMarked at beginning of string
                  NamePos = InStr(StockTitle(Stk), "UnMarked")
                  If NamePos > 0 Then
                     MSFGrid.Item(0, NumContribStk - 1).Value = StockTitle(Stk).Substring(9, StockTitle(Stk).Length - 9)
                  Else
                     MSFGrid.Item(0, NumContribStk - 1).Value = StockTitle(Stk)
                  End If
                  MSFGrid.Item(1, NumContribStk - 1).Value = Age.ToString
                  MSFGrid.Item(2, NumContribStk - 1).Value = CLng(MSFEncounters(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(3, NumContribStk - 1).Value = CLng(MSFLandedCatch(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(4, NumContribStk - 1).Value = CLng(MSFNonRetention(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(5, NumContribStk - 1).Value = CLng(MSFDropOff(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(6, NumContribStk - 1).Value = CLng(MSFShakers(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(7, NumContribStk - 1).Value = CLng(MSFEncounters(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(8, NumContribStk - 1).Value = CLng(MSFLandedCatch(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(9, NumContribStk - 1).Value = CLng(MSFNonRetention(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(10, NumContribStk - 1).Value = CLng(MSFDropOff(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(11, NumContribStk - 1).Value = CLng(MSFShakers(Stk + 1, Age, Fish, TStep).ToString)
               Else
                  '- COHO Stock Title has UnMarked at end of string
                  NamePos = InStr(StockTitle(Stk), "UnMarked")
                  If NamePos > 0 Then
                     MSFGrid.Item(0, NumContribStk - 1).Value = StockTitle(Stk).Substring(0, NamePos - 2)
                  Else
                     MSFGrid.Item(0, NumContribStk - 1).Value = StockTitle(Stk)
                  End If
                  MSFGrid.Item(1, NumContribStk - 1).Value = CLng(MSFEncounters(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(2, NumContribStk - 1).Value = CLng(MSFLandedCatch(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(3, NumContribStk - 1).Value = CLng(MSFNonRetention(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(4, NumContribStk - 1).Value = CLng(MSFDropOff(Stk, Age, Fish, TStep).ToString)
                  MSFGrid.Item(5, NumContribStk - 1).Value = CLng(MSFEncounters(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(6, NumContribStk - 1).Value = CLng(MSFLandedCatch(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(7, NumContribStk - 1).Value = CLng(MSFNonRetention(Stk + 1, Age, Fish, TStep).ToString)
                  MSFGrid.Item(8, NumContribStk - 1).Value = CLng(MSFDropOff(Stk + 1, Age, Fish, TStep).ToString)
               End If
            End If
         Next
      Next

      '- Totals Line
      If SpeciesName = "CHINOOK" Then
         MSFGrid.Item(0, NumContribStk).Value = "Total"
         MSFGrid.Item(1, NumContribStk).Value = "*"
         MSFGrid.Item(2, NumContribStk).Value = CLng(TotUEnc.ToString)
         MSFGrid.Item(3, NumContribStk).Value = CLng(TotUCat.ToString)
         MSFGrid.Item(4, NumContribStk).Value = CLng(TotUNon.ToString)
         MSFGrid.Item(5, NumContribStk).Value = CLng(TotUShk.ToString)
         MSFGrid.Item(6, NumContribStk).Value = CLng(TotUSub.ToString)
         MSFGrid.Item(7, NumContribStk).Value = CLng(TotMEnc.ToString)
         MSFGrid.Item(8, NumContribStk).Value = CLng(TotMCat.ToString)
         MSFGrid.Item(9, NumContribStk).Value = CLng(TotMNon.ToString)
         MSFGrid.Item(10, NumContribStk).Value = CLng(TotMShk.ToString)
         MSFGrid.Item(11, NumContribStk).Value = CLng(TotMSub.ToString)
      Else
         MSFGrid.Item(0, NumContribStk).Value = "Total"
         MSFGrid.Item(1, NumContribStk).Value = CLng(TotUEnc.ToString)
         MSFGrid.Item(2, NumContribStk).Value = CLng(TotUCat.ToString)
         MSFGrid.Item(3, NumContribStk).Value = CLng(TotUNon.ToString)
         MSFGrid.Item(4, NumContribStk).Value = CLng(TotUShk.ToString)
         MSFGrid.Item(5, NumContribStk).Value = CLng(TotMEnc.ToString)
         MSFGrid.Item(6, NumContribStk).Value = CLng(TotMCat.ToString)
         MSFGrid.Item(7, NumContribStk).Value = CLng(TotMNon.ToString)
         MSFGrid.Item(8, NumContribStk).Value = CLng(TotMShk.ToString)
      End If

   End Sub

   Private Sub ClipBoardCopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClipBoardCopyToolStripMenuItem.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      If SpeciesName = "CHINOOK" Then
         ClipStr = "CHINOOK "
      ElseIf SpeciesName = "COHO" Then
         ClipStr = "COHO "
      End If
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr
      If SpeciesName = "CHINOOK" Then
         ClipStr &= "StockName" & vbTab & "Age" & vbTab & "UnMrk Handled" & vbTab & "UnMrk Catch" & vbTab & "UnMrk NonRet" & vbTab & "UnMrk DropOff" & vbTab & "UnMrk SubLeg" & vbTab & "Marked Handled" & vbTab & "Marked Catch" & vbTab & "Marked NonRet" & vbTab & "Marked DropOff" & vbTab & "Marked SubLeg" & vbCr
         For RecNum = 0 To NumContribStk
            For ColNum = 0 To 11
               If ColNum = 0 Then
                  ClipStr &= MSFGrid.Item(ColNum, RecNum).Value
               Else
                  If ColNum = 1 And RecNum = NumContribStk Then
                     ClipStr &= vbTab & "*"
                  Else
                     ClipStr &= vbTab & CLng(MSFGrid.Item(ColNum, RecNum).Value)
                  End If
               End If
            Next
            ClipStr &= vbCr
         Next
      ElseIf SpeciesName = "COHO" Then
         ClipStr &= "StockName" & vbTab & "UnMrk Handled" & vbTab & "UnMrk Catch" & vbTab & "UnMrk NonRet" & vbTab & "UnMrk DropOff" & vbTab & "Marked Handled" & vbTab & "Marked Catch" & vbTab & "Marked NonRet" & vbTab & "Marked DropOff" & vbCr
         For RecNum = 0 To NumContribStk
            For ColNum = 0 To 8
               If ColNum = 0 Then
                  ClipStr = ClipStr & MSFGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & CLng(MSFGrid.Item(ColNum, RecNum).Value)
               End If
            Next
            ClipStr &= vbCr
         Next
      End If
      Clipboard.SetDataObject(ClipStr)

   End Sub

   Private Sub btn_MSFreport_Click(sender As System.Object, e As System.EventArgs) Handles btn_MSFreport.Click
      'This code creates a spreadsheet MSF report that is requested annualy by WDFW-Puget Sound Sampling for management
      'And pre- vs. post-season FRAM vs. Creel comparisons

      Dim MSFList(NumFish, NumSteps) As Integer '0/1 array for no/yes on the existence of MSF regulations
      Dim StepName(NumFish, NumSteps) As String
      Dim TotUEncWDFW(NumFish, NumSteps), TotUCatWDFW(NumFish, NumSteps), TotUNonWDFW(NumFish, NumSteps), TotUShkWDFW(NumFish, NumSteps), TotUSubWDFW(NumFish, NumSteps) As Double
      Dim TotMEncWDFW(NumFish, NumSteps), TotMCatWDFW(NumFish, NumSteps), TotMNonWDFW(NumFish, NumSteps), TotMShkWDFW(NumFish, NumSteps), TotMSubWDFW(NumFish, NumSteps) As Double
      Dim CountMSF As Integer
      Dim R As Integer 'Row
      Dim C As String 'Column Letter
      Dim CellAdd As String


      If SpeciesName = "COHO" Then

         MessageBox.Show("The WDFW MSF Report is currently implemented for Chinook Only")
         Exit Sub

      Else 'It's a Chinook-only report at the moment


         'Step 1: Open and access Microsoft Excel
         Dim appXL As Excel.Application
         Dim wbXl As Excel.Workbook
         Dim shXL As Excel.Worksheet
         Dim raXL As Excel.Range

         ' Start Excel and get Application object.
         appXL = CreateObject("Excel.Application")
         appXL.Visible = True
         ' Add a new workbook.
         wbXl = appXL.Workbooks.Add
         shXL = wbXl.ActiveSheet

         'Format the cells to have integer precision
         ' AutoFit columns A:D.
         raXL = shXL.Range("A1", "M500")
         raXL.NumberFormat = "#,##0"



         'Step 2: Get a list of fisheries and time steps (2-4 only) for which MSF regulations exist

         CountMSF = 0
         For Fish = 1 To NumFish
            For TStep = 2 To NumSteps
               If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 18 Then
                  'Only due WA Sport Fisheries with MSF regs
                  If Fish = 18 Or Fish = 22 Or Fish = 27 Or Fish = 36 Or Fish = 42 Or Fish = 45 Or Fish = 53 Or Fish = 54 Or Fish = 56 Or Fish = 57 Or Fish = 62 Or Fish = 64 Or Fish = 67 Then
                     If MSFFisheryQuota(Fish, TStep) > 0 Or MSFFisheryScaler(Fish, TStep) > 0 Then
                        MSFList(Fish, TStep) = 1
                        CountMSF = CountMSF + 1
                        If TStep = 2 Then
                           StepName(Fish, TStep) = "May-Jun"
                        ElseIf TStep = 3 Then
                           StepName(Fish, TStep) = "Jul-Sep"
                        Else
                           StepName(Fish, TStep) = "Oct-Apr"
                        End If
                     End If
                  End If
               Else
                  MSFList(Fish, TStep) = 0
               End If
            Next
         Next

         'Step 3: Sum up Encounters, Mortalities, etc. for MSFs
         For Fish = 1 To NumFish
            For TStep = 2 To NumSteps
               For Stk As Integer = 1 To NumStk Step 2
                  For Age As Integer = MinAge To MaxAge
                     TotUEncWDFW(Fish, TStep) += MSFEncounters(Stk, Age, Fish, TStep)
                     TotUCatWDFW(Fish, TStep) += MSFLandedCatch(Stk, Age, Fish, TStep)
                     TotUNonWDFW(Fish, TStep) += MSFNonRetention(Stk, Age, Fish, TStep)
                     TotUShkWDFW(Fish, TStep) += MSFDropOff(Stk, Age, Fish, TStep)
                     TotUSubWDFW(Fish, TStep) += MSFShakers(Stk, Age, Fish, TStep)
                     TotMEncWDFW(Fish, TStep) += MSFEncounters(Stk + 1, Age, Fish, TStep)
                     TotMCatWDFW(Fish, TStep) += MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                     TotMNonWDFW(Fish, TStep) += MSFNonRetention(Stk + 1, Age, Fish, TStep)
                     TotMShkWDFW(Fish, TStep) += MSFDropOff(Stk + 1, Age, Fish, TStep)
                     TotMSubWDFW(Fish, TStep) += MSFShakers(Stk + 1, Age, Fish, TStep)
                  Next
               Next
               TotUEncWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotUCatWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotUNonWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotUShkWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotUSubWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotMEncWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotMCatWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotMNonWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotMShkWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
               TotMSubWDFW(Fish, TStep) *= 1 / ModelStockProportion(Fish)
            Next
         Next

         'Step 4: Create General Output Table
         Dim Header As String
         Header = "Table 1. FRAM Estimates of Chinook encounters and mortalities in WA sport MSFs (Model Run: " & _
            RunIDNameSelect & ", Report Created: " & DateTime.Now.ToString & ")"
         shXL.Range("A1").Value = Header
         shXL.Range("A4").Value = "Area"
         shXL.Range("B4").Value = "Period"
         shXL.Range("C3:G3").Value = "Marked"
         shXL.Range("H3:L3").Value = "UnMark"
         shXL.Range("C4").Value = "Handled"
         shXL.Range("D4").Value = "Catch"
         shXL.Range("E4").Value = "NonRet"
         shXL.Range("F4").Value = "Dropof"
         shXL.Range("G4").Value = "SubLeg"
         shXL.Range("H4").Value = "Handled"
         shXL.Range("I4").Value = "Catch"
         shXL.Range("J4").Value = "NonRet"
         shXL.Range("K4").Value = "Dropof"
         shXL.Range("L4").Value = "SubLeg"

         R = 5
         For Fish = 1 To NumFish
            For TStep = 2 To NumSteps
               If MSFList(Fish, TStep) = 1 Then
                  shXL.Range("A" & R).Value = FisheryName(Fish)
                  shXL.Range("B" & R).Value = StepName(Fish, TStep)
                  shXL.Range("C" & R).Value = TotMEncWDFW(Fish, TStep)
                  shXL.Range("D" & R).Value = TotMCatWDFW(Fish, TStep)
                  shXL.Range("E" & R).Value = TotMNonWDFW(Fish, TStep)
                  shXL.Range("F" & R).Value = TotMShkWDFW(Fish, TStep)
                  shXL.Range("G" & R).Value = TotMSubWDFW(Fish, TStep)
                  shXL.Range("H" & R).Value = TotUEncWDFW(Fish, TStep)
                  shXL.Range("I" & R).Value = TotUCatWDFW(Fish, TStep)
                  shXL.Range("J" & R).Value = TotUNonWDFW(Fish, TStep)
                  shXL.Range("K" & R).Value = TotUShkWDFW(Fish, TStep)
                  shXL.Range("L" & R).Value = TotUSubWDFW(Fish, TStep)
                  R = R + 1
               End If
            Next
         Next

         raXL = shXL.Range("A4", "M4")
         raXL.Font.Bold = True
         raXL = shXL.Range("A3", "L4")
         raXL.Font.Bold = True
         raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
         raXL = shXL.Range("A3", "L" & 4 + CountMSF)
         raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
         raXL = shXL.Range("B3", "B" & 4 + CountMSF)
         raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
         raXL.Font.Bold = True
         raXL = shXL.Range("A1", "B1000")
         raXL.Font.Bold = True
         raXL = shXL.Range("A1")
         raXL.ColumnWidth = 10



         'Step 5: Create Individual Output Tables
         R = 6 + CountMSF 'Start dropping in tables a little farther down the worksheet 

         For Fish = 1 To NumFish
            For TStep = 2 To NumSteps
               If MSFList(Fish, TStep) = 1 Then

                  'Some prettying up of things
                  raXL = shXL.Range("A" & R + 2, "M" & R + 3)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
                  raXL = shXL.Range("A" & R, "M" & R + 1)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
                  raXL.Font.Bold = True
                  raXL = shXL.Range("B" & R + 4, "M" & R + 4)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
                  raXL.Font.Bold = True
                  raXL = shXL.Range("B" & R, "B" & R + 4)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
                  raXL = shXL.Range("F" & R, "F" & R + 4)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)
                  raXL = shXL.Range("J" & R, "J" & R + 4)
                  raXL.BorderAround(Weight:=Excel.XlBorderWeight.xlMedium)

                  'Now the tedius part -- headers, etc.
                  CellAdd = "A" & R + 1
                  shXL.Range(CellAdd).Value = FisheryName(Fish)
                  CellAdd = "A" & R + 2
                  shXL.Range(CellAdd).Value = "Legal"
                  CellAdd = "A" & R + 3
                  shXL.Range(CellAdd).Value = "Sublegal"
                  CellAdd = "B" & R + 1
                  shXL.Range(CellAdd).Value = StepName(Fish, TStep)
                  CellAdd = "B" & R + 4
                  shXL.Range(CellAdd).Value = "Total"

                  'Total Encounters
                  CellAdd = "C" & R
                  shXL.Range(CellAdd).Value = "Total Encounters"
                  CellAdd = "C" & R + 1
                  shXL.Range(CellAdd).Value = "Mark"
                  CellAdd = "C" & R + 2
                  shXL.Range(CellAdd).Value = TotMEncWDFW(Fish, TStep)
                  CellAdd = "C" & R + 3
                  shXL.Range(CellAdd).Value = TotMSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)
                  CellAdd = "C" & R + 4
                  shXL.Range(CellAdd).Value = TotMEncWDFW(Fish, TStep) + TotMSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)

                  CellAdd = "D" & R + 1
                  shXL.Range(CellAdd).Value = "Unmark"
                  CellAdd = "D" & R + 2
                  shXL.Range(CellAdd).Value = TotUEncWDFW(Fish, TStep)
                  CellAdd = "D" & R + 3
                  shXL.Range(CellAdd).Value = TotUSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)
                  CellAdd = "D" & R + 4
                  shXL.Range(CellAdd).Value = TotUEncWDFW(Fish, TStep) + TotUSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)

                  CellAdd = "E" & R + 1
                  shXL.Range(CellAdd).Value = "Total"
                  CellAdd = "E" & R + 2
                  shXL.Range(CellAdd).Value = TotUEncWDFW(Fish, TStep) + TotMEncWDFW(Fish, TStep)
                  CellAdd = "E" & R + 3
                  shXL.Range(CellAdd).Value = TotUSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep) + TotMSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)
                  CellAdd = "E" & R + 4
                  shXL.Range(CellAdd).Value = TotUEncWDFW(Fish, TStep) + TotUSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep) + TotMEncWDFW(Fish, TStep) + TotMSubWDFW(Fish, TStep) / ShakerMortRate(Fish, TStep)


                  'Total Mortality
                  CellAdd = "G" & R
                  shXL.Range(CellAdd).Value = "Total Mortality"
                  CellAdd = "G" & R + 1
                  shXL.Range(CellAdd).Value = "Mark"
                  CellAdd = "G" & R + 2
                  shXL.Range(CellAdd).Value = TotMCatWDFW(Fish, TStep) + TotMNonWDFW(Fish, TStep) + TotMShkWDFW(Fish, TStep)
                  CellAdd = "G" & R + 3
                  shXL.Range(CellAdd).Value = TotMSubWDFW(Fish, TStep)
                  CellAdd = "G" & R + 4
                  shXL.Range(CellAdd).Value = TotMSubWDFW(Fish, TStep) + TotMCatWDFW(Fish, TStep) + TotMShkWDFW(Fish, TStep) + TotMNonWDFW(Fish, TStep)

                  CellAdd = "H" & R + 1
                  shXL.Range(CellAdd).Value = "Unmark"
                  CellAdd = "H" & R + 2
                  shXL.Range(CellAdd).Value = TotUCatWDFW(Fish, TStep) + TotUNonWDFW(Fish, TStep) + TotUShkWDFW(Fish, TStep)
                  CellAdd = "H" & R + 3
                  shXL.Range(CellAdd).Value = TotUSubWDFW(Fish, TStep)
                  CellAdd = "H" & R + 4
                  shXL.Range(CellAdd).Value = TotUSubWDFW(Fish, TStep) + TotUCatWDFW(Fish, TStep) + TotUShkWDFW(Fish, TStep) + TotUNonWDFW(Fish, TStep)

                  CellAdd = "I" & R + 1
                  shXL.Range(CellAdd).Value = "Total"
                  CellAdd = "I" & R + 2
                  shXL.Range(CellAdd).Value = TotMCatWDFW(Fish, TStep) + TotMNonWDFW(Fish, TStep) + TotMShkWDFW(Fish, TStep) + TotUCatWDFW(Fish, TStep) + TotUNonWDFW(Fish, TStep) + TotUShkWDFW(Fish, TStep)
                  CellAdd = "I" & R + 3
                  shXL.Range(CellAdd).Value = TotMSubWDFW(Fish, TStep) + TotUSubWDFW(Fish, TStep)
                  CellAdd = "I" & R + 4
                  shXL.Range(CellAdd).Value = TotMSubWDFW(Fish, TStep) + TotMCatWDFW(Fish, TStep) + TotMShkWDFW(Fish, TStep) + TotMNonWDFW(Fish, TStep) + TotUSubWDFW(Fish, TStep) + TotUCatWDFW(Fish, TStep) + TotUShkWDFW(Fish, TStep) + TotUNonWDFW(Fish, TStep)

                  'Total Landed
                  CellAdd = "K" & R
                  shXL.Range(CellAdd).Value = "Total Landed"
                  CellAdd = "K" & R + 1
                  shXL.Range(CellAdd).Value = "Mark"
                  CellAdd = "K" & R + 2
                  shXL.Range(CellAdd).Value = TotMCatWDFW(Fish, TStep)
                  CellAdd = "K" & R + 3
                  shXL.Range(CellAdd).Value = 0
                  CellAdd = "K" & R + 4
                  shXL.Range(CellAdd).Value = TotMCatWDFW(Fish, TStep)

                  CellAdd = "L" & R + 1
                  shXL.Range(CellAdd).Value = "Unmark"
                  CellAdd = "L" & R + 2
                  shXL.Range(CellAdd).Value = TotUCatWDFW(Fish, TStep)
                  CellAdd = "L" & R + 3
                  shXL.Range(CellAdd).Value = 0
                  CellAdd = "L" & R + 4
                  shXL.Range(CellAdd).Value = TotUCatWDFW(Fish, TStep)

                  CellAdd = "M" & R + 1
                  shXL.Range(CellAdd).Value = "Total"
                  CellAdd = "M" & R + 2
                  shXL.Range(CellAdd).Value = TotUCatWDFW(Fish, TStep) + TotMCatWDFW(Fish, TStep)
                  CellAdd = "M" & R + 3
                  shXL.Range(CellAdd).Value = 0
                  CellAdd = "M" & R + 4
                  shXL.Range(CellAdd).Value = TotUCatWDFW(Fish, TStep) + TotMCatWDFW(Fish, TStep)
                  R = R + 6
               End If
            Next
         Next


         ' Make sure Excel is visible and give the user control
         ' of Excel's lifetime.
         appXL.Visible = True
         appXL.UserControl = True
         ' Release object references.
         raXL = Nothing
         shXL = Nothing
         wbXl = Nothing
         'appXL.Quit()
         appXL = Nothing

      End If

   End Sub
End Class