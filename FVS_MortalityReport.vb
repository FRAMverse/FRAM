Public Class FVS_MortalityReport
   Public AgeSumState As Boolean
   Private Sub FVS_MortalityReport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

      'FormHeight = 962
      FormHeight = 972
      FormWidth = 977
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
         If FVS_MortalityReport_ReSize = False Then
            Resize_Form(Me)
            FVS_MortalityReport_ReSize = True
         End If
      End If

      If SpeciesName = "CHINOOK" Then
         AgeSumButton.Visible = True
         AgeSumButton.Text = "Sum Age Only"
         MortalityTypeComboBox.Items.Clear()
         MortalityTypeComboBox.Items.Add("NoneSelected")
         MortalityTypeComboBox.Items.Add("Landed Catch")
         MortalityTypeComboBox.Items.Add("NonRetention")
         MortalityTypeComboBox.Items.Add("Shakers")
         MortalityTypeComboBox.Items.Add("DropOff")
         MortalityTypeComboBox.Items.Add("TotalMortality")
         MortalityTypeComboBox.Items.Add("AEQ-Total Mort")
      Else
         AgeSumButton.Visible = False
         MortalityTypeComboBox.Items.Clear()
         MortalityTypeComboBox.Items.Add("NoneSelected")
         MortalityTypeComboBox.Items.Add("Landed Catch")
         MortalityTypeComboBox.Items.Add("NonRetention")
         MortalityTypeComboBox.Items.Add("Shakers")
         MortalityTypeComboBox.Items.Add("DropOff")
         MortalityTypeComboBox.Items.Add("TotalMortality")
      End If
      AgeSumState = False
      MortalityType = 1
      Call LoadGridValues()

   End Sub

   Sub LoadGridValues()

      'MortalityType = 0
      'MortalityTypeComboBox.Text = "Landed Catch"
      'MortalityType = 1

      If SpeciesName = "CHINOOK" Then
         Select Case MortalityType
            Case 0
               MortalityTypeComboBox.Text = "NoneSelected"
            Case 1
               MortalityTypeComboBox.Text = "Landed Catch"
            Case 2
               MortalityTypeComboBox.Text = "NonRetention"
            Case 3
               MortalityTypeComboBox.Text = "Shakers"
            Case 4
               MortalityTypeComboBox.Text = "DropOff"
            Case 5
               MortalityTypeComboBox.Text = "TotalMortality"
            Case 6
               MortalityTypeComboBox.Text = "AEQ-TotalMort"
         End Select
      ElseIf SpeciesName = "COHO" Then
         Select Case MortalityType
            Case 0
               MortalityTypeComboBox.Text = "NoneSelected"
            Case 1
               MortalityTypeComboBox.Text = "Landed Catch"
            Case 2
               MortalityTypeComboBox.Text = "NonRetention"
            Case 3
               MortalityTypeComboBox.Text = "Shakers"
            Case 4
               MortalityTypeComboBox.Text = "DropOff"
            Case 5
               MortalityTypeComboBox.Text = "TotalMortality"
         End Select
      End If

      If ScreenReportType = 1 Then
         StockTitleLabel.Visible = False
         StocksSelectedLabel.Visible = False
         Me.Text = "Fishery Mortality Report"
      ElseIf ScreenReportType = 2 Then
         StockTitleLabel.Visible = True
         StocksSelectedLabel.Visible = True
         StocksSelectedLabel.Text = ""
         If NumSelectedStocks > 2 Then
            StocksSelectedLabel.Font = New Font("Microsoft San Serif", 8, FontStyle.Bold)
         Else
            StocksSelectedLabel.Font = New Font("Microsoft San Serif", 10, FontStyle.Bold)
         End If
         For Stk As Integer = 1 To NumSelectedStocks
            StocksSelectedLabel.Text &= StockTitle(StockSelection(Stk)) & ","
            If Stk = 2 And NumSelectedStocks > 2 Then
               StocksSelectedLabel.Text &= "--Plus More--"
               Exit For
            End If
         Next
         Me.Text = "Stock Catch Mortality Report"
      End If

      MortalityGrid.Columns.Clear()
      MortalityGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      MortalityGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         MortalityGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         MortalityGrid.Columns.Add("Name", "FisheryName")
         MortalityGrid.Columns(0).Width = 175 / FormWidthScaler
         MortalityGrid.Columns(0).ReadOnly = True
         MortalityGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         MortalityGrid.Columns.Add("Age", "Age")
         MortalityGrid.Columns(1).Width = 50 / FormWidthScaler
         MortalityGrid.Columns(1).DefaultCellStyle.Format = ("0")
         MortalityGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T1", "Oct-Apr1")
         MortalityGrid.Columns(2).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T2", "May-June")
         MortalityGrid.Columns(3).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T3", "July-Sept")
         MortalityGrid.Columns(4).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T4", "Oct-Apr2")
         MortalityGrid.Columns(5).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("Total", "Total")
         MortalityGrid.Columns(6).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.RowCount = NumFish * MaxAge + 1
      ElseIf SpeciesName = "COHO" Then
         MortalityGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         MortalityGrid.Columns.Add("Name", "FisheryName")
         MortalityGrid.Columns(0).Width = 175 / FormWidthScaler
         MortalityGrid.Columns(0).ReadOnly = True
         MortalityGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         MortalityGrid.Columns.Add("T1", "Jan-June")
         MortalityGrid.Columns(1).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(1).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T2", "July")
         MortalityGrid.Columns(2).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T3", "August")
         MortalityGrid.Columns(3).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T4", "September")
         MortalityGrid.Columns(4).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T5", "Oct-Dec")
         MortalityGrid.Columns(5).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("Total", "Total")
         MortalityGrid.Columns(6).Width = 100 / FormWidthScaler
         MortalityGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.RowCount = NumFish
      End If
      Call FillMortalityGrid()
   End Sub

   Sub FillMortalityGrid()

      Dim RowNum, SelStk As Integer
      If SpeciesName = "CHINOOK" Then
         Dim TotCatch(MaxAge, NumFish, NumSteps + 1)
         Select Case MortalityType
            Case 1
               Me.Text = "Landed Catch Report"
               Me.MortalityReportTitle.Text = "CHINOOK Landed Catch"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk1
                     Next
                     GoTo SkipChinStk1
                  End If
SumChinStk1:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Age, Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                           TotCatch(Age, Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                        Next
                     Next
                  Next
SkipChinStk1:
               Next
            Case 2
               Me.Text = "NonRetention Report"
               Me.MortalityReportTitle.Text = "CHINOOK NonRetention"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk2
                     Next
                     GoTo SkipChinStk2
                  End If
SumChinStk2:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Age, Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(Age, Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        Next
                     Next
                  Next
SkipChinStk2:
               Next
            Case 3
               Me.Text = "Sub-Legal Shaker Report"
               Me.MortalityReportTitle.Text = "CHINOOK Sub-Legal Shakers"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk3
                     Next
                     GoTo SkipChinStk3
                  End If
SumChinStk3:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Age, Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(Age, Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        Next
                     Next
                  Next
SkipChinStk3:
               Next
            Case 4
               Me.Text = "DropOff Report"
               Me.MortalityReportTitle.Text = "CHINOOK DropOff"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk4
                     Next
                     GoTo SkipChinStk4
                  End If
SumChinStk4:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Age, Fish, TStep) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                           TotCatch(Age, Fish, NumSteps + 1) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, TStep) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, NumSteps + 1) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        Next
                     Next
                  Next
SkipChinStk4:
               Next
            Case 5
               Me.Text = "Total Mortality Report"
               Me.MortalityReportTitle.Text = "CHINOOK Total Mortality"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk5
                     Next
                     GoTo SkipChinStk5
                  End If
SumChinStk5:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           TotCatch(Age, Fish, TStep) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(Age, Fish, NumSteps + 1) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, TStep) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(1, Fish, NumSteps + 1) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        Next
                     Next
                  Next
SkipChinStk5:
               Next
            Case 6
               Me.Text = "AEQ-Total Mortality Report"
               Me.MortalityReportTitle.Text = "CHINOOK AEQ Total Mortality"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumChinStk6
                     Next
                     GoTo SkipChinStk6
                  End If
SumChinStk6:
                  For Age As Integer = MinAge To MaxAge
                     For Fish As Integer = 1 To NumFish
                        For TStep As Integer = 1 To NumSteps
                           If TerminalFisheryFlag(Fish, TStep) = Term Then
                              TotCatch(Age, Fish, TStep) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                              TotCatch(Age, Fish, NumSteps + 1) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                              TotCatch(1, Fish, TStep) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                              TotCatch(1, Fish, NumSteps + 1) += +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           Else
                              TotCatch(Age, Fish, TStep) += (+LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                              TotCatch(Age, Fish, NumSteps + 1) += (+LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                              TotCatch(1, Fish, TStep) += (+LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                              TotCatch(1, Fish, NumSteps + 1) += (+LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                           End If
                        Next
                     Next
                  Next
SkipChinStk6:
               Next
         End Select
         '- Put CHINOOK Total Array into Grid
         For Fish As Integer = 1 To NumFish
            MortalityGrid.Item(0, ((Fish - 1) * MaxAge)).Value = FisheryName(Fish)
            For Age As Integer = 1 To MaxAge
               If Age = 1 Then
                  RowNum = (Fish - 1) * MaxAge
                  MortalityGrid.Item(1, RowNum).Value = "Sum"
               Else
                  RowNum = ((Fish - 1) * MaxAge) + Age - 1
                  MortalityGrid.Item(1, RowNum).Value = Age
               End If
               For TStep As Integer = 1 To NumSteps + 1
                  If ModelStockProportion(Fish) <> 0 And ScreenReportType = 1 Then
                     MortalityGrid.Item(TStep + 1, RowNum).Value = TotCatch(Age, Fish, TStep) / ModelStockProportion(Fish)
                  ElseIf ScreenReportType = 2 Then
                     MortalityGrid.Item(TStep + 1, RowNum).Value = TotCatch(Age, Fish, TStep)
                  Else
                     MortalityGrid.Item(TStep + 1, RowNum).Value = "-"
                  End If
               Next
            Next
         Next

      ElseIf SpeciesName = "COHO" Then

         Dim TotCatch(NumFish, NumSteps + 1)

         Age = 3
         Select Case MortalityType
            Case 1
               Me.Text = "Landed Catch Report"
               Me.MortalityReportTitle.Text = "COHO Landed Catch"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumSelStk1
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
               Me.Text = "NonRetention Report"
               Me.MortalityReportTitle.Text = "COHO NonRetention"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumSelStk2
                     Next
                     GoTo SkipSelStk2
                  End If
SumSelStk2:
                  For Fish As Integer = 1 To NumFish
                     For TStep As Integer = 1 To NumSteps
                        TotCatch(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                     Next
                  Next
SkipSelStk2:
               Next
            Case 3
               Me.Text = "Sub-Legal Shaker Report"
               Me.MortalityReportTitle.Text = "COHO Shakers"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumSelStk3
                     Next
                     GoTo SkipSelStk3
                  End If
SumSelStk3:
                  For Fish As Integer = 1 To NumFish
                     For TStep As Integer = 1 To NumSteps
                        TotCatch(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     Next
                  Next
SkipSelStk3:
               Next
            Case 4
               Me.Text = "DropOff Report"
               Me.MortalityReportTitle.Text = "COHO DropOff"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumSelStk4
                     Next
                     GoTo SkipSelStk4
                  End If
SumSelStk4:
                  For Fish As Integer = 1 To NumFish
                     For TStep As Integer = 1 To NumSteps
                        TotCatch(Fish, TStep) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                     Next
                  Next
SkipSelStk4:
               Next
            Case 5
               Me.Text = "Total Mortality Report"
               Me.MortalityReportTitle.Text = "COHO Total Mortality"
               For Stk As Integer = 1 To NumStk
                  If ScreenReportType = 2 Then
                     For SelStk = 1 To NumSelectedStocks
                        If Stk = StockSelection(SelStk) Then GoTo SumSelStk5
                     Next
                     GoTo SkipSelStk5
                  End If
SumSelStk5:
                  For Fish As Integer = 1 To NumFish
                     For TStep As Integer = 1 To NumSteps
                        TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     Next
                  Next
SkipSelStk5:
               Next
         End Select
         '- Put COHO Total Array into Grid
         For Fish As Integer = 1 To NumFish
            MortalityGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
            For TStep As Integer = 1 To NumSteps + 1
               MortalityGrid.Item(TStep, Fish - 1).Value = TotCatch(Fish, TStep)
            Next
         Next
      End If


   End Sub

   Private Sub CRCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CRCancelButton.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
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
      If ScreenReportType = 1 Then
         ClipStr &= "FISHERY "
      ElseIf ScreenReportType = 2 Then
         ClipStr &= "STOCK "
      End If
      Select Case MortalityType
         Case 1
            ClipStr &= "Landed Catch Mortality Report"
         Case 2
            ClipStr &= "Non-Retention Mortality Report"
         Case 3
            ClipStr &= "Shaker Mortality Report"
         Case 4
            ClipStr &= "DropOff Mortality Report"
         Case 5
            ClipStr &= "Total Mortality Report"
         Case 6
            ClipStr &= "AEQ-Total Mortality Report"
      End Select
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr
      If ScreenReportType = 2 Then
         ClipStr &= "STOCKS="
         For Stk As Integer = 1 To NumSelectedStocks
            ClipStr &= StockTitle(StockSelection(Stk)) & ","
         Next
         ClipStr &= vbCr
      End If

      If SpeciesName = "CHINOOK" Then
         If AgeSumState = False Then
            ClipStr &= "FisheryName" & vbTab & "Age" & vbTab & "Oct-Apr1" & vbTab & "May-June" & vbTab & "July-Sept" & vbTab & "Oct-Apr2" & vbTab & "Total" & vbCr
            For RecNum = 0 To ((NumFish * MaxAge) - 1)
               For ColNum = 0 To NumSteps + 2
                  If ColNum = 0 Then
                     If ((RecNum) Mod 5) = 0 Then
                        ClipStr &= MortalityGrid.Item(ColNum, RecNum).Value
                        'Else
                        '   ClipStr &= vbTab
                     End If
                  ElseIf ColNum = 1 Then
                     ClipStr &= vbTab & MortalityGrid.Item(ColNum, RecNum).Value
                  Else
                     ClipStr &= vbTab & CInt(MortalityGrid.Item(ColNum, RecNum).Value)
                  End If
               Next
               ClipStr &= vbCr
            Next
         ElseIf AgeSumState = True Then
            ClipStr &= "FisheryName" & vbTab & "Oct-Apr1" & vbTab & "May-June" & vbTab & "July-Sept" & vbTab & "Oct-Apr2" & vbTab & "Total" & vbCr
            For RecNum = 0 To NumFish - 1
               For ColNum = 0 To NumSteps + 1
                  If ColNum = 0 Then
                     ClipStr = ClipStr & MortalityGrid.Item(ColNum, RecNum).Value
                  Else
                     ClipStr = ClipStr & vbTab & CInt(MortalityGrid.Item(ColNum, RecNum).Value)
                  End If
               Next
               ClipStr &= vbCr
            Next
         End If

      ElseIf SpeciesName = "COHO" Then
         ClipStr &= "FisheryName" & vbTab & "Jan-June" & vbTab & "July" & vbTab & "August" & vbTab & "September" & vbTab & "Oct-Dec" & vbTab & "Total" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To NumSteps + 1
               If ColNum = 0 Then
                  ClipStr = ClipStr & MortalityGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & CInt(MortalityGrid.Item(ColNum, RecNum).Value)
               End If
            Next
            ClipStr &= vbCr
         Next
      End If

      Clipboard.SetDataObject(ClipStr)

   End Sub

   Private Sub MortalityTypeComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MortalityTypeComboBox.SelectedIndexChanged
      '- Exit for Form Load
      If MortalityType = 0 Then Exit Sub
      MortalityType = MortalityTypeComboBox.SelectedIndex
      '- Change Mortality Type Back to One if "None Selected" (Index=0)
      If MortalityType = 0 Then MortalityType = 1
      If AgeSumState = False Then
         Call LoadGridValues()
         'Call FillMortalityGrid()
      Else
         Call LoadNoAgeGridValues()
         'Call FillMortalityGrid()
      End If
   End Sub

   Private Sub AgeSumButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AgeSumButton.Click
      If AgeSumState = False Then
         AgeSumState = True
         AgeSumButton.Text = "Show Mort/Age"
         Call LoadNoAgeGridValues()
      Else
         AgeSumState = False
         AgeSumButton.Text = "Sum Age Only"
         Call LoadGridValues()
      End If

   End Sub

   Sub LoadNoAgeGridValues()
      Dim SelStk As Integer

      'MortalityType = 0
      'MortalityTypeComboBox.Text = "Landed Catch"
      'MortalityType = 1

      Select Case MortalityType
         Case 0
            MortalityTypeComboBox.Text = "NoneSelected"
         Case 1
            MortalityTypeComboBox.Text = "Landed Catch"
         Case 2
            MortalityTypeComboBox.Text = "NonRetention"
         Case 3
            MortalityTypeComboBox.Text = "Shakers"
         Case 4
            MortalityTypeComboBox.Text = "DropOff"
         Case 5
            MortalityTypeComboBox.Text = "TotalMortality"
         Case 5
            MortalityTypeComboBox.Text = "AEQ-TotalMortality"
      End Select

      If ScreenReportType = 1 Then
         StockTitleLabel.Visible = False
         StocksSelectedLabel.Visible = False
         Me.Text = "Fishery Mortality Report"
      ElseIf ScreenReportType = 2 Then
         StockTitleLabel.Visible = True
         StocksSelectedLabel.Visible = True
         StocksSelectedLabel.Text = ""
         If NumSelectedStocks > 2 Then
            StocksSelectedLabel.Font = New Font("Microsoft San Serif", 8, FontStyle.Bold)
         Else
            StocksSelectedLabel.Font = New Font("Microsoft San Serif", 10, FontStyle.Bold)
         End If
         For Stk As Integer = 1 To NumSelectedStocks
            StocksSelectedLabel.Text &= StockTitle(StockSelection(Stk)) & ","
            If Stk = 2 And NumSelectedStocks > 2 Then
               StocksSelectedLabel.Text &= "--Plus More--"
               Exit For
            End If
         Next
         Me.Text = "Stock Catch Mortality Report"
      End If

      MortalityGrid.Columns.Clear()
      MortalityGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      MortalityGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         MortalityGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         MortalityGrid.Columns.Add("Name", "FisheryName")
         MortalityGrid.Columns(0).Width = 175
         MortalityGrid.Columns(0).ReadOnly = True
         MortalityGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         MortalityGrid.Columns.Add("T1", "Oct-Apr1")
         MortalityGrid.Columns(1).Width = 100
         MortalityGrid.Columns(1).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T2", "May-June")
         MortalityGrid.Columns(2).Width = 100
         MortalityGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T3", "July-Sept")
         MortalityGrid.Columns(3).Width = 100
         MortalityGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("T4", "Oct-Apr2")
         MortalityGrid.Columns(4).Width = 100
         MortalityGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.Columns.Add("Total", "Total")
         MortalityGrid.Columns(5).Width = 100
         MortalityGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         MortalityGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         MortalityGrid.RowCount = NumFish
      End If

      '- Put CHINOOK Summed Age Data into Grid like COHO
      Dim TotCatch(NumFish, NumSteps + 1)
      Select Case MortalityType
         Case 1
            Me.Text = "Landed Catch Report"
            Me.MortalityReportTitle.Text = "CHINOOK Landed Catch"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk1
                  Next
                  GoTo SkipSelStk1
               End If
SumSelStk1:
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     For Age As Integer = MinAge To MaxAge
                        TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                     Next
                  Next
               Next
SkipSelStk1:
            Next
         Case 2
            Me.Text = "NonRetention Report"
            Me.MortalityReportTitle.Text = "CHINOOK NonRetention"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk2
                  Next
                  GoTo SkipSelStk2
               End If
SumSelStk2:
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     For Age As Integer = MinAge To MaxAge
                        TotCatch(Fish, TStep) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += NonRetention(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep)
                     Next
                  Next
               Next
SkipSelStk2:
            Next
         Case 3
            Me.Text = "Sub-Legal Shaker Report"
            Me.MortalityReportTitle.Text = "CHINOOK Shakers"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk3
                  Next
                  GoTo SkipSelStk3
               End If
SumSelStk3:
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     For Age As Integer = MinAge To MaxAge
                        TotCatch(Fish, TStep) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += Shakers(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     Next
                  Next
               Next
SkipSelStk3:
            Next
         Case 4
            Me.Text = "DropOff Report"
            Me.MortalityReportTitle.Text = "CHINOOK DropOff"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk4
                  Next
                  GoTo SkipSelStk4
               End If
SumSelStk4:
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     For Age As Integer = MinAge To MaxAge
                        TotCatch(Fish, TStep) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += DropOff(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)
                     Next
                  Next
               Next
SkipSelStk4:
            Next
         Case 5
            Me.Text = "Total Mortality Report"
            Me.MortalityReportTitle.Text = "CHINOOK Total Mortality"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk5
                  Next
                  GoTo SkipSelStk5
               End If
SumSelStk5:
               For Fish As Integer = 1 To NumFish
                  For TStep As Integer = 1 To NumSteps
                     For Age As Integer = MinAge To MaxAge
                        TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     Next
                  Next
               Next
SkipSelStk5:
            Next
         Case 6
            Me.Text = "AEQ-Total Mortality Report"
            Me.MortalityReportTitle.Text = "CHINOOK AEQ Total Mortality"
            For Stk As Integer = 1 To NumStk
               If ScreenReportType = 2 Then
                  For SelStk = 1 To NumSelectedStocks
                     If Stk = StockSelection(SelStk) Then GoTo SumSelStk6
                  Next
                  GoTo SkipSelStk6
               End If
SumSelStk6:
               For Age As Integer = MinAge To MaxAge
                  For Fish As Integer = 1 To NumFish
                     For TStep As Integer = 1 To NumSteps
                        If TerminalFisheryFlag(Fish, TStep) = Term Then
                           TotCatch(Fish, TStep) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                           TotCatch(Fish, NumSteps + 1) += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                        Else
                           TotCatch(Fish, TStep) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                           TotCatch(Fish, NumSteps + 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        End If
                     Next
                  Next
               Next
SkipSelStk6:
            Next
      End Select

      '- Put CHINOOK Total Array Summed across Age into Grid
      For Fish As Integer = 1 To NumFish
         MortalityGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
         For TStep As Integer = 1 To NumSteps + 1
            If ModelStockProportion(Fish) <> 0 And ScreenReportType = 1 Then
               MortalityGrid.Item(TStep, Fish - 1).Value = TotCatch(Fish, TStep) / ModelStockProportion(Fish)
            ElseIf ScreenReportType = 2 Then
               MortalityGrid.Item(TStep, Fish - 1).Value = TotCatch(Fish, TStep)
            Else
               MortalityGrid.Item(TStep, Fish - 1).Value = "-"
            End If
         Next
      Next

   End Sub


   Private Sub MortalityGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles MortalityGrid.CellContentClick

   End Sub
End Class