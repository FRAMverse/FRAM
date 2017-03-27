Public Class FVS_FishStkCompScreen
   Public NumContribStk As Integer
   Private Sub FVS_SelectiveFisheryScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 966
      FormWidth = 1090
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
         If FVS_FishStkCompScreen_ReSize = False Then
            Resize_Form(Me)
            FVS_FishStkCompScreen_ReSize = True
         End If
      End If

      FSCComboBox.Items.Clear()
      For Fish As Integer = 1 To NumFish
         FSCComboBox.Items.Add(FisheryTitle(Fish))
      Next
      FSCSelectedLabel.Text = "Selected-Fishery"
      FSCGrid.Columns.Clear()
   End Sub

   Private Sub FSCComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSCComboBox.SelectedIndexChanged
      Dim TempVal, StkTempVal As Double
      Dim TotFisheryMort(NumSteps + 2), StkAgeTempVal, StkTempVal24 As Double

      Fish = FSCComboBox.SelectedIndex + 1
      '- Sum Total Mortality by Fishery and Time Step
      ReDim TotFisheryMort(NumSteps + 2)
      NumContribStk = 0
      For Stk As Integer = 1 To NumStk
         StkTempVal = 0
         For Age As Integer = MinAge To MaxAge
            StkAgeTempVal = 0
            For TStep As Integer = 1 To NumSteps
               TempVal = LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
               TotFisheryMort(TStep) += TempVal
               TotFisheryMort(NumSteps + 1) += TempVal
               If SpeciesName = "CHINOOK" And TStep > 1 Then TotFisheryMort(NumSteps + 2) += TempVal
               StkTempVal += TempVal
               StkAgeTempVal += TempVal
            Next
         Next
         If StkTempVal > 0 Then NumContribStk += 1
      Next
      If NumContribStk = 0 Then NumContribStk = 1
      FSCSelectedLabel.Text = FisheryTitle(Fish)
      FSCGrid.Columns.Clear()
      FSCGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      FSCGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         FSCGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         FSCGrid.Columns.Add("Name", "StockName")
         FSCGrid.Columns(0).Width = 400 / FormWidthScaler
         FSCGrid.Columns(0).ReadOnly = True
         FSCGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         FSCGrid.Columns.Add("T1", TimeStepName(1))
         FSCGrid.Columns(1).Width = 90 / FormWidthScaler
         FSCGrid.Columns(1).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T2", TimeStepName(2))
         FSCGrid.Columns(2).Width = 90 / FormWidthScaler
         FSCGrid.Columns(2).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T3", TimeStepName(3))
         FSCGrid.Columns(3).Width = 90 / FormWidthScaler
         FSCGrid.Columns(3).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T4", TimeStepName(4))
         FSCGrid.Columns(4).Width = 90 / FormWidthScaler
         FSCGrid.Columns(4).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T5", "Time 2-4")
         FSCGrid.Columns(5).Width = 90 / FormWidthScaler
         FSCGrid.Columns(5).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T6", "Total")
         FSCGrid.Columns(6).Width = 90 / FormWidthScaler
         FSCGrid.Columns(6).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.RowCount = NumContribStk
      ElseIf SpeciesName = "COHO" Then
         FSCGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         FSCGrid.Columns.Add("Name", "StockName")
         FSCGrid.Columns(0).Width = 400 / FormWidthScaler
         FSCGrid.Columns(0).ReadOnly = True
         FSCGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         FSCGrid.Columns.Add("T1", TimeStepName(1))
         FSCGrid.Columns(1).Width = 90 / FormWidthScaler
         FSCGrid.Columns(1).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T2", TimeStepName(2))
         FSCGrid.Columns(2).Width = 90 / FormWidthScaler
         FSCGrid.Columns(2).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T3", TimeStepName(3))
         FSCGrid.Columns(3).Width = 90 / FormWidthScaler
         FSCGrid.Columns(3).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T4", TimeStepName(4))
         FSCGrid.Columns(4).Width = 90 / FormWidthScaler
         FSCGrid.Columns(4).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T5", TimeStepName(5))
         FSCGrid.Columns(5).Width = 90 / FormWidthScaler
         FSCGrid.Columns(5).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.Columns.Add("T6", "Total")
         FSCGrid.Columns(6).Width = 90 / FormWidthScaler
         FSCGrid.Columns(6).DefaultCellStyle.Format = ("#####0.00")
         FSCGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FSCGrid.RowCount = NumContribStk
      End If

      '- Stock Composition Lines
      NumContribStk = 0
      For Stk As Integer = 1 To NumStk
         StkTempVal = 0
         StkTempVal24 = 0
         '- Check if Stock Contributes to this Fishery
         For TStep As Integer = 1 To NumSteps
            For Age As Integer = MinAge To MaxAge
               TempVal = LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
               StkTempVal += TempVal
               If SpeciesName = "CHINOOK" And TStep > 1 Then StkTempVal24 += TempVal
            Next
         Next
         '- If Yes, Put Data into Grid by Time Step
         If StkTempVal > 0 Then
            FSCGrid.Item(0, NumContribStk).Value = StockTitle(Stk)
            For TStep As Integer = 1 To NumSteps
               If TotFisheryMort(TStep) = 0 Then
                  FSCGrid.Item(TStep, NumContribStk).Value = "****"
               Else
                  TempVal = 0
                  For Age As Integer = MinAge To MaxAge
                     TempVal += LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                  Next
                  FSCGrid.Item(TStep, NumContribStk).Value = (TempVal / TotFisheryMort(TStep) * 100).ToString("##0.00")
               End If
            Next
            If SpeciesName = "CHINOOK" Then
               If TotFisheryMort(NumSteps + 2) = 0 Then
                  FSCGrid.Item(NumSteps + 1, NumContribStk).Value = "****"
               Else
                  FSCGrid.Item(NumSteps + 1, NumContribStk).Value = ((StkTempVal24 / TotFisheryMort(NumSteps + 2)) * 100).ToString("##0.00")
               End If
               FSCGrid.Item(NumSteps + 2, NumContribStk).Value = ((StkTempVal / TotFisheryMort(NumSteps + 1)) * 100).ToString("##0.00")
            ElseIf SpeciesName = "COHO" Then
               FSCGrid.Item(NumSteps + 1, NumContribStk).Value = ((StkTempVal / TotFisheryMort(NumSteps + 1)) * 100).ToString("##0.00")
            End If
            NumContribStk += 1
         End If
      Next

   End Sub

   Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
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
      '- Column Headings
      If SpeciesName = "CHINOOK" Then
         ClipStr &= "StockName" & vbTab & TimeStepName(1) & vbTab & TimeStepName(2) & vbTab & TimeStepName(3) & vbTab & TimeStepName(4) & vbTab & "Time 2-4" & vbTab & "Total" & vbCr
      ElseIf SpeciesName = "COHO" Then
         ClipStr &= "StockName" & vbTab & TimeStepName(1) & vbTab & TimeStepName(2) & vbTab & TimeStepName(3) & vbTab & TimeStepName(4) & vbTab & TimeStepName(5) & vbTab & "Total" & vbCr
      End If
      '- Grid Fields
      For RecNum = 0 To NumContribStk - 1
         For ColNum = 0 To 6
            If ColNum = 0 Then
               ClipStr &= FSCGrid.Item(ColNum, RecNum).Value
            Else
               If FSCGrid.Item(ColNum, RecNum).Value = "****" Then
                  ClipStr &= vbTab & "****"
               Else
                  ClipStr &= vbTab & CDbl(FSCGrid.Item(ColNum, RecNum).Value)
               End If
            End If
         Next
         ClipStr &= vbCr
      Next
      Clipboard.SetDataObject(ClipStr)

   End Sub

   Private Sub FSCExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSCExitButton.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
   End Sub

End Class