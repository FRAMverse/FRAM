Public Class FVS_StockImpactsPer1000Screen

   Private Sub FVS_StockImpactsPer1000Screen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 903
      FormHeight = 923
      FormWidth = 995
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
         If FVS_StockImpactsPer1000Screen_ReSize = False Then
            Resize_Form(Me)
            FVS_StockImpactsPer1000Screen_ReSize = True
         End If
      End If

      SIPComboBox.SelectedIndex = -1
      SIPComboBox.Items.Clear()
      For Stk As Integer = 1 To NumStk
         SIPComboBox.Items.Add(StockTitle(Stk))
      Next
      SIPSelectedLabel.Text = "Selected-Stock"
   End Sub

   Private Sub SIPExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SIPExitButton.Click
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
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr
      If SpeciesName = "CHINOOK" Then
         ClipStr &= "FisheryName" & vbTab & "Age" & vbTab & "Oct-Apr1" & vbTab & "May-June" & vbTab & "July-Sept" & vbTab & "Oct-Apr2" & vbTab & "Total" & vbCr
         For RecNum = 0 To SIPGrid.RowCount - 1
            For ColNum = 0 To SIPGrid.ColumnCount - 1
               If IsNumeric(SIPGrid.Item(ColNum, RecNum).Value) Then
                  ClipStr &= vbTab & CDbl(SIPGrid.Item(ColNum, RecNum).Value)
               Else
                  If ColNum = 0 Then
                     ClipStr &= SIPGrid.Item(ColNum, RecNum).Value
                  Else
                     ClipStr &= vbTab & SIPGrid.Item(ColNum, RecNum).Value
                  End If
               End If
            Next
            ClipStr &= vbCr
         Next
      ElseIf SpeciesName = "COHO" Then
         ClipStr &= "FisheryName" & vbTab & "Jan-June" & vbTab & "July" & vbTab & "August" & vbTab & "September" & vbTab & "Oct-Dec" & vbTab & "Total" & vbCr
         For RecNum = 0 To SIPGrid.RowCount - 1
            For ColNum = 0 To SIPGrid.ColumnCount - 1
               If ColNum = 0 Then
                  ClipStr &= SIPGrid.Item(ColNum, RecNum).Value
               Else
                  If IsNumeric(SIPGrid.Item(ColNum, RecNum).Value) Then
                     ClipStr &= vbTab & CDbl(SIPGrid.Item(ColNum, RecNum).Value)
                  Else
                     ClipStr &= vbTab & SIPGrid.Item(ColNum, RecNum).Value
                  End If
               End If
            Next
            ClipStr &= vbCr
         Next
      End If
      Clipboard.SetDataObject(ClipStr)

   End Sub

   Private Sub SIPComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SIPComboBox.SelectedIndexChanged
      Dim TempVal, FisherySum, FisheryStepSum, FisheryTotalSum As Double

      Stk = SIPComboBox.SelectedIndex + 1
      SIPSelectedLabel.Text = StockTitle(Stk)

      SIPGrid.Columns.Clear()
      SIPGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      SIPGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         SIPGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         SIPGrid.Columns.Add("Name", "FisheryName")
         SIPGrid.Columns(0).Width = 250 / FormWidthScaler
         SIPGrid.Columns(0).ReadOnly = True
         SIPGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         SIPGrid.Columns.Add("Age", "Age")
         SIPGrid.Columns(1).Width = 40 / FormWidthScaler
         SIPGrid.Columns(1).ReadOnly = True
         SIPGrid.Columns(1).DefaultCellStyle.BackColor = Color.Aquamarine
         SIPGrid.Columns.Add("T1", "Oct-Apr1")
         SIPGrid.Columns(2).Width = 100 / FormWidthScaler
         SIPGrid.Columns(2).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T2", "May-June")
         SIPGrid.Columns(3).Width = 100 / FormWidthScaler
         SIPGrid.Columns(3).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T3", "July-Sept")
         SIPGrid.Columns(4).Width = 100 / FormWidthScaler
         SIPGrid.Columns(4).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T4", "Oct-Apr2")
         SIPGrid.Columns(5).Width = 100 / FormWidthScaler
         SIPGrid.Columns(5).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T5", "Total")
         SIPGrid.Columns(6).Width = 100 / FormWidthScaler
         SIPGrid.Columns(6).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.RowCount = NumFish * (MaxAge - MinAge + 1)
      ElseIf SpeciesName = "COHO" Then
         SIPGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         SIPGrid.Columns.Add("Name", "FisheryName")
         SIPGrid.Columns(0).Width = 250 / FormWidthScaler
         SIPGrid.Columns(0).ReadOnly = True
         SIPGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         SIPGrid.Columns.Add("T1", "Jan-June")
         SIPGrid.Columns(1).Width = 100 / FormWidthScaler
         SIPGrid.Columns(1).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T2", "July")
         SIPGrid.Columns(2).Width = 100 / FormWidthScaler
         SIPGrid.Columns(2).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T3", "August")
         SIPGrid.Columns(3).Width = 100 / FormWidthScaler
         SIPGrid.Columns(3).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T4", "September")
         SIPGrid.Columns(4).Width = 100 / FormWidthScaler
         SIPGrid.Columns(4).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T5", "Oct-Dec")
         SIPGrid.Columns(5).Width = 100 / FormWidthScaler
         SIPGrid.Columns(5).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.Columns.Add("T6", "Total")
         SIPGrid.Columns(6).Width = 100 / FormWidthScaler
         SIPGrid.Columns(6).DefaultCellStyle.Format = ("###0.00")
         SIPGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         SIPGrid.RowCount = NumFish
      End If

      '- Put Stock Impacts Per 1000 into Grid
      For Fish As Integer = 1 To NumFish
         FisheryTotalSum = 0
         For Age As Integer = MinAge To MaxAge
            FisherySum = 0
            If SpeciesName = "CHINOOK" Then
               If Age = MinAge Then
                  SIPGrid.Item(0, (((Fish - 1) * 4) + (Age - 2))).Value = FisheryName(Fish)
               Else
                  SIPGrid.Item(0, (((Fish - 1) * 4) + (Age - 2))).Value = "-"
               End If
               SIPGrid.Item(1, (((Fish - 1) * 4) + (Age - 2))).Value = Age.ToString
               For TStep = 1 To NumSteps
                  FisheryStepSum = TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
                  FisheryTotalSum += FisheryStepSum
                  If FisheryStepSum = 0 Then
                     SIPGrid.Item(TStep + 1, (((Fish - 1) * 4) + (Age - 2))).Value = "-----"
                  Else
                     TempVal = +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     FisherySum += TempVal
                     If TempVal = 0 Then
                        SIPGrid.Item(TStep + 1, (((Fish - 1) * 4) + (Age - 2))).Value = "*****"
                     Else
                        SIPGrid.Item(TStep + 1, (((Fish - 1) * 4) + (Age - 2))).Value = (TempVal * (1000 / (FisheryStepSum / ModelStockProportion(Fish)))).ToString(" ###0.00")
                     End If
                  End If
               Next
               If FisheryTotalSum = 0 Then
                  SIPGrid.Item(TStep + 1, (((Fish - 1) * 4) + (Age - 2))).Value = "-----"
               Else
                  SIPGrid.Item(TStep + 1, (((Fish - 1) * 4) + (Age - 2))).Value = (FisherySum * (1000 / (FisheryTotalSum / ModelStockProportion(Fish)))).ToString(" ###0.00")
               End If

            ElseIf SpeciesName = "COHO" Then
               SIPGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
               Age = 3
               For TStep = 1 To NumSteps
                  FisheryStepSum = TotalLandedCatch(Fish, TStep) + TotalNonRetention(Fish, TStep) + TotalShakers(Fish, TStep) + TotalDropOff(Fish, TStep)
                  FisheryTotalSum += FisheryStepSum
                  If FisheryStepSum = 0 Then
                     SIPGrid.Item(TStep, Fish - 1).Value = "-----"
                  Else
                     TempVal = +LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     FisherySum += TempVal
                     If TempVal = 0 Then
                        SIPGrid.Item(TStep, Fish - 1).Value = "*****"
                     Else
                        SIPGrid.Item(TStep, Fish - 1).Value = (TempVal * (1000 / FisheryStepSum)).ToString(" ###0.00")
                     End If
                  End If
               Next
               If FisheryTotalSum = 0 Then
                  SIPGrid.Item(TStep, Fish - 1).Value = "-----"
               Else
                  SIPGrid.Item(TStep, Fish - 1).Value = (FisherySum * (1000 / FisheryTotalSum)).ToString(" ###0.00")
               End If
            End If
         Next
      Next
   End Sub

Private Sub SIPGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles SIPGrid.CellContentClick

End Sub
End Class