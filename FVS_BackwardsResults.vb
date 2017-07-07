Public Class FVS_BackwardsResults

   Private Sub BROKButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BROKButton.Click
      Me.Close()
      FVS_BackwardsFram.Visible = True
      Exit Sub
   End Sub

   Private Sub FVS_BackwardsResults_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      Dim Stk As Integer, RowCount As Integer, NumTermStk As Integer, TermStk As Integer

      FormHeight = 880
      FormWidth = 1231
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
         If FVS_BackwardsResults_ReSize = False Then
            Resize_Form(Me)
            FVS_BackwardsResults_ReSize = True
         End If
      End If

      '- Fill the DataGrid with Values ... COHO and CHINOOK are different
      BFResultsGrid.Columns.Clear()
      BFResultsGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      BFResultsGrid.Rows.Clear()
      If SpeciesName = "COHO" Then
         Age = 3
         BFResultsGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)

         BFResultsGrid.Columns.Add("StockTitle", "Stock Name")
         BFResultsGrid.Columns("StockTitle").Width = 400 / FormWidthScaler
         BFResultsGrid.Columns("StockTitle").ReadOnly = True
         BFResultsGrid.Columns("StockTitle").DefaultCellStyle.BackColor = Color.Aquamarine
         BFResultsGrid.Columns("StockTitle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.Columns.Add("StockName", "Stk Abbrv")
         BFResultsGrid.Columns("StockName").Width = 120 / FormWidthScaler
         BFResultsGrid.Columns("StockName").ReadOnly = True
         BFResultsGrid.Columns("StockName").DefaultCellStyle.BackColor = Color.Aquamarine
         BFResultsGrid.Columns("StockName").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.Columns.Add("TargetEsc", "Target Esc")
         BFResultsGrid.Columns("TargetEsc").Width = 150 / FormWidthScaler
         BFResultsGrid.Columns("TargetEsc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.Columns.Add("FramEsc", "FRAM Esc")
         BFResultsGrid.Columns("FramEsc").Width = 150 / FormWidthScaler
         BFResultsGrid.Columns("FramEsc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.Columns.Add("StkScaler", "Stock Scaler")
         BFResultsGrid.Columns("StkScaler").Width = 150 / FormWidthScaler
         BFResultsGrid.Columns("StkScaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.Columns.Add("Flag", "Flag")
         BFResultsGrid.Columns("Flag").Width = 60 / FormWidthScaler
         BFResultsGrid.Columns("Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFResultsGrid.RowCount = NumStk

         For Stk = 1 To NumStk
            BFResultsGrid.Item(0, Stk - 1).Value = StockTitle(Stk)
            BFResultsGrid.Item(1, Stk - 1).Value = StockName(Stk)
            BFResultsGrid.Item(2, Stk - 1).Value = CLng(BackwardsTarget(Stk)).ToString
            BFResultsGrid.Item(3, Stk - 1).Value = CLng(Escape(Stk, 3, 5)).ToString
            BFResultsGrid.Item(4, Stk - 1).Value = StockRecruit(Stk, Age, 1).ToString("##0.0000")
            BFResultsGrid.Item(5, Stk - 1).Value = BackwardsFlag(Stk)
         Next

      ElseIf SpeciesName = "CHINOOK" Then

            If NumStk = 38 Or NumStk = 76 Then
                NumChinTermRuns = 37
            ElseIf NumStk = 33 Or NumStk = 66 Then
                NumChinTermRuns = 32
            Else
                NumChinTermRuns = NumStk / 2 - 1
            End If

         BFResultsGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         If BFResultsGrid.ColumnCount = 0 Then
            BFResultsGrid.Columns.Add("StockTitle", "Stock Name")
            BFResultsGrid.Columns("StockTitle").Width = 350 / FormWidthScaler
            BFResultsGrid.Columns("StockTitle").ReadOnly = True
            BFResultsGrid.Columns("StockTitle").DefaultCellStyle.BackColor = Color.Aquamarine

            BFResultsGrid.Columns.Add("StockName", "Stk Abbrv")
            BFResultsGrid.Columns("StockName").Width = 150 / FormWidthScaler
            BFResultsGrid.Columns("StockName").ReadOnly = True
            BFResultsGrid.Columns("StockName").DefaultCellStyle.BackColor = Color.Aquamarine

            BFResultsGrid.Columns.Add("AgeName", "Age")
            BFResultsGrid.Columns("AgeName").Width = 100 / FormWidthScaler
            BFResultsGrid.Columns("AgeName").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFResultsGrid.Columns.Add("FramEsc", "FRAM Esc")
            BFResultsGrid.Columns("FramEsc").Width = 150 / FormWidthScaler
            BFResultsGrid.Columns("FramEsc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFResultsGrid.Columns.Add("TargetEsc", "Target Esc")
            BFResultsGrid.Columns("TargetEsc").Width = 150 / FormWidthScaler
            BFResultsGrid.Columns("TargetEsc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFResultsGrid.Columns.Add("StkScaler", "Stock Scaler")
            BFResultsGrid.Columns("StkScaler").Width = 150 / FormWidthScaler
            BFResultsGrid.Columns("StkScaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFResultsGrid.Columns.Add("Flag", "Flag")
            BFResultsGrid.Columns("Flag").Width = 60 / FormWidthScaler
            BFResultsGrid.Columns("Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            'BFResultsGrid.FormatString = ">Stk#|>                    StkAbbv| Age|>   FRAM-Esc|> TargetEsc|>  StkScaler"
         End If

         '- Determine Number of Rows in DataGrid (Show Only Flagged Stocks)
         RowCount = 1
            For Stk = 1 To NumStk + NumChinTermRuns
                If TermStockNum(Stk) < 0 And BackwardsFlag(Stk) <> 0 Then
                    If NumStk > 65 Then
                        If TermStockNum(Stk) = -2 Then
                            RowCount = RowCount + 15
                            Stk = Stk + 4
                        Else
                            RowCount = RowCount + 9
                            Stk = Stk + 2
                        End If
                    Else
                        If TermStockNum(Stk) = -2 Then
                            RowCount = RowCount + 7
                            Stk = Stk + 2
                        Else
                            RowCount = RowCount + 3
                            Stk = Stk + 1
                        End If
                    End If
                ElseIf TermStockNum(Stk) > 0 And BackwardsFlag(Stk) <> 0 Then
                    RowCount = RowCount + 3
                Else
                    RowCount = RowCount
                End If
            Next Stk

         'BFResultsGrid.RowCount = NumStk + NumChinTermRuns
         BFResultsGrid.RowCount = RowCount

         'Put Stock Names into Array using DRV File order
         'For Stk = 1 To NumStk
         '   Select Case NumStk
         '      Case 66, 76
         '         If Stk = 1 Or Stk = 2 Then
         '            BFResultsGrid.Item(0, Stk).Value = "-----  " & StockTitle(Stk)
         '            BFResultsGrid.Item(1, Stk).Value = "-- " & StockName(Stk)
         '         ElseIf Stk > 2 And Stk < 7 Then
         '            BFResultsGrid.Item(0, Stk + 1).Value = "-----  " & StockTitle(Stk)
         '            BFResultsGrid.Item(1, Stk + 1).Value = "-- " & StockName(Stk)
         '            'TargetEsc.TextArray(fgi(Stk + 2, 0)) = StkName$(Stk)
         '         Else
         '            If (Stk Mod 2) = 0 Then
         '               '- Marked Name
         '               BFResultsGrid.Item(0, TermRunStock(Stk) * 3 + 1).Value = "-----  " & StockTitle(Stk)
         '               BFResultsGrid.Item(1, TermRunStock(Stk) * 3 + 1).Value = "-- " & StockName(Stk)
         '            Else
         '               '- UnMarked Name
         '               BFResultsGrid.Item(0, TermRunStock(Stk) * 3).Value = "-----  " & StockTitle(Stk)
         '               BFResultsGrid.Item(1, TermRunStock(Stk) * 3).Value = "-- " & StockName(Stk)
         '            End If
         '         End If
         '      Case 33, 38
         '         If Stk = 1 Then
         '            BFResultsGrid.Item(0, Stk).Value = "-----  " & StockTitle(Stk)
         '            BFResultsGrid.Item(1, Stk).Value = "-- " & StockName(Stk)
         '         ElseIf Stk > 1 And Stk < 4 Then
         '            BFResultsGrid.Item(0, Stk + 1).Value = "-----  " & StockTitle(Stk)
         '            BFResultsGrid.Item(1, Stk + 1).Value = "-- " & StockName(Stk)
         '         Else
         '            BFResultsGrid.Item(0, TermRunStock(Stk) * 2).Value = "-----  " & StockTitle(Stk)
         '            BFResultsGrid.Item(1, TermRunStock(Stk) * 2).Value = "-- " & StockName(Stk)
         '         End If
         '   End Select
         'Next Stk
         ''- Term Run Names
         'For Stk = 1 To NumChinTermRuns
         '   Select Case NumStk
         '      Case 66, 76
         '         If Stk > 2 Then
         '            BFResultsGrid.Item(0, Stk * 3 - 1).Value = TermRunName(Stk)
         '            BFResultsGrid.Item(1, Stk * 3 - 1).Value = "TOTAL TermRun"
         '         Else
         '            BFResultsGrid.Item(0, Stk * 3 - 3).Value = TermRunName(Stk)
         '            BFResultsGrid.Item(1, Stk * 3 - 3).Value = "TOTAL TermRun"
         '         End If
         '      Case 33, 38
         '         If Stk > 2 Then
         '            BFResultsGrid.Item(0, Stk * 2 - 1).Value = TermRunName(Stk)
         '            BFResultsGrid.Item(1, Stk * 2 - 1).Value = "*NOT USED*"
         '         Else
         '            BFResultsGrid.Item(0, Stk * 2 - 2).Value = TermRunName(Stk)
         '            BFResultsGrid.Item(1, Stk * 2 - 2).Value = "*NOT USED*"
         '         End If
         '   End Select
         'Next Stk

         'For Stk = 1 To NumStk + NumChinTermRuns
         '   For Age = 3 To 5
         '      If TermStockNum(Stk) < 0 And NumStk < 66 Then  '- TermRuns NOT USED for Non-Selective Base
         '         BFResultsGrid.Item(Age - 1, Stk - 1).Value = "*****"
         '      Else
         '         BFResultsGrid.Item(Age - 1, Stk - 1).Value = BackwardsChinook(Stk, Age)
         '      End If
         '   Next
         '   If TermStockNum(Stk) < 0 And NumStk < 66 Then
         '      BFResultsGrid.Item(5, Stk - 1).Value = "*"
         '   Else
         '      BFResultsGrid.Item(5, Stk - 1).Value = BackwardsFlag(Stk)
         '   End If
         'Next Stk
         '===========================================================================================

         '====================================================================================

         '- Results
         RowCount = 0
         For Stk = 1 To NumStk + NumChinTermRuns
            If TermStockNum(Stk) < 0 And BackwardsFlag(Stk) = 0 Then GoTo SkipTermRun
            If TermStockNum(Stk) < 0 And BackwardsFlag(Stk) <> 0 Then
               '- Print Combined Terminal Run plus Stock Components
               For Age As Integer = 3 To 5
                  RowCount += 1
                  '- FRAM Esc for Combined Terminal Run
                  If Age = 3 Then
                     BFResultsGrid.Item(0, RowCount).Value = TermRunName(TermStockNum(Stk) * -1)
                  Else
                     BFResultsGrid.Item(0, RowCount).Value = " "
                  End If
                  BFResultsGrid.Item(2, RowCount).Value = Age.ToString
                        'If TermStockNum(Stk) = -2 Then
                        '- Nooksack Spring Special Case - 4 stocks
                        BFResultsGrid.Item(3, RowCount).Value = CLng(TermChinRun(Stk, Age)).ToString
                        ' Else
                        '- Normal Sum - 2 stocks = Marked and UnMarked
                        ' BFResultsGrid.Item(3, RowCount).Value = CLng(TermChinRun(Stk, Age)).ToString
                        'End If
                        '- Target Escapement
                        BFResultsGrid.Item(4, RowCount).Value = CLng(BackwardsChinook(Stk, Age)).ToString
                        BFResultsGrid.Item(6, RowCount).Value = BackwardsFlag(Stk).ToString
                    Next Age
               'here i am
               If TermStockNum(Stk) = -2 Then
                  NumTermStk = 4
               Else
                  NumTermStk = 2
               End If
               '- Print the Component Stocks of Selected Terminal Run
               For TermStk = 1 To NumTermStk
                  For Age As Integer = 3 To 5
                     RowCount = RowCount + 1
                     BFResultsGrid.Item(0, RowCount).Value = "----" & StockTitle(TermStockNum(Stk + TermStk))
                     BFResultsGrid.Item(1, RowCount).Value = StockName(TermStockNum(Stk + TermStk))
                     BFResultsGrid.Item(2, RowCount).Value = Age.ToString
                     BFResultsGrid.Item(3, RowCount).Value = CLng(TermChinRun(Stk + TermStk, Age)).ToString
                     BFResultsGrid.Item(4, RowCount).Value = CLng(BackwardsChinook(Stk + TermStk, Age)).ToString
                     BFResultsGrid.Item(5, RowCount).Value = StockRecruit(TermStockNum(Stk + TermStk), Age, 1).ToString("##0.0000")
                     BFResultsGrid.Item(6, RowCount).Value = BackwardsFlag(Stk + TermStk).ToString
                  Next Age
               Next TermStk
               '- Change Loop Variable to Skip TermRun Component Stocks
               If TermStockNum(Stk) = -2 Then
                  Stk = Stk + 4
               Else
                  Stk = Stk + 2
               End If
            ElseIf TermStockNum(Stk) > 0 And BackwardsFlag(Stk) <> 0 Then
               '- Print Single Stock Flagged as Targets - TermRun Not Selected
               For Age As Integer = 3 To 5
                  RowCount = RowCount + 1
                  BFResultsGrid.Item(0, RowCount).Value = "----" & StockTitle(TermStockNum(Stk))
                  BFResultsGrid.Item(1, RowCount).Value = StockName(TermStockNum(Stk))
                  BFResultsGrid.Item(2, RowCount).Value = Age.ToString
                  BFResultsGrid.Item(3, RowCount).Value = CLng(TermChinRun(Stk, Age)).ToString
                  BFResultsGrid.Item(4, RowCount).Value = CLng(BackwardsChinook(Stk, Age)).ToString
                  BFResultsGrid.Item(5, RowCount).Value = StockRecruit(TermStockNum(Stk), Age, 1).ToString("##0.0000")
                  BFResultsGrid.Item(6, RowCount).Value = BackwardsFlag(Stk).ToString
               Next Age
            End If
SkipTermRun:
         Next Stk
      End If
   End Sub

Private Sub BFResultsGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles BFResultsGrid.CellContentClick

End Sub
End Class