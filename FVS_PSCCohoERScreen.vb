Public Class FVS_PSCCohoERScreen
   Public PSCER(,) As Double
   Public PSCStockName() As String

   Private Sub FVS_PSCCohoERScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      Dim Col As Integer

      FormHeight = 767
      FormWidth = 1165
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
         If FVS_PSCCohoERScreen_ReSize = False Then
            Resize_Form(Me)
            FVS_PSCCohoERScreen_ReSize = True
         End If
      End If

      PSCERGrid.Columns.Clear()
      PSCERGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      PSCERGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      If SpeciesName = "CHINOOK" Then
         Me.Close()
         FVS_ScreenReports.Visible = True
      ElseIf SpeciesName = "COHO" Then
         PSCERGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         PSCERGrid.Columns.Add("Name", "StockName")
         PSCERGrid.Columns(0).Width = 175 / FormWidthScaler
         PSCERGrid.Columns(0).ReadOnly = True
         PSCERGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         PSCERGrid.Columns.Add("T1", "US Tot-ER")
         PSCERGrid.Columns(1).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(1).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T2", "US Ocean")
         PSCERGrid.Columns(2).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(2).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T3", "US PugSnd")
         PSCERGrid.Columns(3).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(3).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T4", "US Other")
         PSCERGrid.Columns(4).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(4).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T5", "BC Tot-ER")
         PSCERGrid.Columns(5).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(5).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.RowCount = NumFish
         PSCERGrid.Columns.Add("T6", "BC Ocean")
         PSCERGrid.Columns(6).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(6).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T7", "BC GeoStr")
         PSCERGrid.Columns(7).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(7).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T8", "BC Other")
         PSCERGrid.Columns(8).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(8).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.Columns.Add("T9", "Total ER")
         PSCERGrid.Columns(9).Width = 100 / FormWidthScaler
         PSCERGrid.Columns(9).DefaultCellStyle.Format = ("###0.0000")
         PSCERGrid.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PSCERGrid.RowCount = 17
      End If

      '- Sum Mortalities, Calculate ERs
      Call CalculatePSCCohoER()

      '- Put ERs into Grid
      For Stk = 1 To 17
         PSCERGrid.Item(0, Stk - 1).Value = PSCStockName(Stk)
         For Col = 2 To 9
            PSCERGrid.Item(Col - 1, Stk - 1).Value = PSCER(Stk, Col).ToString("###0.0000")
         Next
         '- Totals Line
         PSCERGrid.Item(Col - 1, Stk - 1).Value = (PSCER(Stk, 2) + PSCER(Stk, 6)).ToString("###0.0000")
      Next

   End Sub

   Private Sub PSCERExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PSCERExitButton.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
   End Sub

   Private Sub ClipBoardCopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClipBoardCopyToolStripMenuItem.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      ClipStr = "COHO "
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr
      ClipStr &= "StockName" & vbTab & "US Tot-ER" & vbTab & "US Ocean" & vbTab & "US PgtSnd" & vbTab & "US Other" & vbTab & "BC Tot-ER" & vbTab & "BC Ocean" & vbTab & "BC GeoStr" & vbTab & "BC Other" & vbTab & "Total ER" & vbCr
      For RecNum = 0 To 16
         For ColNum = 0 To 9
            If ColNum = 0 Then
               ClipStr = ClipStr & PSCERGrid.Item(ColNum, RecNum).Value
            Else
               ClipStr = ClipStr & vbTab & CDbl(PSCERGrid.Item(ColNum, RecNum).Value)
            End If
         Next
         ClipStr &= vbCr
      Next
      Clipboard.SetDataObject(ClipStr)

   End Sub

   Sub CalculatePSCCohoER()
      Dim PSCGroup(17, 5) As Integer
      Dim StkGroup, StkList, Col As Integer
      Dim StkTotal As Double
      ReDim PSCER(17, 9)
      ReDim PSCStockName(17)

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
      For Fish = 1 To NumFish
         For StkGroup = 1 To 17
            For StkList = 1 To PSCGroup(StkGroup, 0)
               Stk = PSCGroup(StkGroup, StkList)
               For TStep = 1 To NumSteps
                  If NumStk = 256 Then
                     If Fish > 166 And Fish < 202 Then '- Canadian Catch
                        PSCER(StkGroup, 3) = PSCER(StkGroup, 3) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     Else
                        PSCER(StkGroup, 2) = PSCER(StkGroup, 2) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                  Else
                     If Fish > 166 And Fish < 194 Then
                        '- Canadian Catch
                        PSCER(StkGroup, 6) = PSCER(StkGroup, 6) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish < 167 Or Fish > 193 Then
                        '- US Catch
                        PSCER(StkGroup, 2) = PSCER(StkGroup, 2) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish > 0 And Fish < 23 Or Fish > 32 And Fish < 44 Or Fish = 79 Or Fish > 193 And Fish < 199 Then 'US Ocean
                        PSCER(StkGroup, 3) = PSCER(StkGroup, 3) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish = 44 Or Fish > 79 And Fish < 167 Then
                        '- Puget Sound
                        PSCER(StkGroup, 4) = PSCER(StkGroup, 4) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish > 22 And Fish < 33 Or Fish > 44 And Fish < 79 Then
                        '- US Other
                        PSCER(StkGroup, 5) = PSCER(StkGroup, 5) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish > 170 And Fish < 176 Or Fish > 177 And Fish < 182 Or Fish > 186 And Fish < 189 Or Fish = 190 Then 'BC Ocean
                        PSCER(StkGroup, 7) = PSCER(StkGroup, 7) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish = 176 Or Fish = 183 Or Fish > 190 And Fish < 193 Then
                        '- Georgia Strait
                        PSCER(StkGroup, 8) = PSCER(StkGroup, 8) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                     If Fish > 166 And Fish < 171 Or Fish = 177 Or Fish = 182 Or Fish > 183 And Fish < 187 Or Fish = 189 Or Fish = 193 Then
                        '- BC Other
                        PSCER(StkGroup, 9) = PSCER(StkGroup, 9) + LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep)
                     End If
                  End If
               Next
            Next
         Next
      Next
      '- ReCalcualte Array for Exploitation Rates
      For Stk = 1 To 17
         '- Escapement plus US Mortality plus BC Mortality
         StkTotal = PSCER(Stk, 1) + PSCER(Stk, 2) + PSCER(Stk, 6)
         For Col = 2 To 9
            PSCER(Stk, Col) = PSCER(Stk, Col) / StkTotal
         Next
      Next

   End Sub

End Class