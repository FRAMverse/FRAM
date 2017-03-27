Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class FVS_BackwardsTarget

   Private Sub FVS_BackwardsTarget_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 891
      FormHeight = 905
      FormWidth = 1022
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
         If FVS_BackwardsTarget_ReSize = False Then
            Resize_Form(Me)
            FVS_BackwardsTarget_ReSize = True
         End If
      End If

      '- Fill the DataGrid with Values ... COHO and CHINOOK are different
      BFTargetGrid.Columns.Clear()
      BFTargetGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      BFTargetGrid.Rows.Clear()
      If SpeciesName = "COHO" Then
         BFTargetGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)

         BFTargetGrid.Columns.Add("StockTitle", "Stock Name")
         BFTargetGrid.Columns("StockTitle").Width = 400 / FormWidthScaler
         BFTargetGrid.Columns("StockTitle").ReadOnly = True
         BFTargetGrid.Columns("StockTitle").DefaultCellStyle.BackColor = Color.Aquamarine
         BFTargetGrid.Columns("StockTitle").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFTargetGrid.Columns.Add("StockName", "Stk Abbrv")
         BFTargetGrid.Columns("StockName").Width = 100 / FormWidthScaler
         BFTargetGrid.Columns("StockName").ReadOnly = True
         BFTargetGrid.Columns("StockName").DefaultCellStyle.BackColor = Color.Aquamarine
         BFTargetGrid.Columns("StockName").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFTargetGrid.Columns.Add("TargetEsc", "Target Esc")
         BFTargetGrid.Columns("TargetEsc").Width = 100 / FormWidthScaler
         BFTargetGrid.Columns("TargetEsc").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         BFTargetGrid.Columns.Add("Flag", "FLAG")
         BFTargetGrid.Columns("Flag").Width = 60 / FormWidthScaler
         BFTargetGrid.Columns("Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         BFTargetGrid.RowCount = NumStk

         For Stk As Integer = 1 To NumStk
            BFTargetGrid.Item(0, Stk - 1).Value = StockTitle(Stk)
            BFTargetGrid.Item(1, Stk - 1).Value = StockName(Stk)
            BFTargetGrid.Item(2, Stk - 1).Value = CLng(BackwardsTarget(Stk))
            BFTargetGrid.Item(3, Stk - 1).Value = BackwardsFlag(Stk)
         Next

      ElseIf SpeciesName = "CHINOOK" Then

            If NumStk = 38 Or NumStk = 76 Then
                NumChinTermRuns = 37
            ElseIf NumStk = 33 Or NumStk = 66 Then
                NumChinTermRuns = 32
            Else
                NumChinTermRuns = NumStk / 2 - 1
            End If

         Call FVS_BackwardsFram.BackChinArrays()

         BFTargetGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(8 / FormWidthScaler), FontStyle.Bold)
         If BFTargetGrid.ColumnCount = 0 Then
            BFTargetGrid.Columns.Add("StockTitle", "Stock Name")
            BFTargetGrid.Columns("StockTitle").Width = 350 / FormWidthScaler
            BFTargetGrid.Columns("StockTitle").ReadOnly = True
            BFTargetGrid.Columns("StockTitle").DefaultCellStyle.BackColor = Color.Aquamarine

            BFTargetGrid.Columns.Add("StockName", "Stk Abbrv")
            BFTargetGrid.Columns("StockName").Width = 150 / FormWidthScaler
            BFTargetGrid.Columns("StockName").ReadOnly = True
            BFTargetGrid.Columns("StockName").DefaultCellStyle.BackColor = Color.Aquamarine

            BFTargetGrid.Columns.Add("Age3TermRun", "Age-3")
            BFTargetGrid.Columns("Age3TermRun").Width = 100 / FormWidthScaler
            BFTargetGrid.Columns("Age3TermRun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFTargetGrid.Columns.Add("Age4TermRun", "Age-4")
            BFTargetGrid.Columns("Age4TermRun").Width = 100 / FormWidthScaler
            BFTargetGrid.Columns("Age4TermRun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFTargetGrid.Columns.Add("Age5TermRun", "Age-5")
            BFTargetGrid.Columns("Age5TermRun").Width = 100 / FormWidthScaler
            BFTargetGrid.Columns("Age5TermRun").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            BFTargetGrid.Columns.Add("ChinFlag", "FLAG")
            BFTargetGrid.Columns("ChinFlag").Width = 60 / FormWidthScaler
            BFTargetGrid.Columns("ChinFlag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         End If

            BFTargetGrid.RowCount = NumStk + NumChinTermRuns

         'Put Stock Names into Array using DRV File order
         For Stk As Integer = 1 To NumStk
            Select Case NumStk
                    Case Is > 65
                        If Stk = 1 Or Stk = 2 Then
                            BFTargetGrid.Item(0, Stk).Value = "-----  " & StockTitle(Stk)
                            BFTargetGrid.Item(1, Stk).Value = "-- " & StockName(Stk)
                        ElseIf Stk > 2 And Stk < 7 Then
                            BFTargetGrid.Item(0, Stk + 1).Value = "-----  " & StockTitle(Stk)
                            BFTargetGrid.Item(1, Stk + 1).Value = "-- " & StockName(Stk)
                        Else
                            If (Stk Mod 2) = 0 Then
                                '- Marked Name
                                BFTargetGrid.Item(0, TermRunStock(Stk) * 3 + 1).Value = "-----  " & StockTitle(Stk)
                                BFTargetGrid.Item(1, TermRunStock(Stk) * 3 + 1).Value = "-- " & StockName(Stk)
                            Else
                                '- UnMarked Name
                                BFTargetGrid.Item(0, TermRunStock(Stk) * 3).Value = "-----  " & StockTitle(Stk)
                                BFTargetGrid.Item(1, TermRunStock(Stk) * 3).Value = "-- " & StockName(Stk)
                            End If
                        End If
               Case 33, 38
                  If Stk = 1 Then
                     BFTargetGrid.Item(0, Stk).Value = "-----  " & StockTitle(Stk)
                     BFTargetGrid.Item(1, Stk).Value = "-- " & StockName(Stk)
                  ElseIf Stk > 1 And Stk < 4 Then
                     BFTargetGrid.Item(0, Stk + 1).Value = "-----  " & StockTitle(Stk)
                     BFTargetGrid.Item(1, Stk + 1).Value = "-- " & StockName(Stk)
                  Else
                     BFTargetGrid.Item(0, TermRunStock(Stk) * 2).Value = "-----  " & StockTitle(Stk)
                     BFTargetGrid.Item(1, TermRunStock(Stk) * 2).Value = "-- " & StockName(Stk)
                  End If
            End Select
         Next Stk
         '- Term Run Names
         For Stk As Integer = 1 To NumChinTermRuns

                If NumStk > 65 Then
                    If Stk > 2 Then
                        BFTargetGrid.Item(0, Stk * 3 - 1).Value = TermRunName(Stk)
                        BFTargetGrid.Item(1, Stk * 3 - 1).Value = "TOTAL TermRun"
                    Else
                        BFTargetGrid.Item(0, Stk * 3 - 3).Value = TermRunName(Stk)
                        BFTargetGrid.Item(1, Stk * 3 - 3).Value = "TOTAL TermRun"
                    End If
                Else
                    If Stk > 2 Then
                        BFTargetGrid.Item(0, Stk * 2 - 1).Value = TermRunName(Stk)
                        BFTargetGrid.Item(1, Stk * 2 - 1).Value = "*NOT USED*"
                    Else
                        BFTargetGrid.Item(0, Stk * 2 - 2).Value = TermRunName(Stk)
                        BFTargetGrid.Item(1, Stk * 2 - 2).Value = "*NOT USED*"
                    End If
                End If
            Next Stk

         For Stk As Integer = 1 To NumStk + NumChinTermRuns
            For Age As Integer = 3 To 5
               If TermStockNum(Stk) < 0 And NumStk < 66 Then  '- TermRuns NOT USED for Non-Selective Base
                  BFTargetGrid.Item(Age - 1, Stk - 1).Value = "*****"
               Else
                  BFTargetGrid.Item(Age - 1, Stk - 1).Value = BackwardsChinook(Stk, Age)
               End If
            Next
            If TermStockNum(Stk) < 0 And NumStk < 66 Then
               BFTargetGrid.Item(5, Stk - 1).Value = "*"
            Else
               BFTargetGrid.Item(5, Stk - 1).Value = BackwardsFlag(Stk)
            End If
         Next Stk
      End If

   End Sub

   Private Sub BTCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTCancelButton.Click
      Me.Close()
      FVS_BackwardsFram.Visible = True
      Exit Sub
   End Sub

   Private Sub BTEscapementButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTEscapementButton.Click

      Dim OpenBACKFRAMspreadsheet As New OpenFileDialog()
      Dim I As Integer

        OpenBACKFRAMspreadsheet.Filter = "BACKFRAM Spreadsheets ((*.xls; *.xlsx; *xlsm)|*.xls; *.xlsx; *xlsm|All files (*.*)|*.*"
      OpenBACKFRAMspreadsheet.FilterIndex = 1
      OpenBACKFRAMspreadsheet.RestoreDirectory = True

      If OpenBACKFRAMspreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         BACKFRAMSpreadSheet = OpenBACKFRAMspreadsheet.FileName
         BACKFRAMSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).DirectoryName
      Else
         Exit Sub
      End If

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
      wbName = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(BACKFRAMSpreadSheet)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '-Check if WorkBook contains any FRAMEscapeV2 WorkSheet
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 11 Then
            If xlWorkSheet.Name = "FRAMEscapeV2" Then GoTo FoundBFEscape
         End If
      Next

      MsgBox("Can't Find 'FRAMEscapeV2' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
             "Please Choose appropriate Spreadsheet with Backwards FRAM Escapement!", MsgBoxStyle.OkOnly)
      GoTo CloseEscWorkbook

FoundBFEscape:
      ' Get Columns (Years) with Escapement Values for User Selection


      ' **************** On 6 July 2012, Pete Added an IF SpeciesName = "COHO" and "CHINOOK" component to
      ' **************** FoundBFEscape so that Chinook terminal runs could be easily imported for post-season runs
      '***************** If this material is not desired, delete all lines followed by "*&*&* PM added 7/6/12"

      If SpeciesName = "COHO" Then '*&*&* PM added 7/6/12

         ReDim BFEscYears(50)
         Dim EscCount As Integer
         Dim EscValue As Double
         Dim EscColumn As String
         EscCount = 0
         For I = 1 To 50
            If I < 26 Then
               EscValue = xlWorkSheet.Range(Chr(I + 65) & "2").Value
            Else
               EscValue = xlWorkSheet.Range("A" & Chr(I + 39) & "2").Value
            End If
            If EscValue >= 1973 And EscValue <= 2023 Then
               EscCount = EscCount + 1
               BFEscYears(EscCount) = EscValue
            End If
         Next I

         Me.Enabled = False
         BFYearSelectType = 1
         FVS_BackwardsYearSelect.ShowDialog()
         Me.BringToFront()

         '- User canceled year selection
         If BFYearSelection = 0 Then GoTo CloseEscWorkbook

         If BFYearSelection >= 1973 And BFYearSelection <= 2023 Then '- valid selection
                '- Find correct Column

            For I = 1 To 50
               If I < 26 Then
                  EscValue = xlWorkSheet.Range(Chr(I + 65) & "2").Value
                  EscColumn = Chr(I + 65)
               Else
                  EscValue = xlWorkSheet.Range("A" & Chr(I + 39) & "2").Value
                  EscColumn = "A" & Chr(I + 39)
               End If
               If EscValue = BFYearSelection Then
                  ' Found Column ... Load Esc Values 
                        For Stk As Integer = 1 To NumStk
                            'zero out grid before loading new values
                            BFTargetGrid.Item(2, Stk - 1).Value = 0
                            BFTargetGrid.Item(3, Stk - 1).Value = 0
                            If IsNumeric(xlWorkSheet.Range(EscColumn & CStr(Stk + 2)).Value) Then
                                EscValue = CDbl(CLng(xlWorkSheet.Range(EscColumn & CStr(Stk + 2)).Value))
                            Else
                                EscValue = 0

                            End If
                            If EscValue <> 0 Then
                                BackwardsTarget(Stk) = EscValue
                                If StockName(Stk).Substring(7, 1) = "h" Then
                                    '- Make Hatchery Stocks use FLAG = 2 (Marked-UnMarked Ratio from StkScalers)
                                    BackwardsFlag(Stk) = 2
                                Else
                                    BackwardsFlag(Stk) = 1
                                End If
                                BFTargetGrid.Item(2, Stk - 1).Value = BackwardsTarget(Stk)
                                BFTargetGrid.Item(3, Stk - 1).Value = BackwardsFlag(Stk)
                            End If
                        Next
               End If
            Next I
         End If

         ChangeBackFram = True

CloseEscWorkbook:

         '- Done with TAMM WorkBook for this run .. Close and release object
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

         '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
         'Pete's modification to facilitate reading in of target escapements for Chinook during 2012 post-season runs.
      ElseIf SpeciesName = "CHINOOK" Then '*&*&* PM added 7/6/12

         Dim Esc3Value As Double '*&*&* PM added 7/6/12
         Dim Esc4Value As Double '*&*&* PM added 7/6/12
         Dim Esc5Value As Double '*&*&* PM added 7/6/12
         Dim EscFlagChin As Integer '*&*&* PM added 7/6/12

            For I = 2 To NumStk + NumChinTermRuns + 1  '*&*&* PM added 7/6/12
                Esc3Value = Math.Round(xlWorkSheet.Range("C" & CStr(I)).Value) '*&*&* PM added 7/6/12
                Esc4Value = Math.Round(xlWorkSheet.Range("D" & CStr(I)).Value) '*&*&* PM added 7/6/12
                Esc5Value = Math.Round(xlWorkSheet.Range("E" & CStr(I)).Value) '*&*&* PM added 7/6/12
                EscFlagChin = xlWorkSheet.Range("F" & CStr(I)).Value  '*&*&* PM added 7/6/12

                BFTargetGrid.Item(2, I - 2).Value = Esc3Value   '*&*&* PM added 7/6/12
                BFTargetGrid.Item(3, I - 2).Value = Esc4Value   '*&*&* PM added 7/6/12
                BFTargetGrid.Item(4, I - 2).Value = Esc5Value   '*&*&* PM added 7/6/12
                BFTargetGrid.Item(5, I - 2).Value = EscFlagChin '*&*&* PM added 7/6/12

            Next I   '*&*&* PM added 7/6/12

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

         
      End If  '*&*&* PM added 7/6/12
      '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$



   End Sub

   Private Sub BTOKButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTOKButton.Click

      '- Put User Input from Grid into Backwards FRAM Arrays
        ReDim BackwardsFlag(NumStk + NumChinTermRuns)
        ReDim BackwardsTarget(NumStk + +NumChinTermRuns)
      If SpeciesName = "COHO" Then
         For Stk As Integer = 1 To NumStk
            If BackwardsTarget(Stk) <> CDbl(BFTargetGrid.Item(2, Stk - 1).Value) Then
               ChangeBackFram = True
               BackwardsTarget(Stk) = BFTargetGrid.Item(2, Stk - 1).Value
            End If
            If BackwardsFlag(Stk) <> BFTargetGrid.Item(3, Stk - 1).Value Then
               ChangeBackFram = True
               BackwardsFlag(Stk) = BFTargetGrid.Item(3, Stk - 1).Value
            End If
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         For Stk As Integer = 1 To NumStk + NumChinTermRuns
            If TermStockNum(Stk) < 0 And NumStk < 66 Then '- TermRuns NOT USED for Non-Selective Base
               For Age As Integer = 3 To 5
                  BackwardsChinook(Stk, Age) = 0
               Next Age
               BackwardsFlag(Stk) = 0
            Else
               For Age As Integer = 3 To 5
                  If BackwardsChinook(Stk, Age) <> BFTargetGrid.Item(Age - 1, Stk - 1).Value Then
                     ChangeBackFram = True
                     BackwardsChinook(Stk, Age) = BFTargetGrid.Item(Age - 1, Stk - 1).Value
                  End If
               Next Age
               If BackwardsFlag(Stk) <> BFTargetGrid.Item(5, Stk - 1).Value Then
                  ChangeBackFram = True
                  BackwardsFlag(Stk) = BFTargetGrid.Item(5, Stk - 1).Value
               End If
            End If
         Next Stk
      End If

      Me.Close()
      FVS_BackwardsFram.Visible = True
      Exit Sub

   End Sub

   Private Sub BTFillSSButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTFillSSButton.Click

      Dim OpenBACKFRAMspreadsheet As New OpenFileDialog()
      Dim I As Integer

      MsgBox("Please Choose Year (Column) to Fill with Current Values", MsgBoxStyle.OkOnly)
      OpenBACKFRAMspreadsheet.Filter = "BACKFRAM Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenBACKFRAMspreadsheet.FilterIndex = 1
      OpenBACKFRAMspreadsheet.RestoreDirectory = True

      If OpenBACKFRAMspreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         BACKFRAMSpreadSheet = OpenBACKFRAMspreadsheet.FileName
         BACKFRAMSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).DirectoryName
      Else
         Exit Sub
      End If

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
      wbName = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(BACKFRAMSpreadSheet)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '-Check if WorkBook contains any FRAMEscapeV2 WorkSheet
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 11 Then
            If xlWorkSheet.Name = "FRAMEscapeV2" Then GoTo FoundBFEscape
         End If
      Next

      MsgBox("Can't Find 'FRAMEscapeV2' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
             "Please Choose appropriate Spreadsheet with Backwards FRAM Escapement!", MsgBoxStyle.OkOnly)
      GoTo CloseEscWorkbook

FoundBFEscape:
      ' Get Columns (Years) with Escapement Values for User Selection
      ReDim BFEscYears(50)
      Dim EscCount As Integer
      Dim EscValue As Double
      Dim EscColumn As String
      EscCount = 0
      For I = 1 To 50
         If I < 26 Then
            EscValue = xlWorkSheet.Range(Chr(I + 65) & "2").Value
         Else
            EscValue = xlWorkSheet.Range("A" & Chr(I + 39) & "2").Value
         End If
         If EscValue >= 1973 And EscValue <= 2023 Then
            EscCount = EscCount + 1
            BFEscYears(EscCount) = EscValue
         End If
      Next I

      Me.Enabled = False
      BFYearSelectType = 1
      FVS_BackwardsYearSelect.ShowDialog()
      Me.BringToFront()

      '- User canceled year selection
      If BFYearSelection = 0 Then GoTo CloseEscWorkbook

      If BFYearSelection >= 1973 And BFYearSelection <= 2023 Then '- valid selection
         '- Find correct Column
         For I = 1 To 50
            If I < 26 Then
               EscValue = xlWorkSheet.Range(Chr(I + 65) & "2").Value
               EscColumn = Chr(I + 65)
            Else
               EscValue = xlWorkSheet.Range("A" & Chr(I + 39) & "2").Value
               EscColumn = "A" & Chr(I + 39)
            End If
            If EscValue = BFYearSelection Then
               ' Found Column ... Load Esc Values into Spreadsheet
               For Stk As Integer = 1 To NumStk
                  xlWorkSheet.Range(EscColumn & CStr(Stk + 2)).Value = BackwardsTarget(Stk)
               Next
            End If
         Next I
      End If

CloseEscWorkbook:

      '- Done with TAMM WorkBook for this run .. Close and release object
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



   '--------------------------=============================
   Private Sub BTCatchButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BTCatchButton.Click

      If SpeciesName = "CHINOOK" Then
         MsgBox("Currently there are no Backwards FRAM Chinook Input Spreadsheets" & "Function to be Implemented later", MsgBoxStyle.OkOnly)
         Exit Sub
      End If

      Dim OpenBACKFRAMspreadsheet As New OpenFileDialog()
      Dim I As Integer

      '==================
      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenBACKFRAMspreadsheet.Filter = "BACKFRAM Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenBACKFRAMspreadsheet.FilterIndex = 1
      OpenBACKFRAMspreadsheet.RestoreDirectory = True

      If OpenBACKFRAMspreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         BACKFRAMSpreadSheet = OpenBACKFRAMspreadsheet.FileName
         BACKFRAMSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).DirectoryName
      Else
         Exit Sub
      End If

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
      wbName = My.Computer.FileSystem.GetFileInfo(BACKFRAMSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(BACKFRAMSpreadSheet)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '- Find WorkSheets with FRAM Catch numbers
      ReDim BFCatchYears(50)
      I = 1
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name.Substring(4, 4) = "FRAM" And IsNumeric(xlWorkSheet.Name.Substring(0, 4)) Then
               BFCatchYears(I) = CInt(xlWorkSheet.Name.Substring(0, 4))
               I += 1
            End If
         End If
      Next

      '-Check if WorkBook contains any FRAM#### WorkSheets
      If I = 1 Then
         MsgBox("Can't Find BackFRAM Catch WorkSheet in your DataBase Selection" & vbCrLf & _
                "Please Choose appropriate DataBase with Backwards FRAM Catch!", MsgBoxStyle.OkOnly)
         GoTo CloseBFWorkBook
      End If

      '- User Year Selection
      BFYearSelectType = 2
      Me.Enabled = False
      FVS_BackwardsYearSelect.ShowDialog()
      Me.BringToFront()

      '- Find WorkSheet matching Year Selection
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name.Substring(4, 4) = "FRAM" And IsNumeric(xlWorkSheet.Name.Substring(0, 4)) Then
               If CInt(xlWorkSheet.Name.Substring(0, 4)) = BFYearSelection Then Exit For
            End If
         End If
      Next

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Dim CellAddress As String
      Dim FlagAddress As String
      Dim FlagValue As Integer
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            CellAddress = ""
            FlagAddress = ""
            Select Case TStep
               Case 1
                  CellAddress = "C" & CStr(Fish + 3)
                  FlagAddress = "D" & CStr(Fish + 3)
               Case 2
                  CellAddress = "E" & CStr(Fish + 3)
                  FlagAddress = "F" & CStr(Fish + 3)
               Case 3
                  CellAddress = "G" & CStr(Fish + 3)
                  FlagAddress = "H" & CStr(Fish + 3)
               Case 4
                  CellAddress = "I" & CStr(Fish + 3)
                  FlagAddress = "J" & CStr(Fish + 3)
               Case 5
                  CellAddress = "K" & CStr(Fish + 3)
                  FlagAddress = "L" & CStr(Fish + 3)
            End Select
            If IsNumeric(xlWorkSheet.Range(CellAddress).Value) Then
               If CInt(xlWorkSheet.Range(CellAddress).Value) < 0 Or CInt(xlWorkSheet.Range(CellAddress).Value) > 999999 Then GoTo NextBFCatch
               FisheryQuota(Fish, TStep) = CInt(xlWorkSheet.Range(CellAddress).Value)
               If IsNumeric(xlWorkSheet.Range(FlagAddress).Value) Then
                  FlagValue = xlWorkSheet.Range(FlagAddress).Value
                  If FlagValue = 8 Then
                     FisheryFlag(Fish, TStep) = 8
                  Else
                     FisheryFlag(Fish, TStep) = 2
                  End If
               End If
            End If
NextBFCatch:
         Next
      Next

      ChangeBackFram = True

CloseBFWorkBook:
      '- Done with TAMM WorkBook for this run .. Close and release object
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

    Private Sub BFTargetGrid_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles BFTargetGrid.CellContentClick

    End Sub
End Class