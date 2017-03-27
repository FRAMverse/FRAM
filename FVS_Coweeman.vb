'- 3/15/12 McHugh: commented out the TRShare() summations that were repeated over SOF, PS, and WA Coast, as they were incorrectly added to the PFMC total in Coweeman transfer)
'- 3/16/12 McHugh: added a NoF & SoF rec to the Coweeman transfer to fulfill Treaty (H. Leon, Makah) interests in Coweeman summary outputs
'- 3/19/12 McHugh: automated the transfer of Cindy's summer Chinook summary



Imports Microsoft.Office.Interop
Imports System.IO.File
Public Class FVS_Coweeman
   Public PFMCOption As Integer

   Private Sub Option1Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option1Button.Click
      PFMCOption = 1
      TransferCoweemanData()
      Me.Close()
      FVS_FramUtils.Visible = True
      Exit Sub
   End Sub

   Private Sub Option2Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option2Button.Click
      PFMCOption = 2
      TransferCoweemanData()
      Me.Close()
      FVS_FramUtils.Visible = True
      Exit Sub
   End Sub

   Private Sub Option3Button_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Option3Button.Click
      PFMCOption = 3
      TransferCoweemanData()
      Me.Close()
      FVS_FramUtils.Visible = True
      Exit Sub
   End Sub

   Private Sub CowCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CowCancelButton.Click
      PFMCOption = 0
      Me.Close()
      FVS_FramUtils.Visible = True
      Exit Sub
   End Sub

   Public Sub TransferCoweemanData()

      Dim OpenCOWspreadsheet As New OpenFileDialog()
      Dim COWSpreadSheetName As String

      OpenCOWspreadsheet.Filter = "Spreadsheets (*.xls*)|*.xls*|All files (*.*)|*.*"
      OpenCOWspreadsheet.FilterIndex = 1
      OpenCOWspreadsheet.RestoreDirectory = True
      OpenCOWspreadsheet.Title = "Select Coweeman SpreadSheet"

      COWSpreadSheetName = ""
      If OpenCOWspreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            COWSpreadSheetName = OpenCOWspreadsheet.FileName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file selected. Original error: " & Ex.Message)
         End Try
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

      '- Test if Coweeman Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(COWSpreadSheetName).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(COWSpreadSheetName)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '- Find WorkSheets for Coweeman Options
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 12 Then
            If PFMCOption = 1 Then
               If xlWorkSheet.Name = "PFMC-Option-1" Then Exit For
            ElseIf PFMCOption = 2 Then
               If xlWorkSheet.Name = "PFMC-Option-2" Then Exit For
            ElseIf PFMCOption = 3 Then
               If xlWorkSheet.Name = "PFMC-Option-3" Then Exit For
            End If
         End If
      Next

      '- Check if DataBase contains Coweeman Worksheets
      If xlWorkSheet.Name.Substring(0, 12) <> "PFMC-Option-" Then
         MsgBox("Can't Find 'PFMC-Option-?' WorkSheet in your Selection" & vbCrLf & _
                "Please Choose appropriate SpreadSheet with Coweeman WorkSheets!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      ''- Check first Fishery Name for correct Species Spreadsheet
      'Dim testname As String
      'testname = xlWorkSheet.Range("A4").Value
      'If SpeciesName = "CHINOOK" Then
      '   If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
      '      MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your DataBase Selection" & vbCrLf & _
      '             "Please Choose appropriate CHINOOK DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
      '      GoTo CloseExcelWorkBook
      '   End If
      'ElseIf SpeciesName = "COHO" Then
      '   If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
      '      MsgBox("Can't Find 'No Cal Trm' as first Fishery your DataBase Selection" & vbCrLf & _
      '             "Please Choose appropriate COHO DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
      '      GoTo CloseExcelWorkBook
      '   End If
      'End If

      Me.Cursor = Cursors.WaitCursor
      Call SumCoweemanData()
      Me.Cursor = Cursors.Default


CloseExcelWorkBook:
      '- Done with FRAM-Template WorkBook .. Close and release object
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
      'xlApp.Application.DisplayAlerts = True
      'xlApp = Nothing

      xlApp.Visible = True
      xlApp.Application.DisplayAlerts = True
      Me.Cursor = Cursors.Default

   End Sub

   Sub SumCoweemanData()

      '- Set Dimension for Fishery Summary Arrays (First element is # Fisheries)
      Dim SEAK() As Integer = {3, 1, 2, 3}
      Dim Canada() As Integer = {12, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}
      Dim PFMCAll() As Integer = {15, 16, 17, 18, 20, 21, 22, 26, 27, 29, 30, 31, 32, 33, 34, 35}
      Dim NTOnly() As Integer = {12, 16, 18, 20, 22, 26, 27, 30, 31, 32, 33, 34, 35}
      Dim SOF() As Integer = {6, 30, 31, 32, 33, 34, 35}
      Dim PS() As Integer = {2, 36, 71}
      Dim PSTRFish() As Integer = {14, 38, 40, 41, 44, 47, 50, 52, 55, 59, 61, 63, 66, 69, 71}
      Dim WAC() As Integer = {4, 19, 23, 24, 25}
      Dim TRShare(10, 2, 2) As Double
      Dim Stock, Fishery, BY, TRFish As Integer
      Dim SumAllCatch, SumMark, SumUnMarked As Double
      Dim CellAddress As String
      Dim PrnLine As String
      Dim NonTR As Boolean
      Dim PFMC_Sport(2, 2) As Double



      File_Name = FVSdatabasepath & "\CoweemanCheck.Txt"
      If Exists(File_Name) Then Delete(File_Name)
      sw = CreateText(File_Name)
      PrnLine = "Command File =" + FVSdatabasepath + "\" & RunIDNameSelect.ToString & "     " & Date.Today.ToString
      sw.WriteLine(PrnLine)
      sw.WriteLine(" ")

      '- Call BYERReport to Fill BY Arrays
      OptionChinookBYAEQ = 2
      BYERReport()

      xlWorkSheet.Range("E1").Value = RunIDNameSelect
      xlWorkSheet.Range("E2").Value = RunIDRunTimeDateSelect.ToString
      BY = 2
      For Stock = 1 To 10
         If Stock <> 10 Then
            Stk = Stock * 2 + 35
         Else
            Stk = 67 '- Coweeman is a little out-of-order from base period
         End If

         '- SEAK Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To SEAK(0)
            Fish = SEAK(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "C" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "C" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "C" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- SEAK Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To SEAK(0)
            Fish = SEAK(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "C" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "C" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "C" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Canada Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To Canada(0)
            Fish = Canada(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  If TerminalFisheryFlag(Fish, TStep) = Term Then
                     SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                     SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep))
                     SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                     SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep))
                  Else
                     SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  End If
               Next
            Next
         Next
         CellAddress = "D" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "D" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "D" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Canada Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To Canada(0)
            Fish = Canada(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  If TerminalFisheryFlag(Fish, TStep) = Term Then
                     SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep))
                     SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep))
                     SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep))
                     SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep))
                  Else
                     SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  End If
               Next
            Next
         Next
         CellAddress = "D" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "D" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "D" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- PFMC Total Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To PFMCAll(0)
            Fish = PFMCAll(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  If Fish = 17 Or Fish = 21 Or Fish = 29 Then
                     '- Treaty PFMC
                     TRShare(Stock, 2, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     TRShare(Stock, 2, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  Else
                     '- Non-Treaty PFMC
                     TRShare(Stock, 1, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     TRShare(Stock, 1, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  End If

                  '- NoF & SoF Sport Sums
                  If Stock = 10 Then
                     If Fish = 18 Or Fish = 22 Or Fish = 27 Then
                        PFMC_Sport(1, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        PFMC_Sport(1, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     End If
                     If Fish = 31 Or Fish = 33 Or Fish = 35 Then
                        PFMC_Sport(1, 2) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        PFMC_Sport(1, 2) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     End If
                  End If


               Next
            Next
         Next
         CellAddress = "E" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "E" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "E" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
         CellAddress = "T" & "21"
         xlWorkSheet.Range(CellAddress).Value = PFMC_Sport(1, 1)
         CellAddress = "T" & "22"
         xlWorkSheet.Range(CellAddress).Value = PFMC_Sport(1, 2)


         '- PFMC Total Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To PFMCAll(0)
            Fish = PFMCAll(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  If Fish = 17 Or Fish = 21 Or Fish = 29 Then
                     '- Treaty PFMC
                     TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  Else
                     '- Non-Treaty PFMC
                     TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  End If

                  '- NoF & SoF Sport Sums
                  If Stock = 10 Then
                     If Fish = 18 Or Fish = 22 Or Fish = 27 Then
                        PFMC_Sport(2, 1) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        PFMC_Sport(2, 1) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     End If
                     If Fish = 31 Or Fish = 33 Or Fish = 35 Then
                        PFMC_Sport(2, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                        PFMC_Sport(2, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                     End If
                  End If

               Next
            Next
         Next
         CellAddress = "E" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "E" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "E" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
         CellAddress = "T" & "120"
         xlWorkSheet.Range(CellAddress).Value = PFMC_Sport(2, 1)
         CellAddress = "T" & "121"
         xlWorkSheet.Range(CellAddress).Value = PFMC_Sport(2, 2)

         '- PFMC NonTreaty Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To NTOnly(0)
            Fish = NTOnly(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "F" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "F" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "F" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- PFMC NonTreaty Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To NTOnly(0)
            Fish = NTOnly(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "F" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "F" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "F" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- PFMC SoF Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To SOF(0)
            Fish = SOF(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '- Non-Treaty PFMC
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "G" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "G" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "G" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- PFMC SoF Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To SOF(0)
            Fish = SOF(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '- Non-Treaty PFMC
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
               Next
            Next
         Next
         CellAddress = "G" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "G" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "G" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Puget Sound Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         'Dim FYDiff As Double
         For Fish As Integer = PS(1) To PS(2)
            NonTR = True
            For TRFish = 1 To PSTRFish(0)
               If Fish = PSTRFish(TRFish) Then
                  NonTR = False
                  Exit For
               End If
            Next
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  'If TerminalFisheryFlag(Fish, TStep) = Term Then
                  '   SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                  '   SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep))
                  '   SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep))
                  '   SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep))
                  'Else
                  '   SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '   SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '   SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '   SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  'End If
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-If NonTR = False Then
                  '- Treaty PS
                  '-TRShare(Stock, 2, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 2, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-Else
                  '- Non-Treaty PS
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-End If
               Next
            Next
         Next
         CellAddress = "H" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "H" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "H" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Puget Sound Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fish As Integer = PS(1) To PS(2)
            NonTR = True
            For TRFish = 1 To PSTRFish(0)
               If Fish = PSTRFish(TRFish) Then
                  NonTR = False
                  Exit For
               End If
            Next
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  'If TerminalFisheryFlag(Fish, TStep) = Term Then
                  '   SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep))
                  '   SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep))
                  '   SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep))
                  '   SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep))
                  'Else
                  '   SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '   SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '   SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '   SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  'End If
                  'End If
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep) * AEQ(Stk + 1, Age, TStep)
                  '-If NonTR = False Then
                  '- Treaty PS
                  '-TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-Else
                  '- Non-Treaty PS
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-End If
               Next
            Next
         Next
         CellAddress = "H" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "H" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "H" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Wash Coast Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To WAC(0)
            Fish = WAC(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-If Fish = 19 Or Fish = 24 Then
                  '- Treaty WAC
                  '-TRShare(Stock, 2, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 2, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-Else
                  '- Non-Treaty WAC
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 1, 1) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-End If
               Next
            Next
         Next
         CellAddress = "I" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "I" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "I" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Wash Coast Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 1 To WAC(0)
            Fish = WAC(Fishery)
            For TStep As Integer = 1 To 3
               For Age As Integer = MinAge To MaxAge
                  SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-If Fish = 19 Or Fish = 24 Then
                  '- Treaty WAC
                  '-TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 2, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-Else
                  '- Non-Treaty WAC
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFShakers(BY, Stk, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk, Age, Fish, TStep) + BYMSFDropOff(BY, Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                  '-TRShare(Stock, 1, 2) += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFShakers(BY, Stk + 1, Age, Fish, TStep) + BYMSFNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYMSFDropOff(BY, Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  '-End If
               Next
            Next
         Next
         CellAddress = "I" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "I" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "I" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Terminal Run Fishing Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 71 To 73
            If Fishery = 71 Then
               Fish = 28  '- Columbia River Net
            Else
               Fish = Fishery  '- FWSport and FWNet
            End If
            '- Terminal Fishery Time Steps are Stock Specific
            'If Stk = 49 Or Stk = 51 Then
            '   TStep = 1
            '   For Age = 3 To MaxAge
            '      SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '      SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '   Next
            'ElseIf Stk = 45 Then
            '   TStep = 2
            '   For Age = 3 To MaxAge
            '      SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '      SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '   Next
            '   TStep = 3
            '   For Age = 3 To MaxAge
            '      SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '      SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '   Next
            'Else
            '   TStep = 3
            '   For Age = 3 To MaxAge
            '      SumAllCatch += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumAllCatch += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '      SumUnMarked += (LandedCatch(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + LegalShakers(Stk, Age, Fish, TStep))
            '      SumMark += (LandedCatch(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + LegalShakers(Stk + 1, Age, Fish, TStep))
            '   Next
            'End If
            If Stk = 49 Or Stk = 51 Then
               TStep = 1
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumAllCatch += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                  SumUnMarked += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumMark += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
               Next
            ElseIf Stk = 45 Then
               TStep = 2
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumAllCatch += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                  SumUnMarked += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumMark += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
               Next
               TStep = 3
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumAllCatch += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                  SumUnMarked += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumMark += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
               Next
            Else
               TStep = 3
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumAllCatch += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
                  SumUnMarked += LandedCatch(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep)
                  SumMark += LandedCatch(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep)
               Next
            End If
         Next
         '- Add Model Escapement
         For TStep As Integer = 1 To 3
            For Age As Integer = 3 To MaxAge
               SumAllCatch += Escape(Stk, Age, TStep)
               SumAllCatch += Escape(Stk + 1, Age, TStep)
               SumUnMarked += Escape(Stk, Age, TStep)
               SumMark += Escape(Stk + 1, Age, TStep)
            Next
         Next

         CellAddress = "J" & CStr(Stock + 6)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "J" & CStr(Stock + 39)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "J" & CStr(Stock + 72)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked

         '- Terminal Run Brood Year
         SumAllCatch = 0
         SumMark = 0
         SumUnMarked = 0
         For Fishery = 71 To 73
            If Fishery = 71 Then
               Fish = 28  '- Columbia River Net
            Else
               Fish = Fishery  '- FWSport and FWNet
            End If
            '- Terminal Fishery Time Steps are Stock Specific
            If Stk = 49 Or Stk = 51 Then
               TStep = 1
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumAllCatch += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
                  SumUnMarked += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumMark += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
               Next
            ElseIf Stk = 45 Then
               TStep = 2
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumAllCatch += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
                  SumUnMarked += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumMark += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
               Next
               TStep = 3
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumAllCatch += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
                  SumUnMarked += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumMark += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
               Next
            Else
               TStep = 3
               For Age As Integer = 3 To MaxAge
                  SumAllCatch += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumAllCatch += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
                  SumUnMarked += BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk, Age, Fish, TStep)
                  SumMark += BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYMSFLandedCatch(BY, Stk + 1, Age, Fish, TStep)
               Next
            End If
            'For TStep = 1 To 3
            '   For Age = 3 To MaxAge
            '      SumAllCatch += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYLegalShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep))
            '      SumAllCatch += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYLegalShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep))
            '      SumUnMarked += (BYLandedCatch(BY, Stk, Age, Fish, TStep) + BYShakers(BY, Stk, Age, Fish, TStep) + BYLegalShakers(BY, Stk, Age, Fish, TStep) + BYNonRetention(BY, Stk, Age, Fish, TStep) + BYDropOff(BY, Stk, Age, Fish, TStep))
            '      SumMark += (BYLandedCatch(BY, Stk + 1, Age, Fish, TStep) + BYShakers(BY, Stk + 1, Age, Fish, TStep) + BYLegalShakers(BY, Stk + 1, Age, Fish, TStep) + BYNonRetention(BY, Stk + 1, Age, Fish, TStep) + BYDropOff(BY, Stk + 1, Age, Fish, TStep))
            '   Next
            'Next
         Next
         '- Add Model BY-Escapement
         For TStep As Integer = 1 To 3
            For Age As Integer = 3 To MaxAge
               SumAllCatch += BYEscape(BY, Stk, Age, TStep)
               SumAllCatch += BYEscape(BY, Stk + 1, Age, TStep)
               SumUnMarked += BYEscape(BY, Stk, Age, TStep)
               SumMark += BYEscape(BY, Stk + 1, Age, TStep)
            Next
         Next

         CellAddress = "J" & CStr(Stock + 105)
         xlWorkSheet.Range(CellAddress).Value = SumAllCatch
         CellAddress = "J" & CStr(Stock + 138)
         xlWorkSheet.Range(CellAddress).Value = SumMark
         CellAddress = "J" & CStr(Stock + 171)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked


      '-****************************************************************
      '- Columbia River Summer Chinook summary table for Cindy Le Fleur


      Dim SumSum(1, 75) As Double
      Dim Cindy(1, 15) As Double
      Dim Pete As Integer
      Dim Ops() As String = {"C", "D", "E"}
      Dim S_name As Excel.Worksheet


         If Stock = 5 Then

            For Pete = 1 To 73
                  Fish = Pete
               For TStep As Integer = 1 To 3
                  For Age As Integer = MinAge To MaxAge
                     SumSum(1, Pete) += (LandedCatch(Stk, Age, Fish, TStep) + NonRetention(Stk, Age, Fish, TStep) + Shakers(Stk, Age, Fish, TStep) + DropOff(Stk, Age, Fish, TStep) + MSFLandedCatch(Stk, Age, Fish, TStep) + MSFNonRetention(Stk, Age, Fish, TStep) + MSFShakers(Stk, Age, Fish, TStep) + MSFDropOff(Stk, Age, Fish, TStep)) * AEQ(Stk, Age, TStep)
                     SumSum(1, Pete) += (LandedCatch(Stk + 1, Age, Fish, TStep) + NonRetention(Stk + 1, Age, Fish, TStep) + Shakers(Stk + 1, Age, Fish, TStep) + DropOff(Stk + 1, Age, Fish, TStep) + MSFLandedCatch(Stk + 1, Age, Fish, TStep) + MSFNonRetention(Stk + 1, Age, Fish, TStep) + MSFShakers(Stk + 1, Age, Fish, TStep) + MSFDropOff(Stk + 1, Age, Fish, TStep)) * AEQ(Stk + 1, Age, TStep)
                  Next
               Next
            Next

            Cindy(1, 1) = SumSum(1, 27) '-Area 1 Sport
            Cindy(1, 2) = SumSum(1, 26) '-Area 1 Troll
            Cindy(1, 3) = SumSum(1, 22) '- Area 2 Sport
            Cindy(1, 4) = SumSum(1, 20) '-Area 2 Troll
            Cindy(1, 5) = SumSum(1, 18) '-Area 3:4 Sport
            Cindy(1, 6) = SumSum(1, 16) ' Area 3:4 Troll
            Cindy(1, 7) = SumSum(1, 21) + SumSum(1, 17) + SumSum(1, 41) '-Treaty Troll
            Cindy(1, 8) = SumSum(1, 31) + SumSum(1, 33) + SumSum(1, 35) '-SOF Sport
            Cindy(1, 9) = SumSum(1, 30) + SumSum(1, 32) + SumSum(1, 34) '-SPF Troll
            Cindy(1, 10) = SumSum(1, 23) + SumSum(1, 25) '-WA Coast NT Comm
            Cindy(1, 11) = SumSum(1, 19) + SumSum(1, 24) '-WA Coast Treaty Net
            Cindy(1, 12) = SumSum(1, 36) + SumSum(1, 42) + SumSum(1, 45) + SumSum(1, 48) + SumSum(1, 53) + SumSum(1, 54) + SumSum(1, 56) + SumSum(1, 57) + SumSum(1, 60) + SumSum(1, 62) + SumSum(1, 64) + SumSum(1, 67) '-Puget Snd Sport
            Cindy(1, 13) = SumSum(1, 37) + SumSum(1, 39) + SumSum(1, 43) + SumSum(1, 46) + SumSum(1, 49) + SumSum(1, 51) + SumSum(1, 58) + SumSum(1, 65) + SumSum(1, 68) + SumSum(1, 70) '-Puget Snd NT Comm
            Cindy(1, 14) = SumSum(1, 38) + SumSum(1, 40) + SumSum(1, 41) + SumSum(1, 44) + SumSum(1, 47) + SumSum(1, 50) + SumSum(1, 52) + SumSum(1, 55) + SumSum(1, 59) + SumSum(1, 61) + SumSum(1, 63) + SumSum(1, 66) + SumSum(1, 69) + SumSum(1, 71) '-Puget Snd Treaty


            For Pete = 1 To 9

            CellAddress = CStr(Ops(PFMCOption - 1)) & CStr(31 + Pete)
            S_name = xlWorkBook.Sheets("Cindy's Data")
            S_name.Range(CellAddress).Value = Cindy(1, Pete)

            Next

            For Pete = 10 To 14

            CellAddress = CStr(Ops(PFMCOption - 1)) & CStr(39 + Pete)
            S_name = xlWorkBook.Sheets("Cindy's Data")
            S_name.Range(CellAddress).Value = Cindy(1, Pete)

            Next

         End If

      '- Summer Chinook Summary & Transfer Completed
      '-****************************************************************


      Next

      '- CR Sharing PFMC+River vs Tribal .. Fishing Year and Brood Year
      For Stock = 1 To 10
         '- Fishing Year NonTreaty
         CellAddress = "T" & CStr(Stock + 6)
         SumUnMarked = TRShare(Stock, 1, 1)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
         '- Fishing Year Treaty
         CellAddress = "V" & CStr(Stock + 6)
         SumUnMarked = TRShare(Stock, 2, 1)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
         '- Brood Year NonTreaty
         CellAddress = "T" & CStr(Stock + 105)
         SumUnMarked = TRShare(Stock, 1, 2)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
         '- Brood Year Treaty
         CellAddress = "V" & CStr(Stock + 105)
         SumUnMarked = TRShare(Stock, 2, 2)
         xlWorkSheet.Range(CellAddress).Value = SumUnMarked
      Next


      sw.Close()


   End Sub


   Private Sub FVS_Coweeman_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      FormHeight = 642
      FormWidth = 745
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
         If FVS_Coweeman_ReSize = False Then
            Resize_Form(Me)
            FVS_Coweeman_ReSize = True
         End If
      End If

   End Sub
End Class