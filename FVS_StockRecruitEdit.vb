Imports Microsoft.Office.Interop
Public Class FVS_StockRecruitEdit

   Public GridLoading As Boolean

   Private Sub SRCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SRCancelButton.Click
      Me.Visible = False
      FVS_InputMenu.Visible = True
      FVS_InputMenu.Refresh()
   End Sub

   Private Sub SRDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SRDoneButton.Click

      '- Update Stock Recruit Array and Check for Changes
      If SpeciesName = "COHO" Then
         Age = 3
         For Stk As Integer = 1 To NumStk
            If StockRecruitGrid.Item(3, Stk - 1).Value <> StockRecruit(Stk, Age, 1) Then
               ChangeStockRecruit = True
               StockRecruit(Stk, Age, 1) = StockRecruitGrid.Item(3, Stk - 1).Value
               StockRecruit(Stk, Age, 2) = StockRecruitGrid.Item(4, Stk - 1).Value
            End If
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               If StockRecruitGrid.Item(3, (Stk * 4 - 6) + Age).Value <> StockRecruit(Stk, Age, 1) Then
                  ChangeStockRecruit = True
                  StockRecruit(Stk, Age, 1) = StockRecruitGrid.Item(3, (Stk * 4 - 6) + Age).Value
                  StockRecruit(Stk, Age, 2) = StockRecruitGrid.Item(4, (Stk * 4 - 6) + Age).Value
               End If
            Next
         Next
      End If

      Me.Close()
      FVS_InputMenu.Visible = True

   End Sub

   Private Sub FVS_StockRecruitEdit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 827
      FormHeight = 847
      FormWidth = 1049
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
         If FVS_StockRecruitEdit_ReSize = False Then
            Resize_Form(Me)
            FVS_StockRecruitEdit_ReSize = True
         End If
      End If

      GridLoading = True
      StockRecruitGrid.Columns.Clear()
      StockRecruitGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      StockRecruitGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      StockRecruitGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)

      StockRecruitGrid.Columns.Add("StockName", "Name")
      StockRecruitGrid.Columns("StockName").Width = 400 / FormWidthScaler
      StockRecruitGrid.Columns("StockName").ReadOnly = True
      StockRecruitGrid.Columns("StockName").DefaultCellStyle.BackColor = Color.Aquamarine

      StockRecruitGrid.Columns.Add("Num", "Num")
      StockRecruitGrid.Columns("Num").Width = 50 / FormWidthScaler
      StockRecruitGrid.Columns("Num").ReadOnly = True
      StockRecruitGrid.Columns("Num").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      StockRecruitGrid.Columns("Num").DefaultCellStyle.BackColor = Color.Azure

      StockRecruitGrid.Columns.Add("Age", "Age")
      StockRecruitGrid.Columns("Age").Width = 50 / FormWidthScaler
      StockRecruitGrid.Columns("Age").ReadOnly = True
      StockRecruitGrid.Columns("Age").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      StockRecruitGrid.Columns("Age").DefaultCellStyle.BackColor = Color.Azure

      StockRecruitGrid.Columns.Add("RecScaler", "Recruit Scaler")
      StockRecruitGrid.Columns("RecScaler").Width = 150 / FormWidthScaler
      StockRecruitGrid.Columns("RecScaler").DefaultCellStyle.Format = ("###0.0000")
      StockRecruitGrid.Columns("RecScaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      StockRecruitGrid.Columns.Add("RecSize", "Recruit Cohort")
      StockRecruitGrid.Columns("RecSize").Width = 150 / FormWidthScaler
      StockRecruitGrid.Columns("RecSize").DefaultCellStyle.Format = ("########0")
      StockRecruitGrid.Columns("RecSize").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      StockRecruitGrid.Columns.Add("Version", "Vers")
      StockRecruitGrid.Columns("Version").Width = 50 / FormWidthScaler
      'StockRecruitGrid.Columns("Version").CellTemplate.AdjustCellBorderStyle()
      StockRecruitGrid.Columns("Version").DefaultCellStyle.Format = ("###0")
      StockRecruitGrid.Columns("Version").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomLeft

      If SpeciesName = "COHO" Then
         Age = 3
         StockRecruitGrid.RowCount = NumStk
         For Stk As Integer = 1 To NumStk
            StockRecruitGrid.Item(0, Stk - 1).Value = StockTitle(Stk)
            StockRecruitGrid.Item(1, Stk - 1).Value = Stk.ToString
            StockRecruitGrid.Item(2, Stk - 1).Value = Age.ToString
            StockRecruitGrid.Item(3, Stk - 1).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
            StockRecruitGrid.Item(4, Stk - 1).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
            StockRecruitGrid.Item(5, Stk - 1).Value = StockVersion.ToString
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         StockRecruitGrid.RowCount = NumStk * (MaxAge - MinAge + 1)
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               If Age = MinAge Then
                  StockRecruitGrid.Item(0, (Stk * 4 - 6) + Age).Value = StockName(Stk)
               Else
                  StockRecruitGrid.Item(0, (Stk * 4 - 6) + Age).Value = "----"
               End If
               StockRecruitGrid.Item(0, (Stk * 4 - 6) + Age).Value = StockTitle(Stk)
               StockRecruitGrid.Item(1, (Stk * 4 - 6) + Age).Value = Stk.ToString
               StockRecruitGrid.Item(2, (Stk * 4 - 6) + Age).Value = Age.ToString
               StockRecruitGrid.Item(3, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
               StockRecruitGrid.Item(4, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
               StockRecruitGrid.Item(5, (Stk * 4 - 6) + Age).Value = StockVersion.ToString
            Next
         Next
      End If
      GridLoading = False

   End Sub

   Private Sub StockRecruitGrid_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles StockRecruitGrid.CellValueChanged

      Dim Row, Col As Integer

      If GridLoading = True Then Exit Sub
      Row = StockRecruitGrid.CurrentCell.RowIndex
      Col = StockRecruitGrid.CurrentCell.ColumnIndex
      If SpeciesName = "COHO" Then
         Stk = Row + 1
         Age = 3
      ElseIf SpeciesName = "CHINOOK" Then
         '- Use \ division to return integer (no rounding)
         Stk = (Row \ 4) + 1
         Age = Row - (Stk * 4 - 4) + 2
      End If
      '- User Changed Recruit Scaler ... Update Cohort Size
      If Col = 3 Then
         StockRecruitGrid.Item(Col + 1, Row).Value = StockRecruitGrid.Item(Col, Row).Value * BaseCohortSize(Stk, Age)
      End If
      '- User Changed Forecasted Cohort Size ... Update Scaler
      If Col = 4 Then
         StockRecruitGrid.Item(Col - 1, Row).Value = StockRecruitGrid.Item(Col, Row).Value / BaseCohortSize(Stk, Age)
      End If
   End Sub

   Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      ClipStr = "StockLongName" & vbTab & "StockID" & vbTab & "Age" & vbTab & "RecruitScaleFactor" & vbTab & "RecruitCohortSize" & vbCr
      For RecNum = 0 To StockRecruitGrid.RowCount - 1
         For ColNum = 0 To 4
            If ColNum = 0 Then
               ClipStr = ClipStr & StockRecruitGrid.Item(ColNum, RecNum).Value
            Else
               ClipStr = ClipStr & vbTab & StockRecruitGrid.Item(ColNum, RecNum).Value
            End If
         Next
         ClipStr = ClipStr & vbCr
      Next
      Clipboard.SetDataObject(ClipStr)
   End Sub

   Private Sub ReadRecruitsButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadRecruitsButton.Click

      Dim OpenFileDialog1 As New OpenFileDialog()
      Dim FRAMCatchSpreadSheet, FRAMCatchSpreadSheetPath As String

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenFileDialog1.Filter = "FRAM-Catch Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenFileDialog1.FilterIndex = 1
      OpenFileDialog1.RestoreDirectory = True

      If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         FRAMCatchSpreadSheet = OpenFileDialog1.FileName
         FRAMCatchSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(FRAMCatchSpreadSheet).DirectoryName
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

      '- Test if FRAM-Template Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(FRAMCatchSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(FRAMCatchSpreadSheet)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '- Find WorkSheets with FRAM Catch numbers
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name = "FRAM_Recruits" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAM_Recruits" Then
         MsgBox("Can't Find 'FRAM_Recruits' WorkSheet in your DataBase Selection" & vbCrLf & _
                "Please Choose appropriate DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "UnMarked Nooksack/Samish Fall" Then
            MsgBox("Can't Find 'UnMarked Nooksack/Samish Fall' as first Stock your DataBase Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "Nooksack River Wild UnMarked" Then
            MsgBox("Can't Find 'Nooksack River Wild UnMarked' as first Stock your DataBase Selection" & vbCrLf & _
                   "Please Choose appropriate COHO DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress As String
      Dim FlagAddress As String
      GridLoading = True
      If SpeciesName = "CHINOOK" Then
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               CellAddress = "D" & CStr(Stk * 4 + Age - 2)
               FlagAddress = "E" & CStr(Stk * 4 + Age - 2)
               If IsNumeric(xlWorkSheet.Range(CellAddress).Value) Then
                  StockRecruit(Stk, Age, 1) = CDbl(xlWorkSheet.Range(CellAddress).Value)
                  StockRecruitGrid.Item(3, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
                  StockRecruit(Stk, Age, 2) = StockRecruit(Stk, Age, 1) * BaseCohortSize(Stk, Age)
                  StockRecruitGrid.Item(4, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
               Else
                  StockRecruit(Stk, Age, 1) = 0
                  StockRecruitGrid.Item(3, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
                  StockRecruit(Stk, Age, 2) = 0
                  StockRecruitGrid.Item(4, (Stk * 4 - 6) + Age).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
               End If
            Next
         Next
      ElseIf SpeciesName = "COHO" Then
         For Stk As Integer = 1 To NumStk
            CellAddress = "D" & CStr(Stk + 3)
            FlagAddress = "E" & CStr(Stk + 3)
            If IsNumeric(xlWorkSheet.Range(CellAddress).Value) Then
               StockRecruit(Stk, Age, 1) = CDbl(xlWorkSheet.Range(CellAddress).Value)
               StockRecruitGrid.Item(3, (Stk - 1)).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
               StockRecruit(Stk, Age, 2) = StockRecruit(Stk, Age, 1) * BaseCohortSize(Stk, Age)
               StockRecruitGrid.Item(4, (Stk - 1)).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
            Else
               StockRecruit(Stk, Age, 1) = 0
               StockRecruitGrid.Item(3, (Stk - 1)).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
               StockRecruit(Stk, Age, 2) = 0
               StockRecruitGrid.Item(4, (Stk - 1)).Value = StockRecruit(Stk, Age, 2).ToString("#######0")
            End If
         Next
      End If
      ChangeStockRecruit = True

CloseExcelWorkBook:
      '- Done with FRAM-Template WorkBook .. Close and release object
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

      Me.Cursor = Cursors.Default
      GridLoading = False

      Exit Sub

   End Sub

   Private Sub FillRecruitSSButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FillRecruitSSButton.Click

      Dim OpenFileDialog1 As New OpenFileDialog()
      Dim FRAMCatchSpreadSheet, FRAMCatchSpreadSheetPath As String

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenFileDialog1.Filter = "FRAM-Catch Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenFileDialog1.FilterIndex = 1
      OpenFileDialog1.RestoreDirectory = True

      If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         FRAMCatchSpreadSheet = OpenFileDialog1.FileName
         FRAMCatchSpreadSheetPath = My.Computer.FileSystem.GetFileInfo(FRAMCatchSpreadSheet).DirectoryName
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

      '- Test if FRAM-Template Workbook is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(FRAMCatchSpreadSheet).Name
      For Each xlWorkBook In xlApp.Workbooks
         If xlWorkBook.Name = wbName Then
            xlWorkBook.Activate()
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp.Workbooks.Open(FRAMCatchSpreadSheet)
      xlApp.WindowState = Excel.XlWindowState.xlMinimized
SkipWBOpen:

      xlApp.Application.DisplayAlerts = False
      xlApp.Visible = False
      xlApp.WindowState = Excel.XlWindowState.xlMinimized

      '- Find WorkSheets with FRAM Catch numbers
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name = "FRAM_Recruits" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAM_Recruits" Then
         MsgBox("Can't Find 'FRAM_Recruits' WorkSheet in your DataBase Selection" & vbCrLf & _
                "Please Choose appropriate DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "UnMarked Nooksack/Samish Fall" Then
            MsgBox("Can't Find 'UnMarked Nooksack/Samish Fall' as first Stock your DataBase Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "Nooksack River Wild UnMarked" Then
            MsgBox("Can't Find 'Nooksack River Wild UnMarked' as first Stock your DataBase Selection" & vbCrLf & _
                   "Please Choose appropriate COHO DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress As String
      Dim FlagAddress As String
      GridLoading = True
      If SpeciesName = "CHINOOK" Then
         For Stk As Integer = 1 To NumStk
            For Age As Integer = MinAge To MaxAge
               CellAddress = "D" & CStr(Stk * 4 + Age - 2)
               FlagAddress = "E" & CStr(Stk * 4 + Age - 2)
               xlWorkSheet.Range(CellAddress).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
               xlWorkSheet.Range(FlagAddress).Value = StockRecruit(Stk, Age, 2).ToString("########0")
            Next
         Next
      ElseIf SpeciesName = "COHO" Then
         For Stk As Integer = 1 To NumStk
            CellAddress = "D" & CStr(Stk + 3)
            FlagAddress = "E" & CStr(Stk + 3)
            xlWorkSheet.Range(CellAddress).Value = StockRecruit(Stk, Age, 1).ToString("###0.0000")
            xlWorkSheet.Range(FlagAddress).Value = StockRecruit(Stk, Age, 2).ToString("########0")
         Next
      End If
      ChangeStockRecruit = True

CloseExcelWorkBook:
      '- Done with FRAM-Template WorkBook .. Close and release object
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

      Me.Cursor = Cursors.Default
      GridLoading = False

      Exit Sub


   End Sub

End Class