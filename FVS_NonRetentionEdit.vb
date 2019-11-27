Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class FVS_NonRetentionEdit

   Private Sub FVS_NonRetentionEdit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 977
      FormHeight = 997
      FormWidth = 1138
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
         If FVS_NonRetentionEdit_ReSize = False Then
            Resize_Form(Me)
            FVS_NonRetentionEdit_ReSize = True
         End If
      End If

      If SpeciesName = "COHO" Then
         '- Hide Chinook CNR Flag/Value labels for COHO
         Label2.Text = "COHO CNR Estimates are Total Dead Fish"
         Label3.Visible = False
         Label4.Visible = False
         Label5.Visible = False
         Label6.Visible = False
         Label7.Visible = False
         Label8.Visible = False
         Label9.Visible = False
         Label10.Visible = False
         Label11.Visible = False
         Label12.Visible = False
         Label13.Visible = False
         Label14.Visible = False
         Label15.Visible = False
         Label16.Visible = False
         Label17.Visible = False
         Label18.Visible = False
         Label19.Visible = False
         Label20.Visible = False
         Label21.Visible = False
         Label22.Visible = False
         Label23.Visible = False
         Label24.Visible = False
         Label25.Visible = False
      ElseIf SpeciesName = "CHINOOK" Then
         Label2.Text = "1 = Computed CNR"
         Label3.Visible = True
         Label4.Visible = True
         Label5.Visible = True
         Label6.Visible = True
         Label7.Visible = True
         Label8.Visible = True
         Label9.Visible = True
         Label10.Visible = True
         Label11.Visible = True
         Label12.Visible = True
         Label13.Visible = True
         Label14.Visible = True
         Label15.Visible = True
         Label16.Visible = True
         Label17.Visible = True
         Label18.Visible = True
         Label19.Visible = True
         Label20.Visible = True
         Label21.Visible = True
         Label22.Visible = True
         Label23.Visible = True
         Label24.Visible = True
         Label25.Visible = True
      End If

      '- Fill the DataGrid with Values ... COHO and CHINOOK are different
      NonRetentionGrid.Columns.Clear()
      NonRetentionGrid.Rows.Clear()
      NonRetentionGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      NonRetentionGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      NonRetentionGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      If SpeciesName = "COHO" Then
         NonRetentionGrid.Columns.Add("FisheryName", "Name")
            NonRetentionGrid.Columns("FisheryName").Width = 120 / FormWidthScaler
         NonRetentionGrid.Columns("FisheryName").ReadOnly = True
         NonRetentionGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine
         NonRetentionGrid.Columns.Add("FishNum", "#")
         NonRetentionGrid.Columns("FishNum").Width = 40 / FormWidthScaler
         NonRetentionGrid.Columns("FishNum").ReadOnly = True
         NonRetentionGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

         NonRetentionGrid.Columns.Add("Time1Estimate", "Jan-June")
         NonRetentionGrid.Columns("Time1Estimate").Width = 150 / FormWidthScaler
         NonRetentionGrid.Columns("Time1Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time2Estimate", "July")
         NonRetentionGrid.Columns("Time2Estimate").Width = 150 / FormWidthScaler
         NonRetentionGrid.Columns("Time2Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time3Estimate", "August")
         NonRetentionGrid.Columns("Time3Estimate").Width = 150 / FormWidthScaler
         NonRetentionGrid.Columns("Time3Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time4Estimate", "September")
         NonRetentionGrid.Columns("Time4Estimate").Width = 150 / FormWidthScaler
         NonRetentionGrid.Columns("Time4Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time5Estimate", "Oct-Dec")
         NonRetentionGrid.Columns("Time5Estimate").Width = 150 / FormWidthScaler
            NonRetentionGrid.Columns("Time5Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.RowCount = NumFish

         For Fish As Integer = 1 To NumFish
            NonRetentionGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
            NonRetentionGrid.Item(1, Fish - 1).Value = Fish.ToString
            For TStep As Integer = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 1 Then
                  NonRetentionGrid.Item(TStep + 1, Fish - 1).Value = NonRetentionInput(Fish, TStep, 1)
               Else
                  NonRetentionGrid.Item(TStep + 1, Fish - 1).Value = "****"
                  NonRetentionGrid.Item(TStep + 1, Fish - 1).Style.BackColor = Color.LightBlue
               End If
            Next
         Next

      ElseIf SpeciesName = "CHINOOK" Then

         NonRetentionGrid.Columns.Add("FisheryName", "Name")
         NonRetentionGrid.Columns("FisheryName").Width = 200 / FormWidthScaler
         NonRetentionGrid.Columns("FisheryName").ReadOnly = True
         NonRetentionGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine
         NonRetentionGrid.Columns.Add("FishNum", "Values")
         NonRetentionGrid.Columns("FishNum").Width = 80 / FormWidthScaler
         NonRetentionGrid.Columns("FishNum").ReadOnly = True
         NonRetentionGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

         NonRetentionGrid.Columns.Add("Time1Flag", "Flg-1")
         NonRetentionGrid.Columns("Time1Flag").Width = 50 / FormWidthScaler
         NonRetentionGrid.Columns("Time1Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         NonRetentionGrid.Columns.Add("Time1Estimate", "Oct-Apr-1")
         NonRetentionGrid.Columns("Time1Estimate").Width = 100 / FormWidthScaler
         NonRetentionGrid.Columns("Time1Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time2Flag", "Flg-2")
         NonRetentionGrid.Columns("Time2Flag").Width = 50 / FormWidthScaler
         NonRetentionGrid.Columns("Time2Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         NonRetentionGrid.Columns.Add("Time2Estimate", "May-June")
         NonRetentionGrid.Columns("Time2Estimate").Width = 100 / FormWidthScaler
         NonRetentionGrid.Columns("Time2Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time3Flag", "Flg-3")
         NonRetentionGrid.Columns("Time3Flag").Width = 50 / FormWidthScaler
         NonRetentionGrid.Columns("Time3Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         NonRetentionGrid.Columns.Add("Time3Estimate", "July-Sept")
         NonRetentionGrid.Columns("Time3Estimate").Width = 100 / FormWidthScaler
         NonRetentionGrid.Columns("Time3Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.Columns.Add("Time4Flag", "Flg-4")
         NonRetentionGrid.Columns("Time4Flag").Width = 50 / FormWidthScaler
         NonRetentionGrid.Columns("Time4Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         NonRetentionGrid.Columns.Add("Time4Estimate", "Oct-Apr-2")
         NonRetentionGrid.Columns("Time4Estimate").Width = 100 / FormWidthScaler
         NonRetentionGrid.Columns("Time4Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         NonRetentionGrid.RowCount = NumFish * 4

         For Fish As Integer = 1 To NumFish
            NonRetentionGrid.Item(0, Fish * 4 - 4).Value = FisheryTitle(Fish)
            NonRetentionGrid.Item(1, Fish * 4 - 4).Value = "Value-1"
            NonRetentionGrid.Item(1, Fish * 4 - 3).Value = "Value-2"
            NonRetentionGrid.Item(1, Fish * 4 - 2).Value = "Value-3"
            NonRetentionGrid.Item(1, Fish * 4 - 1).Value = "Value-4"
            For TStep As Integer = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 1 Then
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value = NonRetentionFlag(Fish, TStep)
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 3).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 3).Style.BackColor = Color.LightSalmon
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 2).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 2).Style.BackColor = Color.LightSalmon
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 1).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 1).Style.BackColor = Color.LightSalmon

                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value = NonRetentionInput(Fish, TStep, 1)
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value = NonRetentionInput(Fish, TStep, 2)
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value = NonRetentionInput(Fish, TStep, 3)
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value = NonRetentionInput(Fish, TStep, 4)
               Else
                  'NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value = NonRetentionFlag(Fish, 1)
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 3).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 3).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 2).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 2).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 1).Value = "*"
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 1).Style.BackColor = Color.LightBlue

                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value = "****"
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value = "****"
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value = "****"
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Style.BackColor = Color.LightBlue
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value = "****"
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Style.BackColor = Color.LightBlue
               End If
            Next

         Next
      End If
   End Sub

   Private Sub NRCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NRCancelButton.Click
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub NRDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NRDoneButton.Click

      'Dim ChangeNonRetention As Boolean

      '- Put Grid Values into NonRetention Arrays  
      'ChangeNonRetention = False
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 0 Then
               If SpeciesName = "COHO" Then
                  If NonRetentionInput(Fish, TStep, 1) <> 0 Then
                     NonRetentionInput(Fish, TStep, 1) = 0
                     NonRetentionFlag(Fish, TStep) = 0
                     ChangeNonRetention = True
                  End If
                  GoTo NextTStep3
               End If
               If NonRetentionFlag(Fish, TStep) <> 0 Then
                  NonRetentionFlag(Fish, TStep) = 0
                  ChangeNonRetention = True
               End If
               If NonRetentionInput(Fish, TStep, 1) <> 0 Then
                  NonRetentionInput(Fish, TStep, 1) = 0
                  ChangeNonRetention = True
               End If
               If NonRetentionInput(Fish, TStep, 2) <> 0 Then
                  NonRetentionInput(Fish, TStep, 2) = 0
                  ChangeNonRetention = True
               End If
               If NonRetentionInput(Fish, TStep, 3) <> 0 Then
                  NonRetentionInput(Fish, TStep, 3) = 0
                  ChangeNonRetention = True
               End If
               If NonRetentionInput(Fish, TStep, 4) <> 0 Then
                  NonRetentionInput(Fish, TStep, 4) = 0
                  ChangeNonRetention = True
               End If
               GoTo NextTStep3
            End If
            If SpeciesName = "COHO" Then
               If CDbl(NonRetentionGrid.Item(TStep + 1, Fish - 1).Value) <> NonRetentionInput(Fish, TStep, 1) Then
                  NonRetentionInput(Fish, TStep, 1) = CDbl(NonRetentionGrid.Item(TStep + 1, Fish - 1).Value)
                  ChangeNonRetention = True
                  If NonRetentionInput(Fish, TStep, 1) = 0 Then
                     NonRetentionFlag(Fish, TStep) = 0
                  Else
                     NonRetentionFlag(Fish, TStep) = 1
                  End If
               End If
            ElseIf SpeciesName = "CHINOOK" Then
               '- Check if Flag Value Changed
               If CInt(NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value) <> NonRetentionFlag(Fish, TStep) Then
                  NonRetentionFlag(Fish, TStep) = CInt(NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value)
                  If Not (NonRetentionFlag(Fish, TStep) <= 4 And NonRetentionFlag(Fish, TStep) >= 0) Then
                     MsgBox("ERROR - Chinook CNR Flag must be Zero, 1, 2, 3, or 4" & vbCrLf & "Check Fish,TStep=" & FisheryName(Fish) & "-" & TStep.ToString, MsgBoxStyle.OkOnly)
                     Exit Sub
                  End If
                  ChangeNonRetention = True
               End If
               '- Check for CNR Input Value Changes
               If CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value) <> NonRetentionInput(Fish, TStep, 1) Then
                  NonRetentionInput(Fish, TStep, 1) = CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value)
                  ChangeNonRetention = True
               End If
               If CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value) <> NonRetentionInput(Fish, TStep, 2) Then
                  NonRetentionInput(Fish, TStep, 2) = CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value)
                  ChangeNonRetention = True
               End If
               If CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value) <> NonRetentionInput(Fish, TStep, 3) Then
                  NonRetentionInput(Fish, TStep, 3) = CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value)
                  ChangeNonRetention = True
               End If
               If CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value) <> NonRetentionInput(Fish, TStep, 4) Then
                  NonRetentionInput(Fish, TStep, 4) = CDbl(NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value)
                  ChangeNonRetention = True
               End If
               '- Check if CNR Input Values Exist when Flag is NonZero ... If Not Zero Flag
               If NonRetentionFlag(Fish, TStep) <> 0 And NonRetentionInput(Fish, TStep, 1) = 0 And NonRetentionInput(Fish, TStep, 2) = 0 _
                  And NonRetentionInput(Fish, TStep, 3) = 0 And NonRetentionInput(Fish, TStep, 4) = 0 Then
                  NonRetentionFlag(Fish, TStep) = 0
               End If
            End If
NextTStep3:
         Next
      Next

      Me.Close()
      FVS_InputMenu.Visible = True

   End Sub


   Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim JimStr As String
      Dim RecNum, ColNum, Row As Integer

      '- The Clipboard Copy Column Names are specific for each of the Species 
      If SpeciesName = "CHINOOK" Then
         '- Clipboard Copy for CHINOOK Non-Retention Screen
         ClipStr = ""
         Clipboard.Clear()
         ClipStr = "FisheryName" & vbTab & "Values" & vbTab & "Flg-1" & vbTab & "Oct-Apr-1" & vbTab & "Flg-2" & vbTab & "May-June" & vbTab & "Flg-3" & vbTab & "July-Sept" & vbTab & "Flg-4" & vbTab & "Oct-Apr-2" & vbCr
         For RecNum = 0 To NumFish - 1
            For Row = 0 To 3
               For ColNum = 0 To 9
                  If ColNum = 0 Then
                     ClipStr = ClipStr & NonRetentionGrid.Item(ColNum, RecNum * 4 + Row).Value
                  Else
                     ClipStr = ClipStr & vbTab & NonRetentionGrid.Item(ColNum, RecNum * 4 + Row).Value
                  End If
               Next
               ClipStr = ClipStr & vbCr
            Next
         Next
         Clipboard.SetDataObject(ClipStr)
      ElseIf SpeciesName = "COHO" Then
         '- Clipboard Copy for COHO Non-Retention Screen
         ClipStr = ""
         Clipboard.Clear()
         ClipStr = "Name" & vbTab & "#" & vbTab & "Jan-June" & vbTab & "July" & vbTab & "August" & vbTab & "Septmbr" & vbTab & "Oct-Dec" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To 6
               JimStr = NonRetentionGrid.Item(ColNum, RecNum).Value
               If ColNum = 0 Then
                  ClipStr = ClipStr & NonRetentionGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & NonRetentionGrid.Item(ColNum, RecNum).Value
               End If
            Next
            ClipStr = ClipStr & vbCr
         Next
         Clipboard.SetDataObject(ClipStr)
      End If

   End Sub

   Private Sub ZeroNRButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ZeroNRButton.Click

      If SpeciesName = "COHO" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 1 Then
                  NonRetentionGrid.Item(TStep + 1, Fish - 1).Value = 0
               End If
            Next
         Next
      ElseIf SpeciesName = "CHINOOK" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               If AnyBaseRate(Fish, TStep) = 1 Then
                  NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value = 0
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value = 0
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value = 0
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value = 0
                  NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value = 0
               End If
            Next
         Next
      End If

   End Sub

   Private Sub LoadNRButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadNRButton.Click

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
            If xlWorkSheet.Name = "FRAM_CNR" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAM_CNR" Then
         MsgBox("Can't Find 'FRAM_CNR' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
            MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
            MsgBox("Can't Find 'No Cal Trm' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate COHO Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress, NewAddress As String
      Dim FlagAddress As String
      If SpeciesName = "CHINOOK" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               CellAddress = ""
               FlagAddress = ""
               If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextNRVal
               Select Case TStep
                  Case 1
                     FlagAddress = "C" & CStr(Fish * 4)
                     CellAddress = "D"
                  Case 2
                     FlagAddress = "E" & CStr(Fish * 4)
                     CellAddress = "F"
                  Case 3
                     FlagAddress = "G" & CStr(Fish * 4)
                     CellAddress = "H"
                  Case 4
                     FlagAddress = "I" & CStr(Fish * 4)
                     CellAddress = "J"
                  Case 5
                     FlagAddress = "K" & CStr(Fish * 4)
                     CellAddress = "L"
               End Select
               If IsNumeric(xlWorkSheet.Range(FlagAddress).Value) Then
                  If CInt(xlWorkSheet.Range(FlagAddress).Value) < 0 Or CInt(xlWorkSheet.Range(FlagAddress).Value) > 4 Then GoTo NextNRVal
                  If IsNumeric(xlWorkSheet.Range(FlagAddress).Value) Then
                     NonRetentionGrid.Item(TStep * 2, Fish * 4 - 4).Value = xlWorkSheet.Range(FlagAddress).Value
                     NewAddress = CellAddress & CStr(Fish * 4)
                     NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 4).Value = xlWorkSheet.Range(NewAddress).Value
                     NewAddress = CellAddress & CStr(Fish * 4 + 1)
                     NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 3).Value = xlWorkSheet.Range(NewAddress).Value
                     NewAddress = CellAddress & CStr(Fish * 4 + 2)
                     NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 2).Value = xlWorkSheet.Range(NewAddress).Value
                     NewAddress = CellAddress & CStr(Fish * 4 + 3)
                     NonRetentionGrid.Item(TStep * 2 + 1, Fish * 4 - 1).Value = xlWorkSheet.Range(NewAddress).Value
                  End If
               End If
NextNRVal:
            Next
         Next
      ElseIf SpeciesName = "COHO" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               CellAddress = ""
               FlagAddress = ""
               If AnyBaseRate(Fish, TStep) = 0 Then GoTo NextNRVal2
               Select Case TStep
                  Case 1
                     CellAddress = "C" & CStr(Fish + 3)
                  Case 2
                     CellAddress = "D" & CStr(Fish + 3)
                  Case 3
                     CellAddress = "E" & CStr(Fish + 3)
                  Case 4
                     CellAddress = "F" & CStr(Fish + 3)
                  Case 5
                     CellAddress = "G" & CStr(Fish + 3)
               End Select
               If IsNumeric(xlWorkSheet.Range(CellAddress).Value) Then
                        NonRetentionGrid.Item(TStep + 1, Fish - 1).Value = xlWorkSheet.Range(CellAddress).Value
               End If
NextNRVal2:
            Next
         Next
      End If

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

      Exit Sub

   End Sub

   Private Sub FillNRSSButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FillNRSSButton.Click

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
            If xlWorkSheet.Name = "FRAM_CNR" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAM_CNR" Then
         MsgBox("Can't Find 'FRAM_CNR' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
            MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
            MsgBox("Can't Find 'No Cal Trm' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate COHO Spreadsheet with FRAM CNR WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress, NewAddress As String
      Dim FlagAddress As String
      If SpeciesName = "CHINOOK" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               CellAddress = ""
               FlagAddress = ""
               Select Case TStep
                  Case 1
                     FlagAddress = "C"
                     CellAddress = "D"
                  Case 2
                     FlagAddress = "E"
                     CellAddress = "F"
                  Case 3
                     FlagAddress = "G"
                     CellAddress = "H"
                  Case 4
                     FlagAddress = "I"
                     CellAddress = "J"
                  Case 5
                     FlagAddress = "K"
                     CellAddress = "L"
               End Select
               If AnyBaseRate(Fish, TStep) = 0 Then
                  NewAddress = FlagAddress & CStr(Fish * 4)
                  xlWorkSheet.Range(NewAddress).Value = "*"
                  xlWorkSheet.Range(NewAddress).Interior.Color = RGB(148, 150, 232)
                  NewAddress = CellAddress & CStr(Fish * 4)
                  xlWorkSheet.Range(NewAddress).Interior.Color = RGB(148, 150, 232)
                  NewAddress = CellAddress & CStr(Fish * 4 + 1)
                  xlWorkSheet.Range(NewAddress).Interior.Color = RGB(148, 150, 232)
                  NewAddress = CellAddress & CStr(Fish * 4 + 2)
                  xlWorkSheet.Range(NewAddress).Interior.Color = RGB(148, 150, 232)
                  NewAddress = CellAddress & CStr(Fish * 4 + 3)
                  xlWorkSheet.Range(NewAddress).Interior.Color = RGB(148, 150, 232)
                  GoTo NextNRVal
               End If

               NewAddress = FlagAddress & CStr(Fish * 4)
               xlWorkSheet.Range(NewAddress).Value = NonRetentionFlag(Fish, TStep).ToString
               NewAddress = FlagAddress & CStr(Fish * 4 + 1)
               xlWorkSheet.Range(NewAddress).Value = "*"
               xlWorkSheet.Range(NewAddress).Interior.Color = RGB(229, 150, 100)
               NewAddress = FlagAddress & CStr(Fish * 4 + 2)
               xlWorkSheet.Range(NewAddress).Value = "*"
               xlWorkSheet.Range(NewAddress).Interior.Color = RGB(229, 150, 100)
               NewAddress = FlagAddress & CStr(Fish * 4 + 3)
               xlWorkSheet.Range(NewAddress).Value = "*"
               xlWorkSheet.Range(NewAddress).Interior.Color = RGB(229, 150, 100)

               NewAddress = CellAddress & CStr(Fish * 4)
               xlWorkSheet.Range(NewAddress).Value = NonRetentionInput(Fish, TStep, 1).ToString("#####0")
               NewAddress = CellAddress & CStr(Fish * 4 + 1)
               xlWorkSheet.Range(NewAddress).Value = NonRetentionInput(Fish, TStep, 2).ToString("#####0")
               NewAddress = CellAddress & CStr(Fish * 4 + 2)
               If NonRetentionInput(Fish, TStep, 3) > 0 Then
                  xlWorkSheet.Range(NewAddress).Value = NonRetentionInput(Fish, TStep, 3).ToString("###0.0000")
               Else
                  xlWorkSheet.Range(NewAddress).Value = "0"
               End If
               NewAddress = CellAddress & CStr(Fish * 4 + 3)
               If NonRetentionInput(Fish, TStep, 4) > 0 Then
                  xlWorkSheet.Range(NewAddress).Value = NonRetentionInput(Fish, TStep, 4).ToString("###0.0000")
               Else
                  xlWorkSheet.Range(NewAddress).Value = "0"
               End If
NextNRVal:
            Next
         Next
      ElseIf SpeciesName = "COHO" Then
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               CellAddress = ""
               Select Case TStep
                  Case 1
                     CellAddress = "C" & CStr(Fish + 3)
                  Case 2
                     CellAddress = "D" & CStr(Fish + 3)
                  Case 3
                     CellAddress = "E" & CStr(Fish + 3)
                  Case 4
                     CellAddress = "F" & CStr(Fish + 3)
                  Case 5
                     CellAddress = "G" & CStr(Fish + 3)
               End Select
               If AnyBaseRate(Fish, TStep) = 0 Then
                  xlWorkSheet.Range(CellAddress).Value = "*"
                  xlWorkSheet.Range(CellAddress).Interior.Color = RGB(148, 150, 232)
               Else
                        xlWorkSheet.Range(CellAddress).Value = NonRetentionInput(Fish, TStep, 1).ToString("#####0.0000")
                        'xlWorkSheet.Range(CellAddress).Interior.Color = RGB(255, 255, 255)
               End If
            Next
         Next
      End If

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

      Exit Sub


   End Sub
End Class