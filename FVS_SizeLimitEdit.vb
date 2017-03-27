''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
'Imports System.Data.OleDb
'Imports Microsoft.Office.Interop
''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

Public Class FVS_SizeLimitEdit

   Private Sub FVS_SizeLimitEdit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      '#################################################################################
      'Pete-Feb 2013 Code to Make Invisible the size limit loading button & checkbox
      btnLimitChange.Visible = False
      btnLimitChange.Enabled = False
      SizeLimitBox.Visible = False
      SizeLimitBox.Enabled = False
      'Pete-Feb 2013 Code to Make Invisible the size limit loading button & checkbox
      '#################################################################################


      'FormHeight = 865
      FormHeight = 885
      FormWidth = 920
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
         If FVS_SizeLimitEdit_ReSize = False Then
            Resize_Form(Me)
            FVS_SizeLimitEdit_ReSize = True
         End If
      End If

      '- Fill the DataGrid with Values ... COHO and CHINOOK are different
      SizeLimitGrid.Columns.Clear()
      SizeLimitGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      SizeLimitGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      SizeLimitGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      If SpeciesName = "COHO" Then
         '- Save this for Future Coho Size Limits
         If SizeLimitGrid.ColumnCount = 0 Then
            SizeLimitGrid.Columns.Add("FisheryName", "Name")
            SizeLimitGrid.Columns("FisheryName").Width = 100 / FormWidthScaler
            SizeLimitGrid.Columns("FisheryName").ReadOnly = True
            SizeLimitGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine
            SizeLimitGrid.Columns.Add("FishNum", "#")
            SizeLimitGrid.Columns("FishNum").Width = 40 / FormWidthScaler
            SizeLimitGrid.Columns("FishNum").ReadOnly = True
            SizeLimitGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

            SizeLimitGrid.Columns.Add("Time1Estimate", "T1-MinSize")
            SizeLimitGrid.Columns("Time1Estimate").Width = 85 / FormWidthScaler
            SizeLimitGrid.Columns("Time1Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time2Estimate", "T2-MinSize")
            SizeLimitGrid.Columns("Time2Estimate").Width = 85 / FormWidthScaler
            SizeLimitGrid.Columns("Time2Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time3Estimate", "T3-MinSize")
            SizeLimitGrid.Columns("Time3Estimate").Width = 85 / FormWidthScaler
            SizeLimitGrid.Columns("Time3Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time4Estimate", "T4-MinSize")
            SizeLimitGrid.Columns("Time4Estimate").Width = 85 / FormWidthScaler
            SizeLimitGrid.Columns("Time4Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time5Estimate", "T5-MinSize")
            SizeLimitGrid.Columns("Time5Estimate").Width = 85 / FormWidthScaler
            SizeLimitGrid.Columns("Time5Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.RowCount = NumFish

         End If

      ElseIf SpeciesName = "CHINOOK" Then

         If SizeLimitGrid.ColumnCount = 0 Then
            SizeLimitGrid.Columns.Add("FisheryName", "Name")
            SizeLimitGrid.Columns("FisheryName").Width = 200 / FormWidthScaler
            SizeLimitGrid.Columns("FisheryName").ReadOnly = True
            SizeLimitGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine
            SizeLimitGrid.Columns.Add("FishNum", "#")
            SizeLimitGrid.Columns("FishNum").Width = 40 / FormWidthScaler
            SizeLimitGrid.Columns("FishNum").ReadOnly = True
            SizeLimitGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

            SizeLimitGrid.Columns.Add("Time1Estimate", "Oct-Apr-1")
            SizeLimitGrid.Columns("Time1Estimate").Width = 100 / FormWidthScaler
            SizeLimitGrid.Columns("Time1Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time2Estimate", "May-June")
            SizeLimitGrid.Columns("Time2Estimate").Width = 100 / FormWidthScaler
            SizeLimitGrid.Columns("Time2Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time3Estimate", "July-Sept")
            SizeLimitGrid.Columns("Time3Estimate").Width = 100 / FormWidthScaler
            SizeLimitGrid.Columns("Time3Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.Columns.Add("Time4Estimate", "Oct-Apr-2")
            SizeLimitGrid.Columns("Time4Estimate").Width = 100 / FormWidthScaler
            SizeLimitGrid.Columns("Time4Estimate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

            SizeLimitGrid.RowCount = NumFish

         End If

      End If

      '- Load MinSizeLimit Array into SizeLimitGrid
      For Fish As Integer = 1 To NumFish
         SizeLimitGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
         SizeLimitGrid.Item(1, Fish - 1).Value = Fish.ToString
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 1 Then
               SizeLimitGrid.Item(TStep + 1, Fish - 1).Value = MinSizeLimit(Fish, TStep)
            Else
               SizeLimitGrid.Item(TStep + 1, Fish - 1).Value = "****"
               SizeLimitGrid.Item(TStep + 1, Fish - 1).Style.BackColor = Color.LightBlue
            End If
         Next
      Next

   End Sub

   Private Sub SLCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SLCancelButton.Click
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub SLDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SLDoneButton.Click
      '- Save SizeLimitGrid into MinSizeLimit Array 
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 1 Then
               If SizeLimitGrid.Item(TStep + 1, Fish - 1).Value <> MinSizeLimit(Fish, TStep) Then
                  MinSizeLimit(Fish, TStep) = CInt(SizeLimitGrid.Item(TStep + 1, Fish - 1).Value)
                  ChangeSizeLimit = True
               End If
            End If
         Next
      Next
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      ClipStr = "Name" & vbTab & "#" & vbTab & "Oct-Apr-1" & vbTab & "May-June" & vbTab & "July-Sept" & vbTab & "Oct-Apr-2" & vbCr
      For RecNum = 0 To SizeLimitGrid.RowCount - 1
         For ColNum = 0 To 5
            If ColNum = 0 Then
               ClipStr = ClipStr & SizeLimitGrid.Item(ColNum, RecNum).Value
            Else
               ClipStr = ClipStr & vbTab & SizeLimitGrid.Item(ColNum, RecNum).Value
            End If
         Next
         ClipStr = ClipStr & vbCr
      Next
      Clipboard.SetDataObject(ClipStr)
   End Sub

   Private Sub SizeLimitGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles SizeLimitGrid.CellContentClick

   End Sub

   '   '#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
   'Private Sub btnLimitChange_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLimitChange.Click


   '      Dim OpenSizeLimitSpreadsheet As New OpenFileDialog()
   '      Dim LimitSpreadsheet, LimitSpreadsheetPath As String
   '      ReDim AltFlag(NumFish, NumSteps)
   '      ReDim AltLimitNS(NumFish, NumSteps), AltLimitMSF(NumFish, NumSteps)
   '      ReDim ShakerFlagNS(NumFish, NumSteps), ShakerFlagMSF(NumFish, NumSteps)
   '      ReDim LSRatioNS(NumFish, NumSteps), LSRatioMSF(NumFish, NumSteps)
   '      ReDim ExtShakerNS(NumFish, NumSteps), ExtShakerMSF(NumFish, NumSteps)
   '      ReDim ExternalBaseRatio(NumFish, NumSteps)

   '      '- Test if Excel was Running
   '      ExcelWasNotRunning = True
   '      Try
   '         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
   '         ExcelWasNotRunning = False
   '      Catch ex As Exception
   '         xlApp = New Microsoft.Office.Interop.Excel.Application()
   '      End Try

   '      OpenSizeLimitSpreadsheet.Filter = "Size Limit Change Templates (*.xls)|*.xls|All files (*.*)|*.*"
   '      OpenSizeLimitSpreadsheet.FilterIndex = 1
   '      OpenSizeLimitSpreadsheet.RestoreDirectory = True

   '      If OpenSizeLimitSpreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
   '         LimitSpreadsheet = OpenSizeLimitSpreadsheet.FileName
   '         LimitSpreadsheetPath = My.Computer.FileSystem.GetFileInfo(LimitSpreadsheet).DirectoryName
   '      Else
   '         Exit Sub
   '      End If

   '      '- Test if Excel was Running
   '      ExcelWasNotRunning = True
   '      Try
   '         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
   '         ExcelWasNotRunning = False
   '      Catch ex As Exception
   '         xlApp = New Microsoft.Office.Interop.Excel.Application()
   '      End Try

   '      '- Test if Size Limit Workbook is Open
   '      WorkBookWasNotOpen = True
   '      Dim wbName As String
   '      wbName = My.Computer.FileSystem.GetFileInfo(LimitSpreadsheet).Name
   '      For Each xlWorkBook In xlApp.Workbooks
   '         If xlWorkBook.Name = wbName Then
   '            'xlApp.Workbooks.Close()
   '            'xlWorkBook = xlApp.Workbooks.Open(FRAMCatchSpreadSheet)
   '            'xlWorkBook.Activate()
   '            WorkBookWasNotOpen = False
   '            GoTo SkipWBOpen
   '         End If
   '      Next
   '      xlWorkBook = xlApp.Workbooks.Open(LimitSpreadsheet)
   '      xlApp.WindowState = Excel.XlWindowState.xlMinimized
   'SkipWBOpen:

   '      xlApp.Application.DisplayAlerts = False
   '      xlApp.Visible = False
   '      xlApp.WindowState = Excel.XlWindowState.xlMinimized

   '      '- Find WorkSheets with FRAM Catch numbers
   '      For Each xlWorkSheet In xlWorkBook.Worksheets
   '         If xlWorkSheet.Name.Length > 7 Then
   '            If xlWorkSheet.Name = "SizeLimitTemplate" Then Exit For
   '         End If
   '      Next

   '      '- Check if DataBase contains FRAMInput Worksheet
   '      If xlWorkSheet.Name <> "SizeLimitTemplate" Then
   '         MsgBox("Can't Find 'SizeLimitTemplate' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
   '                "Please Choose appropriate Spreadsheet with Size Limit Inputs!", MsgBoxStyle.OkOnly)
   '         GoTo CloseExcelWorkBook
   '      End If

   '      '- Check first Fishery Name for correct Species Spreadsheet
   '      Dim testname As String
   '      testname = xlWorkSheet.Range("A4").Value
   '      If SpeciesName = "CHINOOK" Then
   '         If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
   '            MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your Spreadsheet Selection" & vbCrLf & _
   '                   "Please Choose appropriate CHINOOK Spreadsheet with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
   '            GoTo CloseExcelWorkBook
   '         End If
   '      ElseIf SpeciesName = "COHO" Then
   '         If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
   '            MsgBox("Can't Find 'No Cal Trm' as first Fishery your Spreadsheet Selection" & vbCrLf & _
   '                   "Please Choose appropriate COHO Spreadsheet with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
   '            GoTo CloseExcelWorkBook
   '         End If
   '      End If

   '      '- Load WorkSheet Catch into Quota Array (Change Flag)
   '      Me.Cursor = Cursors.WaitCursor
   '      Dim CellAddress1, CellAddress2, CellAddress3, CellAddress4 As String
   '      Dim CellAddress5, CellAddress6, CellAddress7, CellAddress8 As String
   '      Dim CellAddress9, CellAddress10, CellAddress11, CellAddress12 As String
   '      Dim FlagAddress As String
   '      For Fish = 1 To NumFish
   '         For TStep = 1 To NumSteps
   '            CellAddress1 = ""
   '            CellAddress2 = ""
   '            FlagAddress = ""
   '            If SpeciesName = "CHINOOK" Then
   '               Select Case TStep
   '                  Case 1
   '                     FlagAddress = "C" & CStr(Fish + 3)
   '                     CellAddress1 = "D" & CStr(Fish + 3)
   '                     CellAddress2 = "E" & CStr(Fish + 3)
   '                     CellAddress3 = "F" & CStr(Fish + 3)
   '                     CellAddress4 = "G" & CStr(Fish + 3)
   '                     CellAddress5 = "H" & CStr(Fish + 3)
   '                     CellAddress6 = "I" & CStr(Fish + 3)
   '                     CellAddress7 = "J" & CStr(Fish + 3)
   '                     CellAddress8 = "K" & CStr(Fish + 3)
   '                     CellAddress9 = "AM" & CStr(Fish + 3)
   '                  Case 2
   '                     FlagAddress = "L" & CStr(Fish + 3)
   '                     CellAddress1 = "M" & CStr(Fish + 3)
   '                     CellAddress2 = "N" & CStr(Fish + 3)
   '                     CellAddress3 = "O" & CStr(Fish + 3)
   '                     CellAddress4 = "P" & CStr(Fish + 3)
   '                     CellAddress5 = "Q" & CStr(Fish + 3)
   '                     CellAddress6 = "R" & CStr(Fish + 3)
   '                     CellAddress7 = "S" & CStr(Fish + 3)
   '                     CellAddress8 = "T" & CStr(Fish + 3)
   '                     CellAddress9 = "AN" & CStr(Fish + 3)
   '                  Case 3
   '                     FlagAddress = "U" & CStr(Fish + 3)
   '                     CellAddress1 = "V" & CStr(Fish + 3)
   '                     CellAddress2 = "W" & CStr(Fish + 3)
   '                     CellAddress3 = "X" & CStr(Fish + 3)
   '                     CellAddress4 = "Y" & CStr(Fish + 3)
   '                     CellAddress5 = "Z" & CStr(Fish + 3)
   '                     CellAddress6 = "AA" & CStr(Fish + 3)
   '                     CellAddress7 = "AB" & CStr(Fish + 3)
   '                     CellAddress8 = "AC" & CStr(Fish + 3)
   '                     CellAddress9 = "AO" & CStr(Fish + 3)
   '                  Case 4
   '                     FlagAddress = "AD" & CStr(Fish + 3)
   '                     CellAddress1 = "AE" & CStr(Fish + 3)
   '                     CellAddress2 = "AF" & CStr(Fish + 3)
   '                     CellAddress3 = "AG" & CStr(Fish + 3)
   '                     CellAddress4 = "AH" & CStr(Fish + 3)
   '                     CellAddress5 = "AI" & CStr(Fish + 3)
   '                     CellAddress6 = "AJ" & CStr(Fish + 3)
   '                     CellAddress7 = "AK" & CStr(Fish + 3)
   '                     CellAddress8 = "AL" & CStr(Fish + 3)
   '                     CellAddress9 = "AP" & CStr(Fish + 3)
   '               End Select
   '            End If

   '            If IsNumeric(xlWorkSheet.Range(FlagAddress).Value) Then
   '               AltFlag(Fish, TStep) = xlWorkSheet.Range(FlagAddress).Value
   '               AltLimitNS(Fish, TStep) = xlWorkSheet.Range(CellAddress3).Value
   '               AltLimitMSF(Fish, TStep) = xlWorkSheet.Range(CellAddress4).Value
   '               ShakerFlagNS(Fish, TStep) = xlWorkSheet.Range(CellAddress1).Value
   '               ShakerFlagMSF(Fish, TStep) = xlWorkSheet.Range(CellAddress2).Value
   '               LSRatioNS(Fish, TStep) = xlWorkSheet.Range(CellAddress5).Value
   '               LSRatioMSF(Fish, TStep) = xlWorkSheet.Range(CellAddress6).Value
   '               ExtShakerNS(Fish, TStep) = xlWorkSheet.Range(CellAddress7).Value
   '               ExtShakerMSF(Fish, TStep) = xlWorkSheet.Range(CellAddress8).Value
   '               ExternalBaseRatio(Fish, TStep) = xlWorkSheet.Range(CellAddress9).Value
   '            End If
   '         Next
   '      Next

   '      For Fish = 1 To NumFish
   '         For TStep = 1 To NumSteps
   '            Debug.Print(FisheryName(Fish) & "TS " & TStep & "," & _
   '            "Alt Flag = " & AltFlag(Fish, TStep) & " ," & _
   '            "Alt Limit NS = " & AltLimitNS(Fish, TStep) & " ," & _
   '            "Alt Limit MSF = " & AltLimitMSF(Fish, TStep) & " ," & _
   '            "Shaker Flag NS = " & ShakerFlagNS(Fish, TStep) & " ," & _
   '            "Shaker Flag MSF = " & ShakerFlagMSF(Fish, TStep) & " ," & _
   '            "LS Ratio NS = " & LSRatioNS(Fish, TStep) & " ," & _
   '            "LS Ratio msf = " & LSRatioMSF(Fish, TStep) & " ," & _
   '            "ExtShak NS = " & ExtShakerNS(Fish, TStep) & " ," & _
   '            "ExtShak MSF = " & ExtShakerMSF(Fish, TStep) & " ," & _
   '            "22 Inch Ratio = " & ExternalBaseRatio(Fish, TStep))
   '         Next TStep
   '      Next Fish

   'CloseExcelWorkBook:
   '      '- Done with FRAM-Template WorkBook .. Close and release object
   '      xlApp.Application.DisplayAlerts = False
   '      'xlWorkBook.Save()
   '      If WorkBookWasNotOpen = True Then
   '         xlWorkBook.Close()
   '      End If
   '      If ExcelWasNotRunning = True Then
   '         xlApp.Application.Quit()
   '         xlApp.Quit()
   '      Else
   '         xlApp.Visible = True
   '         xlApp.WindowState = Excel.XlWindowState.xlMinimized
   '      End If
   '      xlApp.Application.DisplayAlerts = True
   '      xlApp = Nothing

   '      Me.Cursor = Cursors.Default

   '      Exit Sub

   'End Sub
   '   '#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.



   ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.
   ''#Broad scope variable to tell FRAM to use the CNR sublegals as external shakers calc methods
   'Public Sub SizeLimitBox_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles SizeLimitBox.CheckedChanged
   '   'If SizeLimitBox.Checked = True Then
   '   '   SizeLimitScenario = True
   '   'End If
   'End Sub
   ''#################### Size Limit & External Shaker Code ###########################  -- Pete Dec 2012.

End Class