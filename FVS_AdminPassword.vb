Imports System.Data.OleDb
Imports Microsoft.Office.Interop

Public Class FVS_AdminPassword

   Public FailedLoad As Boolean
   Public PasswordCheck As String

   Private Sub FVS_AdminPassword_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        'txt_pwentry.Clear()
      'txt_pwentry.Focus()
   End Sub

   Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
      Me.Close()
      FVS_RunModel.Visible() = True
   End Sub


   Private Sub Button1_Click_1(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        'PasswordCheck = "password"
      FailedLoad = False

        'If txt_pwentry.Text = "" Then
        '   MessageBox.Show("Enter the secret nuclear codes or click cancel to exit admin panel")
        '   txt_pwentry.Clear()
        'ElseIf txt_pwentry.Text = PasswordCheck Then
        'Dim result = MessageBox.Show("Are you ABSOLUTELY sure you want to do an SLRatio update?" & _
        '                                Environment.NewLine & "(If yes, click 'Yes' and complete model run as per usual.)", "SLRatio", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        'If result = DialogResult.No Then
        '    UpdateRunEncounterRateAdjustment = False
        '    Me.Close()
        '    FVS_RunModel.Visible() = True
        'ElseIf result = DialogResult.Yes Then

        'Call LoadSLRatio Subroutine if loading from spreadsheet is desired
        If chk_LoadSLRatio.Checked = True Then
            Call LoadSLRatio()
            If FailedLoad = True Then
                chk_LoadSLRatio.Checked = False
                Exit Sub
            End If
        End If

        UpdateRunEncounterRateAdjustment = True
        WhoUpdated = Environment.UserName
        Me.Close()
        FVS_RunModel.Visible() = True
        '**************************************
        'Call the RunEncounterRateAdjustment process here or back on run screen?
        '**************************************
        '  End If
        'Else
        'MessageBox.Show("Incorrect password. Re-enter or click cancel to exit admin panel")
        'UpdateRunEncounterRateAdjustment = False
        'txt_pwentry.Clear()
        'End If
   End Sub

   Sub LoadSLRatio()

      'If loading from spreadsheet is desired, dump anything that's associated with a run and replace with the new values
      'Values will be cleaned/replaced within the database when SaveDat() runs within FRAMCalcs.

      Dim xlApp1 As Excel.Application
      Dim ExcelWasNotRunning As Boolean
      Dim WorkBookWasNotOpen As Boolean
      Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
      Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
      Dim n As Integer
      Dim OpenRatExcel As New OpenFileDialog()
      Dim TargRat, TargRatPath As String
      'Dim TempTargetRatio(NumFish, MaxAge, NumSteps)

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp1 = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp1 = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenRatExcel.Filter = "All Excel Files (*.xls; *.xlsx; *xlsm)|*.xls; *.xlsx; *xlsm|All files (*.*)|*.*"
      OpenRatExcel.FilterIndex = 1
      OpenRatExcel.RestoreDirectory = True

      If OpenRatExcel.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         TargRat = OpenRatExcel.FileName
         TargRatPath = My.Computer.FileSystem.GetFileInfo(TargRat).DirectoryName
      Else
         Exit Sub
         FailedLoad = True 'Don't load anything 
      End If

      ''- Test if Excel was Running
      'ExcelWasNotRunning = True
      'Try
      '   xlApp1 = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
      '   ExcelWasNotRunning = False
      'Catch ex As Exception
      '   xlApp1 = New Microsoft.Office.Interop.Excel.Application()
      'End Try

      '- Test if TargetRatio spreadsheet is Open
      WorkBookWasNotOpen = True
      Dim wbName As String
      wbName = My.Computer.FileSystem.GetFileInfo(TargRat).Name
      For Each xlWorkBook In xlApp1.Workbooks
         If xlWorkBook.Name = wbName Then
            WorkBookWasNotOpen = False
            GoTo SkipWBOpen
         End If
      Next
      xlWorkBook = xlApp1.Workbooks.Open(TargRat)
      xlApp1.WindowState = Excel.XlWindowState.xlMinimized

SkipWBOpen:

      xlApp1.Application.DisplayAlerts = False
      xlApp1.Visible = False
      xlApp1.WindowState = Excel.XlWindowState.xlMinimized

      '- Find the import spreadsheet
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 1 Then
            If xlWorkSheet.Name = "SLRatioImport" Then Exit For
         End If
      Next

      '- Check if spreadsheet doesn't contain targetratio import worksheet...
      If xlWorkSheet.Name <> "SLRatioImport" Then
         MsgBox("Can't find the 'SLRatioImport' worksheet in the file you selected." & vbCrLf & _
                 "The SLRatio update process has been terminated!" & vbCrLf & _
                "Reinitialize with an approriate file to continue the update process.", MsgBoxStyle.OkOnly)
         FailedLoad = True
         GoTo CloseExcelWorkBook
      End If


      Dim xlen As Boolean
      Dim i As Integer
      Dim Addy As String
      Dim AgeSpecific As Boolean

      AgeSpecific = False 'Start with the assumption that values are provided for all ages and timesteps

      'First, determine the number of observations within the dataset
      i = 5
      Do While xlen = False
         Addy = "A" & i
         xlen = IsNothing(xlWorkSheet.Range(Addy).Value)
         i = i + 1
      Loop
      n = i - 6

      'Second, determine whether or not age-specific values are provided
      If xlWorkSheet.Range("D2").Value = "Yes" Then
         AgeSpecific = True
      End If

      'Third, read in the values and replace whatever (if anything) is in the array before running.
      For i = 5 To n + 4
         If AgeSpecific = False Then

            If xlWorkSheet.Range("D" & i).Value > 0 Then
               'Allows for flagging -1 for no data.
               TargetRatio(xlWorkSheet.Range("A" & i).Value, xlWorkSheet.Range("B" & i).Value, xlWorkSheet.Range("C" & i).Value) = xlWorkSheet.Range("D" & i).Value
            End If
         Else
            If xlWorkSheet.Range("D" & i).Value > 0 Then
               TargetRatio(xlWorkSheet.Range("A" & i).Value, 2, xlWorkSheet.Range("C" & i).Value) = xlWorkSheet.Range("D" & i).Value
               TargetRatio(xlWorkSheet.Range("A" & i).Value, 3, xlWorkSheet.Range("C" & i).Value) = xlWorkSheet.Range("D" & i).Value
               TargetRatio(xlWorkSheet.Range("A" & i).Value, 4, xlWorkSheet.Range("C" & i).Value) = xlWorkSheet.Range("D" & i).Value
               TargetRatio(xlWorkSheet.Range("A" & i).Value, 5, xlWorkSheet.Range("C" & i).Value) = xlWorkSheet.Range("D" & i).Value
            End If
         End If
      Next


CloseExcelWorkBook:
      '- Done with loading SLRatio Data..Close and release object
      xlApp1.Application.DisplayAlerts = False
      'xlWorkBook.Save() 'Don't need to save this...
      If WorkBookWasNotOpen = True Then
         xlWorkBook.Close()
      End If
      If ExcelWasNotRunning = True Then
         xlApp1.Application.Quit()
         xlApp1.Quit()
      Else
         xlApp1.Visible = True
         xlApp1.WindowState = Excel.XlWindowState.xlMinimized
      End If
      xlApp1.Application.DisplayAlerts = True
      xlApp1 = Nothing
      Me.Cursor = Cursors.Default

   End Sub

End Class