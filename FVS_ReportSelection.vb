Imports System.Data.OleDb
Public Class FVS_ReportSelection
   Public ReportTypes() As String
   Public ReportOpts(,) As String
   Public RepDriverName As String
   Public FRAMReps(), RepNums() As Integer

   Private Sub FVS_ReportSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 911
      FormHeight = 931
      FormWidth = 964
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
         If FVS_ReportSelection_ReSize = False Then
            Resize_Form(Me)
            FVS_ReportSelection_ReSize = True
         End If
      End If

      ReDim ReportOpts(250, 6)
      ReDim ReportTypes(14)
      ReDim FRAMReps(18), RepNums(18)
      ReportTypes(1) = "Fishery Mortality Summary"
      ReportTypes(2) = "Stock Mortality Summary"
      ReportTypes(3) = "Terminal Run Summary"
      ReportTypes(4) = "Stock Summary"
      ReportTypes(5) = "Fishery Scalers"
      ReportTypes(6) = "Population Statistics"
      ReportTypes(7) = "Selective Fishery Impacts"
      ReportTypes(8) = "PSC Coho Stock ER"
      ReportTypes(9) = "Stock Exploitation Rate"
      ReportTypes(10) = "Stock/Fishery Impacts Per 1000"
      ReportTypes(11) = "Fishery Stock Composition"
      ReportTypes(12) = "Snake River Fall Chinook Index"
      ReportTypes(13) = "Exploitation Rate Distribution"
      ReportTypes(14) = "Mortality by Age/Time-Step"
      ReportCheckedListBox.Items.Clear()
      For Stk As Integer = 1 To 14
         ReportCheckedListBox.Items.Add(ReportTypes(Stk))
      Next
      NumDriverReports = 0
      ReportSelectedListBox.Items.Clear()
      RepDriverNameTextBox.Text = ""
      DrvListLabel.Text = "Select Report and Requested Options" & vbCrLf & "from Checked List Box Above" _
      & vbCrLf & "Selections will be Displayed in" & vbCrLf & "List Box to the Left"

      '- Set Return Point for Stock, Fishery, Mortality SubRoutines 1=Driver Reports
      CallingRoutine = 1

      '- Report Numbers (from Old FRAM Program)
      RepNums(1) = 1
      RepNums(2) = 3
      RepNums(3) = 2
      RepNums(4) = 7
      RepNums(5) = 6
      RepNums(6) = 8
      RepNums(7) = 14
      RepNums(8) = 15
      RepNums(9) = 16
      RepNums(10) = 18
      RepNums(11) = 17
      RepNums(12) = 4
      RepNums(13) = 9
      RepNums(14) = 5

      '- FRAM Report Numbers
      FRAMReps(1) = 1   ' 1 - Fishery Mortality
      FRAMReps(2) = 3   ' 2 - Terminal Run
      FRAMReps(3) = 2   ' 3 - Stock Mortality
      FRAMReps(4) = 12  ' 4 - Snake River Fall Index
      FRAMReps(5) = 14  ' 5 - Mortality by Stock/Age
      FRAMReps(6) = 5   ' 6 - Fishery Scalers
      FRAMReps(7) = 4   ' 7 - Stock Summary
      FRAMReps(8) = 6   ' 8 - Population Statistics
      FRAMReps(9) = 13  ' 9 - ER Distribution by Fishery Group
      FRAMReps(10) = 0  ' 10 - 
      FRAMReps(11) = 0  ' 11 - 
      FRAMReps(12) = 0  ' 12 - 
      FRAMReps(13) = 0  ' 13 - 
      FRAMReps(14) = 7  ' 14 - Selective Fishery Impacts
      FRAMReps(15) = 8  ' 15 - PSC Coho ER
      FRAMReps(16) = 9  ' 16 - Stock Exploitation Rates
      FRAMReps(17) = 11 ' 17 - Fishery Stock Composition
      FRAMReps(18) = 10 ' 18 - Stock Impacts Per 1000
   End Sub

   Private Sub DrvDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrvDoneButton.Click

      '- SAVE Report Driver to DataBase Table

      If RepDriverNameTextBox.Text = "" Then
         MsgBox("Please Enter Report Driver Name before Saving!!!", MsgBoxStyle.OkOnly)
         Exit Sub
      ElseIf RepDriverNameTextBox.Text.Contains(" ") Then
         MsgBox("Please Enter Report Driver Name WITHOUT Spaces!!!", MsgBoxStyle.OkOnly)
         Exit Sub
      ElseIf RepDriverNameTextBox.Text.Contains(Chr(34)) Then
         MsgBox("Please Enter Report Driver Name WITHOUT Quotation Marks!!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      RepDriverName = RepDriverNameTextBox.Text

      '- Setup Data Adapter Commands
      Dim CmdStr As String
      CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & RepDriverName & Chr(34)
      Dim DrvDA As New OleDb.OleDbCommand(CmdStr, FramDB)
      Dim DriverDA As New System.Data.OleDb.OleDbDataAdapter
      DriverDA.SelectCommand = DrvDA

      CmdStr = "DELETE * FROM ReportDriver WHERE DriverName = " & Chr(34) & RepDriverName & Chr(34)
      Dim DrvDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
      DriverDA.DeleteCommand = DrvDcm

      Dim DRVcb As New OleDb.OleDbCommandBuilder
      DRVcb = New OleDb.OleDbCommandBuilder(DriverDA)

      Dim DrvTrans As OleDb.OleDbTransaction
      Dim DrvCnn As New OleDbCommand

      '- Check if Report Driver Name is already in use
      If FramDataSet.Tables.Contains("ReportDriver") Then
         FramDataSet.Tables("ReportDriver").Rows.Clear()
      End If
      DriverDA.Fill(FramDataSet, "ReportDriver")
      Dim NumRD, Rep As Integer
      NumRD = FramDataSet.Tables("ReportDriver").Rows.Count
      If NumRD <> 0 Then
         MsgBox("Report Driver Name already in use" & vbCrLf & "Please Choose another name!", MsgBoxStyle.OkOnly)
         DriverDA = Nothing
         Exit Sub
      End If

      FramDB.Open()
      'DriverDA.DeleteCommand.ExecuteScalar()
      DrvTrans = FramDB.BeginTransaction
      DrvCnn.Connection = FramDB
      DrvCnn.Transaction = DrvTrans

      For Rep = 1 To NumDriverReports
         DrvCnn.CommandText = "INSERT INTO ReportDriver (DriverName,SpeciesName,ReportNumber,Option1,Option2,Option3,Option4,Option5,Option6) " & _
            "VALUES(" & Chr(34) & RepDriverName.ToString & Chr(34) & "," & _
            Chr(34) & SpeciesName & Chr(34) & "," & _
            RepNums(CInt(ReportOpts(Rep, 0))).ToString & "," & _
            Chr(34) & ReportOpts(Rep, 1) & Chr(34) & "," & _
            Chr(34) & ReportOpts(Rep, 2) & Chr(34) & "," & _
            Chr(34) & ReportOpts(Rep, 3) & Chr(34) & "," & _
            Chr(34) & ReportOpts(Rep, 4) & Chr(34) & "," & _
            Chr(34) & ReportOpts(Rep, 5) & Chr(34) & "," & _
            Chr(34) & ReportOpts(Rep, 6) & Chr(34) & ")"
         DrvCnn.ExecuteNonQuery()
      Next
      DrvTrans.Commit()
      FramDB.Close()

      DriverDA = Nothing

      Me.Close()
      FVS_OutputDriver.Visible = True
   End Sub

   Private Sub DrvCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DrvCancelButton.Click
      Me.Close()
      FVS_OutputDriver.Visible = True
   End Sub

   Private Sub ReportCheckedListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReportCheckedListBox.Click
      '- Add a Report to Driver List
      Dim OptionLine As String
      Dim Grp As Integer
      Dim Result
      ReportNumber = ReportCheckedListBox.SelectedIndex + 1
      NumDriverReports += 1
      ReportOpts(NumDriverReports, 0) = CStr(ReportNumber)
      Select Case ReportNumber

         Case 1 '- Fishery Mortality
            Me.Visible = False
            FVS_MortalityTypeSelection.ShowDialog()
            Me.BringToFront()
            ReportOpts(NumDriverReports, 1) = CStr(MortalityType)
            ReportOpts(NumDriverReports, 2) = ""
            ReportOpts(NumDriverReports, 3) = ""
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case 2 '- Stock Mortality
NextStockReport:
            Me.Visible = False
            FVS_MortalityTypeSelection.ShowDialog()
            Me.BringToFront()
            If MortalityType = 0 Then Exit Sub
            ReportOpts(NumDriverReports, 1) = CStr(MortalityType)
            StockSelectionType = 2  '- Multi Stock Group
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            Me.BringToFront()
            '- Put Stock Selections into Report Option
            OptionLine = ""
            For Stk As Integer = 1 To NumSelectedStocks
               OptionLine &= String.Format("{0,1}", StockSelection(Stk).ToString)
               If Stk <> NumSelectedStocks Then
                  OptionLine &= ","
               End If
            Next
            ReportOpts(NumDriverReports, 2) = CStr(NumSelectedStocks)
            ReportOpts(NumDriverReports, 3) = OptionLine
            ReportOpts(NumDriverReports, 4) = StockGroupName
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""
            Result = MsgBox("Do You Want to do Another Stock Catch Report" & vbCrLf & "for this Report Driver ???", MsgBoxStyle.YesNo)
            If Result = vbYes Then
               ReportNumber = ReportCheckedListBox.SelectedIndex + 1
               NumDriverReports += 1
               ReportOpts(NumDriverReports, 0) = CStr(ReportNumber)
               ReportSelectedListBox.Items.Add(ReportCheckedListBox.SelectedItem & "-" & StockGroupName)
               GoTo NextStockReport
            End If

         Case 3 '- Terminal Run
            If SpeciesName = "CHINOOK" Then
               Result = MsgBox("Do You Want a Brood-Year-AEQ Style Terminal Run Report?", MsgBoxStyle.YesNo)
               If Result = vbYes Then
                  TermRunBYAEQ = True
               Else
                  TermRunBYAEQ = False
               End If
            End If
            'TermRunBYAEQ = True
NextTermRunRep:
            StockSelectionType = 2  '- Multi Stock Selection
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            Me.BringToFront()
            '- Put Stock Selections into Report Option
            OptionLine = ""
            For Stk As Integer = 1 To NumSelectedStocks
               OptionLine &= String.Format("{0,2}", StockSelection(Stk).ToString)
               If Stk <> NumSelectedStocks Then
                  OptionLine &= ","
               End If
            Next
            ReportOpts(NumDriverReports, 1) = OptionLine
            ReportOpts(NumDriverReports, 5) = StockGroupName
            FisherySelectionType = 2 '- Fishery Group w/o GrpName
            Me.Visible = False
            FVS_FisherySelect.ShowDialog()
            Me.BringToFront()
            OptionLine = ""
            For Fish As Integer = 1 To NumSelectedFisheries
               OptionLine &= FisherySelection(Fish).ToString
               If Fish <> NumSelectedFisheries Then
                  OptionLine &= ","
               End If
            Next
            ReportOpts(NumDriverReports, 2) = OptionLine
            Me.Visible = False
            FVS_TRunTimeSteps.ShowDialog()
            OptionLine = TimeStepSelection1.ToString & "," & TimeStepSelection2.ToString
            ReportOpts(NumDriverReports, 3) = OptionLine
            If SpeciesName = "CHINOOK" And TermRunBYAEQ = True Then
               If TermRunTypeSelection = 0 Then
                  ReportOpts(NumDriverReports, 4) = "ETRS Brood Year"
               Else
                  ReportOpts(NumDriverReports, 4) = "TAA Brood Year"
               End If
            Else
               If TermRunTypeSelection = 0 Then
                  ReportOpts(NumDriverReports, 4) = "ETRS"
               Else
                  ReportOpts(NumDriverReports, 4) = "TAA"
               End If
            End If
            If NumDriverReports < 10 Then
               ReportOpts(NumDriverReports, 6) = "00" & NumDriverReports.ToString
            ElseIf NumDriverReports > 9 And NumDriverReports < 100 Then
               ReportOpts(NumDriverReports, 6) = "0" & NumDriverReports.ToString
            Else
               ReportOpts(NumDriverReports, 6) = NumDriverReports.ToString
            End If
            Me.BringToFront()
            Result = MsgBox("Do Another Terminal Run Report for this Selection?", MsgBoxStyle.YesNo)
            If Result = vbYes Then
               NumDriverReports += 1
               ReportOpts(NumDriverReports, 0) = CStr(ReportNumber)
               GoTo NextTermRunRep
            End If

         Case 9 '- Stock ER
            StockSelectionType = 1  '- Single Stock Selection
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            Me.BringToFront()
            ReportOpts(NumDriverReports, 1) = StockSelection(1).ToString
            ReportOpts(NumDriverReports, 2) = ""
            ReportOpts(NumDriverReports, 3) = ""
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case 10 '- Stock Impacts Per 1000
            StockSelectionType = 1  '- Single Stock Selection
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            Me.BringToFront()
            ReportOpts(NumDriverReports, 1) = StockSelection(1).ToString
            ReportOpts(NumDriverReports, 2) = ""
            ReportOpts(NumDriverReports, 3) = ""
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case 11 '- Fishery Stock Composition
            FisherySelectionType = 2 '- Fishery Group w/o GrpName
            Me.Visible = False
            FVS_FisherySelect.ShowDialog()
            Me.BringToFront()
            OptionLine = ""
            For Fish As Integer = 1 To NumSelectedFisheries
               OptionLine &= FisherySelection(Fish).ToString
               If Fish <> NumSelectedFisheries Then
                  OptionLine &= ","
               End If
            Next
            ReportOpts(NumDriverReports, 1) = CStr(OptionLine)
            ReportOpts(NumDriverReports, 2) = ""
            ReportOpts(NumDriverReports, 3) = ""
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case 13 '- ER Distribution
            Me.Visible = False
            FVS_MortalityTypeSelection.ShowDialog()
            Me.BringToFront()
            '- Put Mortality Type into Option1
            ReportOpts(NumDriverReports, 1) = CStr(MortalityType)
            StockSelectionType = 2  '- Multi Stock Selection
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            Me.BringToFront()
            '- Put Stock Selections into Option2
            OptionLine = ""
            For Stk As Integer = 1 To NumSelectedStocks
               OptionLine &= String.Format("{0,2}", StockSelection(Stk).ToString)
               If Stk <> NumSelectedStocks Then
                  OptionLine &= ","
               End If
            Next
            ReportOpts(NumDriverReports, 2) = CStr(OptionLine)
            Me.Visible = False
            FVS_FisheryMultiGroupSelect.ShowDialog()
            Me.BringToFront()
            '- Put Multi-Fishery Group Selections into Option3
            OptionLine = ""
            '- Number of Fishery Groups is First Variable
            OptionLine = NumFisheryGroups.ToString & ","
            For Grp = 1 To NumFisheryGroups
               '- Number Fisheries in this Group
               OptionLine &= SelectFisheryList(Grp, 0).ToString & ","
               For Fish As Integer = 1 To SelectFisheryList(Grp, 0)
                  OptionLine &= SelectFisheryList(Grp, Fish).ToString
                  If (Fish = SelectFisheryList(Grp, 0)) And (Grp = NumFisheryGroups) Then
                     '- Last Fishery in Last Group
                     Exit For
                  Else
                     OptionLine &= ","
                  End If
               Next
            Next
            ReportOpts(NumDriverReports, 4) = CStr(OptionLine)
            '- Put Fishery Group Names into Option4
            OptionLine = ""
            For Grp = 1 To NumFisheryGroups
               OptionLine &= FisheryGroupNames(Grp)
               If Grp <> NumFisheryGroups Then OptionLine &= ","
            Next
            ReportOpts(NumDriverReports, 4) = CStr(OptionLine)
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case 14 '- Mortality by Age/Time-Step
            Me.Visible = False
            FVS_MortalityTypeSelection.ShowDialog()
            ReportOpts(NumDriverReports, 1) = CStr(MortalityType)
            StockSelectionType = 2  '- Multi Stock Selection
            Me.Visible = False
            FVS_StockSelect.ShowDialog()
            '- Put Stock Selections into Report Option
            OptionLine = ""
            For Stk As Integer = 1 To NumStk
               '- Fill Line with Zeros
               OptionLine &= "0"
            Next
            For Stk As Integer = 1 To NumSelectedStocks
               '- Put Ones in Selected Stock Positions
               Mid(OptionLine, Stk, 1) = "1"
            Next
            ReportOpts(NumDriverReports, 3) = CStr(OptionLine)
            Me.BringToFront()
            Me.Visible = False
            FisherySelectionType = 2 '- Fishery Group w/o GrpName
            FVS_FisherySelect.ShowDialog()
            For Fish As Integer = 1 To NumFish
               '- Fill Line with Zeros
               OptionLine &= "0"
            Next
            For Fish As Integer = 1 To NumSelectedFisheries
               '- Put Ones in Selected Fishery Positions
               Mid(OptionLine, Fish, 1) = "1"
            Next
            ReportOpts(NumDriverReports, 2) = CStr(OptionLine)
            Me.BringToFront()
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

         Case Else
            '- All Other Reports without Parameters to pass
            '- Stock-Summary, Pop Stats, PSC-Coho_ER, SnkRvrFallIndx
            ReportOpts(NumDriverReports, 1) = ""
            ReportOpts(NumDriverReports, 2) = ""
            ReportOpts(NumDriverReports, 3) = ""
            ReportOpts(NumDriverReports, 4) = ""
            ReportOpts(NumDriverReports, 5) = ""
            ReportOpts(NumDriverReports, 6) = ""

      End Select

      '- Put Report Selection in ReportCheckedListBox
      If ReportNumber = 2 Then
         ReportSelectedListBox.Items.Add(ReportCheckedListBox.SelectedItem & "-" & StockGroupName)
      Else
         ReportSelectedListBox.Items.Add(ReportCheckedListBox.SelectedItem)
      End If

   End Sub

   Private Sub ReportSelectedListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReportSelectedListBox.Click
      '- Remove a Report from Driver Selected List
      Dim Result, RepDelete, RepNum As Integer
      RepDelete = ReportSelectedListBox.SelectedIndex + 1
      Result = MsgBox("Remove this Report from Driver List ???" & vbCrLf & "Report=" & ReportTypes(ReportOpts(RepDelete, 0)), MsgBoxStyle.YesNo)
      If Result = vbYes Then
         '- Remove selection from RepOpts storage array
         For RepNum = RepDelete To NumDriverReports - 1
            For Fish As Integer = 0 To 6
               ReportOpts(RepNum, Fish) = ReportOpts(RepNum + 1, Fish)
            Next
         Next
         '- Set last report in array to nothing
         For Fish As Integer = 0 To 6
            ReportOpts(NumDriverReports, Fish) = Nothing
         Next
         NumDriverReports -= 1
      Else
         Exit Sub
      End If
      '- Update Driver Selected List
      ReportSelectedListBox.Items.Clear()
      For RepNum = 1 To NumDriverReports
         ReportSelectedListBox.Items.Add(ReportTypes(CInt(ReportOpts(RepNum, 0))))
      Next
      ReportSelectedListBox.Update()
   End Sub

    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class