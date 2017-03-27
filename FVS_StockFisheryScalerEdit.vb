Imports System.Data.OleDb
Public Class FVS_StockFisheryScalerEdit

   Private Sub StockFisheryScalerEdit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 884
      FormHeight = 900
      FormWidth = 985
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
         If FVS_StockFisheryScalerEdit_ReSize = False Then
            Resize_Form(Me)
            FVS_StockFisheryScalerEdit_ReSize = True
         End If
      End If

      '- Initialize ComboBox
      SFSComboBox.Items.Clear()
      For Fish As Integer = 1 To NumFish
         SFSComboBox.Items.Add(FisheryTitle(Fish))
      Next
      Fish = 0
      SFSComboBox.SelectedItem = 0

      '- Initialize DataViewGrid
      StockFisheryGrid.Columns.Clear()
      StockFisheryGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      StockFisheryGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      StockFisheryGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", 10, FontStyle.Bold)

      If SpeciesName = "CHINOOK" Then
         StockFisheryGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", 10, FontStyle.Bold)
         StockFisheryGrid.Columns.Add("FisheryName", "Name")
         StockFisheryGrid.Columns("FisheryName").Width = 300 / FormWidthScaler
         StockFisheryGrid.Columns("FisheryName").ReadOnly = True
         StockFisheryGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine

         StockFisheryGrid.Columns.Add("FishNum", "#")
         StockFisheryGrid.Columns("FishNum").Width = 40 / FormWidthScaler
         StockFisheryGrid.Columns("FishNum").ReadOnly = True
         StockFisheryGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine
         StockFisheryGrid.Columns("FishNum").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time1Rate", "Oct-Apr1")
         StockFisheryGrid.Columns("Time1Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time1Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time1Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time2Rate", "May-June")
         StockFisheryGrid.Columns("Time2Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time2Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time2Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time3Rate", "July-Sept")
         StockFisheryGrid.Columns("Time3Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time3Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time3Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time4Rate", "Oct-Apr2")
         StockFisheryGrid.Columns("Time4Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time4Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time4Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.RowCount = NumStk

      ElseIf SpeciesName = "COHO" Then
         StockFisheryGrid.Columns.Add("FisheryName", "Name")
         StockFisheryGrid.Columns("FisheryName").Width = 300 / FormWidthScaler
         StockFisheryGrid.Columns("FisheryName").ReadOnly = True
         StockFisheryGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine

         StockFisheryGrid.Columns.Add("FishNum", "#")
         StockFisheryGrid.Columns("FishNum").Width = 40 / FormWidthScaler
         StockFisheryGrid.Columns("FishNum").ReadOnly = True
         StockFisheryGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine
         StockFisheryGrid.Columns("FishNum").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time1Rate", "Jan-June")
         StockFisheryGrid.Columns("Time1Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time1Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time1Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time2Rate", "July")
         StockFisheryGrid.Columns("Time2Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time2Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time2Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time3Rate", "August")
         StockFisheryGrid.Columns("Time3Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time3Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time3Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time4Rate", "September")
         StockFisheryGrid.Columns("Time4Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time4Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time4Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.Columns.Add("Time5Rate", "Oct-Dec")
         StockFisheryGrid.Columns("Time5Rate").Width = 100 / FormWidthScaler
         StockFisheryGrid.Columns("Time5Rate").DefaultCellStyle.BackColor = Color.White
         StockFisheryGrid.Columns("Time5Rate").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

         StockFisheryGrid.RowCount = NumStk
      End If

      '- Fill Grid with Array Values
      For Stk As Integer = 1 To NumStk
         StockFisheryGrid.Item(0, Stk - 1).Value = StockTitle(Stk)
         StockFisheryGrid.Item(1, Stk - 1).Value = Stk.ToString
         For TStep As Integer = 1 To NumSteps
            StockFisheryGrid.Item(TStep + 1, Stk - 1).Value = StockFishRateScalers(Stk, Fish, TStep)
         Next
      Next

      '- Set Return Point for Stock, Fishery, Mortality SubRoutines 3=StockFisheryEdit
      CallingRoutine = 3

   End Sub

   Private Sub SFCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SFCancelButton.Click
      '- This only affects the Last Fishery Selected
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub SFDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SFDoneButton.Click

      '- Exit if No Fishery Selected
      If Fish = 0 Then
         Me.Close()
         FVS_InputMenu.Visible = True
      End If

      '- Put Changes from Last Fishery Selected into StockFishInput Array
      For Stk As Integer = 1 To NumStk
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 1 Then
               If StockFisheryGrid.Item(TStep + 1, Stk - 1).Value <> StockFishRateScalers(Stk, Fish, TStep) Then
                  StockFishRateScalers(Stk, Fish, TStep) = CDbl(StockFisheryGrid.Item(TStep + 1, Stk - 1).Value)
                  ChangeStockFishScaler = True
               End If
            End If
         Next
      Next

      Me.Close()
      FVS_InputMenu.Visible = True

   End Sub


   Private Sub SFSComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles SFSComboBox.SelectedIndexChanged

      If Fish = 0 Then GoTo SkipSFSSave
      '- Put Changes (if any) from previous Fishery Selection into StockFishInput Array
      For Stk As Integer = 1 To NumStk
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 1 Then
               If StockFisheryGrid.Item(TStep + 1, Stk - 1).Value <> StockFishRateScalers(Stk, Fish, TStep) Then
                  StockFishRateScalers(Stk, Fish, TStep) = CDbl(StockFisheryGrid.Item(TStep + 1, Stk - 1).Value)
                  ChangeStockFishScaler = True
               End If
            End If
         Next
      Next

SkipSFSSave:
      '- Fill DataViewGrid with New Fishery Selection
      Fish = SFSComboBox.SelectedIndex + 1
      For Stk As Integer = 1 To NumStk
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 0 Then
               StockFisheryGrid.Item(TStep + 1, Stk - 1).Value = "****"
               StockFisheryGrid.Item(TStep + 1, Stk - 1).Style.BackColor = Color.LightBlue
            Else
               If StockFishRateScalers(Stk, Fish, TStep) = 1 Then
                  StockFisheryGrid.Item(TStep + 1, Stk - 1).Value = StockFishRateScalers(Stk, Fish, TStep).ToString("0")
               Else
                  StockFisheryGrid.Item(TStep + 1, Stk - 1).Value = StockFishRateScalers(Stk, Fish, TStep).ToString("###0.0000")
               End If
               StockFisheryGrid.Item(TStep + 1, Stk - 1).Style.BackColor = Color.White
            End If
         Next
      Next

   End Sub

End Class