Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Public Class FVS_FisheryScalerEdit
   Public Event CellEndEdit As DataGridViewCellEventHandler
   Public MarkSelectiveEdit As Boolean
   Public NumMSF As Integer

   Private Sub FisheryScalerGrid_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FisheryScalerGrid.CellEndEdit
      Dim TempFlag, TempFish, TempStep As Integer
      If MarkSelectiveEdit = True Then Exit Sub
      If e.ColumnIndex = 2 Or e.ColumnIndex = 5 Or e.ColumnIndex = 8 Or e.ColumnIndex = 11 Or e.ColumnIndex = 14 Then

         TempFlag = FisheryScalerGrid.Item(e.ColumnIndex, e.RowIndex).Value
         Select Case e.ColumnIndex
            Case 2
               TempStep = 1
            Case 5
               TempStep = 2
            Case 8
               TempStep = 3
            Case 11
               TempStep = 4
            Case 14
               TempStep = 5
         End Select
         TempFish = e.RowIndex + 1

         'If FisheryFlag(e.RowIndex + 1, e.ColumnIndex + 1) <> 7 And FisheryFlag(e.RowIndex + 1, e.ColumnIndex + 1) <> 8 Then
         If TempFlag = 7 Or TempFlag = 8 Then
            FisheryScalerGrid.Item(e.ColumnIndex + 1, e.RowIndex).Value = "MSF"
            FisheryScalerGrid.Item(e.ColumnIndex + 2, e.RowIndex).Value = "MSF"
            FisheryScalerGrid.Item(e.ColumnIndex + 1, e.RowIndex).Style.BackColor = Color.DeepPink
            FisheryScalerGrid.Item(e.ColumnIndex + 2, e.RowIndex).Style.BackColor = Color.DeepPink
         Else
            FisheryScalerGrid.Item(e.ColumnIndex + 1, e.RowIndex).Value = FisheryScaler(TempFish, TempStep).ToString("0.0000")
            FisheryScalerGrid.Item(e.ColumnIndex + 2, e.RowIndex).Value = FisheryQuota(TempFish, TempStep).ToString
            FisheryScalerGrid.Item(e.ColumnIndex + 1, e.RowIndex).Style.BackColor = Color.White
            FisheryScalerGrid.Item(e.ColumnIndex + 2, e.RowIndex).Style.BackColor = Color.White
         End If
         'End If
      End If
      'FisheryScalerGrid(e.ColumnIndex, e.RowIndex).Style SelectionBackColor = Color.Empty

   End Sub


   Private Sub FVS_FisheryScalerEdit_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      'Dim FisheryScalerEdit As DataGridView

      'Dim FisheryScalerGrid_CellEndEdit As DataGridViewCellEventHandler
      'AddHandler CellEndEdit, FisheryScalerGrid_CellEndEdit

      'FormHeight = 948
      FormHeight = 975
      FormWidth = 1257
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
         If FVS_FisheryScalerEdit_ReSize = False Then
            Resize_Form(Me)
            FVS_FisheryScalerEdit_ReSize = True
         End If
      End If

      MarkSelectiveEdit = False
      FSCancelButton.Visible = True
      LoadCatchButton.Visible = True
      LoadSheetButton.Visible = True
      FisheryScalerGrid.Columns.Clear()
      FisheryScalerGrid.Rows.Clear()
      FisheryScalerGrid.Columns.Clear()

      If FormWidthScaler > 1 Then
         FisheryScalerGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         FisheryScalerGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      Else
         FisheryScalerGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      End If
      If SpeciesName = "COHO" Then

         FisheryScalerGrid.Columns.Add("FisheryName", "Name")
         FisheryScalerGrid.Columns("FisheryName").Width = 120 / FormWidthScaler
         FisheryScalerGrid.Columns("FisheryName").ReadOnly = True
         FisheryScalerGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine

         'FisheryScalerGrid.ColumnHeadersBorderStyle
         FisheryScalerGrid.Columns.Add("FishNum", "#")
         FisheryScalerGrid.Columns("FishNum").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("FishNum").ReadOnly = True
         FisheryScalerGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

         FisheryScalerGrid.Columns.Add("Time1Flag", "Flg")
         FisheryScalerGrid.Columns("Time1Flag").Width = 35 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time1Scaler", "T1Scaler")
         FisheryScalerGrid.Columns("Time1Scaler").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time1Quota", "T1Quota")
         FisheryScalerGrid.Columns("Time1Quota").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time2Flag", "Flg")
         FisheryScalerGrid.Columns("Time2Flag").Width = 35 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time2Scaler", "T2Scaler")
         FisheryScalerGrid.Columns("Time2Scaler").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time2Quota", "T2Quota")
         FisheryScalerGrid.Columns("Time2Quota").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time3Flag", "Flg")
         FisheryScalerGrid.Columns("Time3Flag").Width = 35 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time3Scaler", "T3Scaler")
         FisheryScalerGrid.Columns("Time3Scaler").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time3Quota", "T3Quota")
         FisheryScalerGrid.Columns("Time3Quota").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time4Flag", "Flg")
         FisheryScalerGrid.Columns("Time4Flag").Width = 35 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time4Scaler", "T4Scaler")
         FisheryScalerGrid.Columns("Time4Scaler").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time4Quota", "T4Quota")
         FisheryScalerGrid.Columns("Time4Quota").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time5Flag", "Flg")
         FisheryScalerGrid.Columns("Time5Flag").Width = 35 / FormWidthScaler
         FisheryScalerGrid.Columns("Time5Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time5Scaler", "T5Scaler")
         FisheryScalerGrid.Columns("Time5Scaler").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time5Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time5Quota", "T5Quota")
         FisheryScalerGrid.Columns("Time5Quota").Width = 82 / FormWidthScaler
         FisheryScalerGrid.Columns("Time5Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.RowCount = NumFish

         For Fish As Integer = 1 To NumFish
            FisheryScalerGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
            FisheryScalerGrid.Item(1, Fish - 1).Value = Fish.ToString
            '- Time Step 1
            If AnyBaseRate(Fish, 1) = 1 Then
               FisheryScalerGrid.Item(2, Fish - 1).Value = FisheryFlag(Fish, 1)
               FisheryScalerGrid.Item(2, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 1) = 7 Or FisheryFlag(Fish, 1) = 8 Then
                  FisheryScalerGrid.Item(3, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(4, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(3, Fish - 1).Value = FisheryScaler(Fish, 1).ToString("####0.0000")
                  FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(4, Fish - 1).Value = CDbl(FisheryQuota(Fish, 1))
                  FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(2, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(2, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(3, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(4, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 2
            If AnyBaseRate(Fish, 2) = 1 Then
               FisheryScalerGrid.Item(5, Fish - 1).Value = FisheryFlag(Fish, 2)
               FisheryScalerGrid.Item(5, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 2) = 7 Or FisheryFlag(Fish, 2) = 8 Then
                  FisheryScalerGrid.Item(6, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(7, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(6, Fish - 1).Value = FisheryScaler(Fish, 2).ToString("####0.0000")
                  FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(7, Fish - 1).Value = CDbl(FisheryQuota(Fish, 2))
                  FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(5, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(5, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(6, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(7, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 3
            If AnyBaseRate(Fish, 3) = 1 Then
               FisheryScalerGrid.Item(8, Fish - 1).Value = FisheryFlag(Fish, 3)
               FisheryScalerGrid.Item(8, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 3) = 7 Or FisheryFlag(Fish, 3) = 8 Then
                  FisheryScalerGrid.Item(9, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(10, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(9, Fish - 1).Value = FisheryScaler(Fish, 3).ToString("####0.0000")
                  FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(10, Fish - 1).Value = CDbl(FisheryQuota(Fish, 3))
                  FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(8, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(8, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(9, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(10, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 4
            If AnyBaseRate(Fish, 4) = 1 Then
               FisheryScalerGrid.Item(11, Fish - 1).Value = FisheryFlag(Fish, 4)
               FisheryScalerGrid.Item(11, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 4) = 7 Or FisheryFlag(Fish, 4) = 8 Then
                  FisheryScalerGrid.Item(12, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(13, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(12, Fish - 1).Value = FisheryScaler(Fish, 4).ToString("####0.0000")
                  FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(13, Fish - 1).Value = CDbl(FisheryQuota(Fish, 4))
                  FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(11, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(11, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(12, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(13, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 5
            If AnyBaseRate(Fish, 5) = 1 Then
               FisheryScalerGrid.Item(14, Fish - 1).Value = FisheryFlag(Fish, 5)
               FisheryScalerGrid.Item(14, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 5) = 7 Or FisheryFlag(Fish, 5) = 8 Then
                  FisheryScalerGrid.Item(15, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(15, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(16, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(16, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(15, Fish - 1).Value = FisheryScaler(Fish, 5).ToString("####0.0000")
                  FisheryScalerGrid.Item(15, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(16, Fish - 1).Value = CDbl(FisheryQuota(Fish, 5))
                  FisheryScalerGrid.Item(16, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(14, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(14, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(15, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(15, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(16, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(16, Fish - 1).Style.BackColor = Color.Azure
            End If
         Next

      ElseIf SpeciesName = "CHINOOK" Then

         FisheryScalerGrid.Columns.Add("FisheryName", "Name")
         FisheryScalerGrid.Columns("FisheryName").Width = 250 / FormWidthScaler
         FisheryScalerGrid.Columns("FisheryName").ReadOnly = True
         FisheryScalerGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine

         FisheryScalerGrid.Columns.Add("FishNum", "#")
         FisheryScalerGrid.Columns("FishNum").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("FishNum").ReadOnly = True
         FisheryScalerGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

         FisheryScalerGrid.Columns.Add("Time1Flag", "Flg")
         FisheryScalerGrid.Columns("Time1Flag").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time1Scaler", "T1Scaler")
         FisheryScalerGrid.Columns("Time1Scaler").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time1Quota", "T1Quota")
         FisheryScalerGrid.Columns("Time1Quota").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time1Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time2Flag", "Flg")
         FisheryScalerGrid.Columns("Time2Flag").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time2Scaler", "T2Scaler")
         FisheryScalerGrid.Columns("Time2Scaler").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time2Quota", "T2Quota")
         FisheryScalerGrid.Columns("Time2Quota").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time2Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time3Flag", "Flg")
         FisheryScalerGrid.Columns("Time3Flag").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time3Scaler", "T3Scaler")
         FisheryScalerGrid.Columns("Time3Scaler").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time3Quota", "T3Quota")
         FisheryScalerGrid.Columns("Time3Quota").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time3Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.Columns.Add("Time4Flag", "Flg")
         FisheryScalerGrid.Columns("Time4Flag").Width = 40 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Flag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
         FisheryScalerGrid.Columns.Add("Time4Scaler", "T4Scaler")
         FisheryScalerGrid.Columns("Time4Scaler").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Scaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("Time4Quota", "T4Quota")
         FisheryScalerGrid.Columns("Time4Quota").Width = 85 / FormWidthScaler
         FisheryScalerGrid.Columns("Time4Quota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         FisheryScalerGrid.RowCount = NumFish

         For Fish As Integer = 1 To NumFish
            FisheryScalerGrid.Item(0, Fish - 1).Value = FisheryTitle(Fish)
            FisheryScalerGrid.Item(1, Fish - 1).Value = Fish.ToString
            '- Time Step 1
            If AnyBaseRate(Fish, 1) = 1 Then
               FisheryScalerGrid.Item(2, Fish - 1).Value = FisheryFlag(Fish, 1)
               FisheryScalerGrid.Item(2, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 1) = 7 Or FisheryFlag(Fish, 1) = 8 Then
                  FisheryScalerGrid.Item(3, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(4, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(3, Fish - 1).Value = FisheryScaler(Fish, 1).ToString("####0.0000")
                  FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(4, Fish - 1).Value = CLng(FisheryQuota(Fish, 1))
                  FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(2, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(2, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(3, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(3, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(4, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(4, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 2
            If AnyBaseRate(Fish, 2) = 1 Then
               FisheryScalerGrid.Item(5, Fish - 1).Value = FisheryFlag(Fish, 2)
               FisheryScalerGrid.Item(5, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 2) = 7 Or FisheryFlag(Fish, 2) = 8 Then
                  FisheryScalerGrid.Item(6, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(7, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(6, Fish - 1).Value = FisheryScaler(Fish, 2).ToString("####0.0000")
                  FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(7, Fish - 1).Value = CLng(FisheryQuota(Fish, 2))
                  FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(5, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(5, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(6, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(6, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(7, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(7, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 3
            If AnyBaseRate(Fish, 3) = 1 Then
               FisheryScalerGrid.Item(8, Fish - 1).Value = FisheryFlag(Fish, 3)
               FisheryScalerGrid.Item(8, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 3) = 7 Or FisheryFlag(Fish, 3) = 8 Then
                  FisheryScalerGrid.Item(9, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(10, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(9, Fish - 1).Value = FisheryScaler(Fish, 3).ToString("####0.0000")
                  FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(10, Fish - 1).Value = CLng(FisheryQuota(Fish, 3))
                  FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(8, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(8, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(9, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(9, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(10, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(10, Fish - 1).Style.BackColor = Color.Azure
            End If
            '- Time Step 4
            If AnyBaseRate(Fish, 4) = 1 Then
               FisheryScalerGrid.Item(11, Fish - 1).Value = FisheryFlag(Fish, 4)
               FisheryScalerGrid.Item(11, Fish - 1).Style.BackColor = Color.White
               If FisheryFlag(Fish, 4) = 7 Or FisheryFlag(Fish, 4) = 8 Then
                  FisheryScalerGrid.Item(12, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.DeepPink
                  FisheryScalerGrid.Item(13, Fish - 1).Value = "MSF"
                  FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.DeepPink
               Else
                  FisheryScalerGrid.Item(12, Fish - 1).Value = FisheryScaler(Fish, 4).ToString("####0.0000")
                  FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.White
                  FisheryScalerGrid.Item(13, Fish - 1).Value = CLng(FisheryQuota(Fish, 4))
                  FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.LavenderBlush
               End If
            Else
               FisheryScalerGrid.Item(11, Fish - 1).Value = "*"
               FisheryScalerGrid.Item(11, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(12, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(12, Fish - 1).Style.BackColor = Color.Azure
               FisheryScalerGrid.Item(13, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(13, Fish - 1).Style.BackColor = Color.Azure
            End If
         Next

      End If

   End Sub

   Private Sub FSCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSCancelButton.Click
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub FSDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSDoneButton.Click

      Dim AnyMarkFisheries As Boolean
      Dim MSFFish, MSFTStep, FlagVal As Integer

      '- After return from MSF Form, make changes and exit Form after Selective Fishery Inputs have been Editted
      If MarkSelectiveEdit = True Then
         NumMSF = 0
         For Fish As Integer = 1 To NumFish
            For TStep As Integer = 1 To NumSteps
               If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
                  NumMSF += 1
               Else
                  GoTo NextMTStep
               End If
               MSFFish = CInt(FisheryScalerGrid.Item(1, NumMSF - 1).Value)
               MSFTStep = CInt(FisheryScalerGrid.Item(2, NumMSF - 1).Value)
               If FisheryFlag(MSFFish, MSFTStep) <> CInt(FisheryScalerGrid.Item(3, NumMSF - 1).Value) Then
                  FisheryFlag(MSFFish, MSFTStep) = CInt(FisheryScalerGrid.Item(3, NumMSF - 1).Value)
                        ChangeFishScalers = True
               End If
                    If Math.Round(MSFFisheryScaler(MSFFish, MSFTStep), 4, MidpointRounding.AwayFromZero) <> CDbl(FisheryScalerGrid.Item(4, NumMSF - 1).Value) Then                  
                        MSFFisheryScaler(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(4, NumMSF - 1).Value)
                        ChangeFishScalers = True
                    End If
                    If Math.Round(MSFFisheryQuota(MSFFish, MSFTStep), 4, MidpointRounding.AwayFromZero) <> CDbl(FisheryScalerGrid.Item(5, NumMSF - 1).Value) Then
                        ChangeFishScalers = True
                        MSFFisheryQuota(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(5, NumMSF - 1).Value)
                    End If
        If MarkSelectiveMortRate(MSFFish, MSFTStep) <> CDbl(FisheryScalerGrid.Item(6, NumMSF - 1).Value) Then
            MarkSelectiveMortRate(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(6, NumMSF - 1).Value)
            ChangeFishScalers = True
        End If
        If MarkSelectiveMarkMisID(MSFFish, MSFTStep) <> CDbl(FisheryScalerGrid.Item(7, NumMSF - 1).Value) Then
            MarkSelectiveMarkMisID(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(7, NumMSF - 1).Value)
            ChangeFishScalers = True
        End If
        If MarkSelectiveUnMarkMisID(MSFFish, MSFTStep) <> CDbl(FisheryScalerGrid.Item(8, NumMSF - 1).Value) Then
            MarkSelectiveUnMarkMisID(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(8, NumMSF - 1).Value)
            ChangeFishScalers = True
        End If
        If MarkSelectiveIncRate(MSFFish, MSFTStep) <> CDbl(FisheryScalerGrid.Item(9, NumMSF - 1).Value) Then
            MarkSelectiveIncRate(MSFFish, MSFTStep) = CDbl(FisheryScalerGrid.Item(9, NumMSF - 1).Value)
            ChangeFishScalers = True
        End If
NextMTStep:
                Next
         Next
         FSCancelButton.Visible = True
         LoadCatchButton.Visible = True
         LoadSheetButton.Visible = True
         Me.Close()
         FVS_InputMenu.Visible = True
         '- Done with Second Pass for MSF Parameters ... Exit Sub
         Exit Sub
      End If

      AnyMarkFisheries = False
      '- Put DataGridView Variables back into Arrays
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            '- Zero out any "wrong" inputs (ie no base period fishery)
            If AnyBaseRate(Fish, TStep) = 0 Then
               If FisheryFlag(Fish, TStep) <> 0 Then
                  FisheryFlag(Fish, TStep) = 0
                  ChangeFishScalers = True
               End If
               If FisheryScaler(Fish, TStep) <> 0 Then
                  FisheryScaler(Fish, TStep) = 0
                  ChangeFishScalers = True
               End If
               If FisheryQuota(Fish, TStep) <> 0 Then
                  FisheryQuota(Fish, TStep) = 0
                  ChangeFishScalers = True
               End If
               GoTo NextTStep
            End If
            '- Check for Any Data Changes
            If FisheryFlag(Fish, TStep) <> CInt(FisheryScalerGrid.Item(TStep * 3 - 1, Fish - 1).Value) Then
               FlagVal = CInt(FisheryScalerGrid.Item(TStep * 3 - 1, Fish - 1).Value)
               '- Check for Valid Flag Value
               If FlagVal = 0 Or FlagVal = 1 Or FlagVal = 2 Or FlagVal = 7 Or FlagVal = 8 Or FlagVal = 17 Or FlagVal = 18 Or FlagVal = 27 Or FlagVal = 28 Then
                  FisheryFlag(Fish, TStep) = CInt(FisheryScalerGrid.Item(TStep * 3 - 1, Fish - 1).Value)
                  ChangeFishScalers = True
               Else
                  MsgBox("Bad Flag-Value Input !!!" & vbCrLf & "Fish=" & FisheryName(Fish) & " Time=" & TStep.ToString & vbCrLf & _
                         "Please change and re-save form" & "Use 0,1,2,7,8,17,18,27,28 for Flag!", vbOKOnly)
                  Exit Sub
               End If
            End If
            '- Don't Check Retention Values when MSF Flag Selected
            If FisheryFlag(Fish, TStep) <> 7 And FisheryFlag(Fish, TStep) <> 8 Then
                    If Math.Round(FisheryScaler(Fish, TStep), 4, MidpointRounding.AwayFromZero) <> CDbl(FisheryScalerGrid.Item(TStep * 3, Fish - 1).Value) Then
                        FisheryScaler(Fish, TStep) = CDbl(FisheryScalerGrid.Item(TStep * 3, Fish - 1).Value)
                        ChangeFishScalers = True
                    End If
                    If Math.Round(FisheryQuota(Fish, TStep), 4, MidpointRounding.AwayFromZero) <> CDbl(FisheryScalerGrid.Item(TStep * 3 + 1, Fish - 1).Value) Then
                        FisheryQuota(Fish, TStep) = CDbl(FisheryScalerGrid.Item(TStep * 3 + 1, Fish - 1).Value)
                        ChangeFishScalers = True
                    End If
            End If
            '- Check if Any MSF Selected
            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
               AnyMarkFisheries = True
            End If
NextTStep:
         Next
      Next
      '- "New" form for MSF
      If AnyMarkFisheries = True Then
         Call MarkSelectiveFisheryInputs()
         Exit Sub
      Else
         '- Exit form if No MSF
         Me.Close()
         FVS_InputMenu.Visible = True
         Exit Sub
      End If

   End Sub

   Sub MarkSelectiveFisheryInputs()

      MarkSelectiveEdit = True

      FSCancelButton.Visible = False
      LoadCatchButton.Visible = False
      LoadSheetButton.Visible = False

      FisheryScalerGrid.Columns.Clear()
      FisheryScalerGrid.Rows.Clear()

      FisheryScalerGrid.Columns.Add("FisheryName", "Fishery Name")
      FisheryScalerGrid.Columns("FisheryName").Width = 150
      FisheryScalerGrid.Columns("FisheryName").ReadOnly = True
      FisheryScalerGrid.Columns("FisheryName").DefaultCellStyle.BackColor = Color.Aquamarine

      FisheryScalerGrid.Columns.Add("FishNum", "Fish #")
      FisheryScalerGrid.Columns("FishNum").Width = 50
      FisheryScalerGrid.Columns("FishNum").ReadOnly = True
      FisheryScalerGrid.Columns("FishNum").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      FisheryScalerGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Aquamarine

      FisheryScalerGrid.Columns.Add("TimeStep", "Time Step")
      FisheryScalerGrid.Columns("TimeStep").Width = 50
      FisheryScalerGrid.Columns("TimeStep").ReadOnly = True
      FisheryScalerGrid.Columns("TimeStep").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      FisheryScalerGrid.Columns("TimeStep").DefaultCellStyle.BackColor = Color.Aquamarine

      FisheryScalerGrid.Columns.Add("MarkFlag", "Flag")
      FisheryScalerGrid.Columns("MarkFlag").Width = 47
      FisheryScalerGrid.Columns("MarkFlag").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      FisheryScalerGrid.Columns.Add("MarkScaler", "Scaler")
      FisheryScalerGrid.Columns("MarkScaler").Width = 130
      FisheryScalerGrid.Columns("MarkScaler").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      FisheryScalerGrid.Columns.Add("MarkQuota", "Quota")
      FisheryScalerGrid.Columns("MarkQuota").Width = 90
      FisheryScalerGrid.Columns("MarkQuota").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      FisheryScalerGrid.Columns.Add("MarkRelease", "Release Rate")
      FisheryScalerGrid.Columns("MarkRelease").Width = 90
      FisheryScalerGrid.Columns("MarkRelease").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      FisheryScalerGrid.Columns.Add("MarkMisID", "Marked MisID")
      FisheryScalerGrid.Columns("MarkMisID").Width = 90
      FisheryScalerGrid.Columns("MarkMisID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      FisheryScalerGrid.Columns.Add("UnMarkMisID", "UnMark MisID")
      FisheryScalerGrid.Columns("UnMarkMisID").Width = 90
      FisheryScalerGrid.Columns("UnMarkMisID").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      FisheryScalerGrid.Columns.Add("MarkDropOff", "DropOff Rate")
      FisheryScalerGrid.Columns("MarkDropOff").Width = 90
      FisheryScalerGrid.Columns("MarkDropOff").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

      '- Count Number of Mark-Selective Fishery/Time-Steps
      NumMSF = 0
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
               NumMSF += 1
            End If
         Next
      Next

      FisheryScalerGrid.RowCount = NumMSF
      NumMSF = 0
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
               FisheryScalerGrid.Item(0, NumMSF).Value = FisheryName(Fish)
               FisheryScalerGrid.Item(1, NumMSF).Value = Fish.ToString
               FisheryScalerGrid.Item(2, NumMSF).Value = TStep.ToString
               FisheryScalerGrid.Item(3, NumMSF).Value = FisheryFlag(Fish, TStep).ToString("0")
               FisheryScalerGrid.Item(4, NumMSF).Value = MSFFisheryScaler(Fish, TStep).ToString("###0.0000")
                    FisheryScalerGrid.Item(5, NumMSF).Value = CDbl(MSFFisheryQuota(Fish, TStep)).ToString("######0.0000")
               FisheryScalerGrid.Item(6, NumMSF).Value = MarkSelectiveMortRate(Fish, TStep).ToString("0.0000")
               FisheryScalerGrid.Item(7, NumMSF).Value = MarkSelectiveMarkMisID(Fish, TStep).ToString("0.0000")
               FisheryScalerGrid.Item(8, NumMSF).Value = MarkSelectiveUnMarkMisID(Fish, TStep).ToString("0.0000")
               FisheryScalerGrid.Item(9, NumMSF).Value = MarkSelectiveIncRate(Fish, TStep).ToString("0.0000")
               NumMSF += 1
            End If
         Next
      Next

   End Sub

   Private Sub MenuStrip1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      '- The Clipboard Copy Column Names are specific to the Number of Time Steps
      '- for each of the Species - 5 for Coho and 4 for Chinook
      If MarkSelectiveEdit = False And SpeciesName = "COHO" Then
         '- Clipboard Copy for COHO Quota/Scaler Screen
         ClipStr = ""
         Clipboard.Clear()
         ClipStr = "FisheryName" & vbTab & "FishID" & vbTab & "T1-Flag" & vbTab & "T1-Scaler" & vbTab & "T1-Quota" & vbTab & "T2-Flag" & vbTab & "T2-Scaler" & vbTab & "T2-Quota" & vbTab & "T3-Flag" & vbTab & "T3-Scaler" & vbTab & "T3-Quota" & vbTab & "T4-Flag" & vbTab & "T4-Scaler" & vbTab & "T4-Quota" & vbTab & "T5-Flag" & vbTab & "T5-Scaler" & vbTab & "T5-Quota" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To 16
               If ColNum = 0 Then
                  ClipStr = ClipStr & FisheryScalerGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & FisheryScalerGrid.Item(ColNum, RecNum).Value
               End If
            Next
            ClipStr = ClipStr & vbCr
         Next
         Clipboard.SetDataObject(ClipStr)
      ElseIf MarkSelectiveEdit = False And SpeciesName = "CHINOOK" Then
         '- Clipboard Copy for CHINOOK Quota/Scaler Screen
         ClipStr = ""
         Clipboard.Clear()
         ClipStr = "FisheryName" & vbTab & "FishID" & vbTab & "T1-Flag" & vbTab & "T1-Scaler" & vbTab & "T1-Quota" & vbTab & "T2-Flag" & vbTab & "T2-Scaler" & vbTab & "T2-Quota" & vbTab & "T3-Flag" & vbTab & "T3-Scaler" & vbTab & "T3-Quota" & vbTab & "T4-Flag" & vbTab & "T4-Scaler" & vbTab & "T4-Quota" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To 13
               If ColNum = 0 Then
                  ClipStr = ClipStr & FisheryScalerGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & FisheryScalerGrid.Item(ColNum, RecNum).Value
               End If
            Next
            ClipStr = ClipStr & vbCr
         Next
         Clipboard.SetDataObject(ClipStr)
      Else
         '- Clipboard Copy for Mark Selective Fishery Screen
         ClipStr = ""
         Clipboard.Clear()
         ClipStr = "FisheryName" & vbTab & "FishID" & vbTab & "TimeStep" & vbTab & "MSF Flag" & vbTab & "MSFScaler" & vbTab & "MSFQuota" & vbTab & "RelMrtRate" & vbTab & "MarkMisID" & vbTab & "UnMrkMisID" & vbTab & "DropOffRt" & vbCr
         For RecNum = 0 To NumMSF - 1
            For ColNum = 0 To 9
               If ColNum = 0 Then
                  ClipStr = ClipStr & FisheryScalerGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & FisheryScalerGrid.Item(ColNum, RecNum).Value
               End If
            Next
            ClipStr = ClipStr & vbCr
         Next
         Clipboard.SetDataObject(ClipStr)
      End If

   End Sub

   Private Sub LoadCatchButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadCatchButton.Click

      Dim OpenFRAMCatchSpreadsheet As New OpenFileDialog()
      Dim FRAMCatchSpreadSheet, FRAMCatchSpreadSheetPath As String

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenFRAMCatchSpreadsheet.Filter = "FRAM-Catch Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenFRAMCatchSpreadsheet.FilterIndex = 1
      OpenFRAMCatchSpreadsheet.RestoreDirectory = True

      If OpenFRAMCatchSpreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         FRAMCatchSpreadSheet = OpenFRAMCatchSpreadsheet.FileName
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
            'xlApp.Workbooks.Close()
            'xlWorkBook = xlApp.Workbooks.Open(FRAMCatchSpreadSheet)
            'xlWorkBook.Activate()
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
            If xlWorkSheet.Name = "FRAMInput" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAMInput" Then
         MsgBox("Can't Find 'FRAMInput' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate Spreadsheet with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
            MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK Spreadsheet with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
            MsgBox("Can't Find 'No Cal Trm' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate COHO Spreadsheet with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress1 As String
      Dim CellAddress2 As String
      Dim FlagAddress As String
      Dim FlagValue As Integer
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            CellAddress1 = ""
            CellAddress2 = ""
            FlagAddress = ""
            If SpeciesName = "CHINOOK" Then
               Select Case TStep
                  Case 1
                     FlagAddress = "C" & CStr(Fish + 3)
                     CellAddress1 = "D" & CStr(Fish + 3)
                     CellAddress2 = "E" & CStr(Fish + 3)
                  Case 2
                     FlagAddress = "F" & CStr(Fish + 3)
                     CellAddress1 = "G" & CStr(Fish + 3)
                     CellAddress2 = "H" & CStr(Fish + 3)
                  Case 3
                     FlagAddress = "I" & CStr(Fish + 3)
                     CellAddress1 = "J" & CStr(Fish + 3)
                     CellAddress2 = "K" & CStr(Fish + 3)
                  Case 4
                     FlagAddress = "L" & CStr(Fish + 3)
                     CellAddress1 = "M" & CStr(Fish + 3)
                     CellAddress2 = "N" & CStr(Fish + 3)
               End Select
            ElseIf SpeciesName = "COHO" Then
               Select Case TStep
                  Case 1
                     FlagAddress = "D" & CStr(Fish + 3)
                     CellAddress1 = "E" & CStr(Fish + 3)
                     CellAddress2 = "F" & CStr(Fish + 3)
                  Case 2
                     FlagAddress = "G" & CStr(Fish + 3)
                     CellAddress1 = "H" & CStr(Fish + 3)
                     CellAddress2 = "I" & CStr(Fish + 3)
                  Case 3
                     FlagAddress = "J" & CStr(Fish + 3)
                     CellAddress1 = "K" & CStr(Fish + 3)
                     CellAddress2 = "L" & CStr(Fish + 3)
                  Case 4
                     FlagAddress = "M" & CStr(Fish + 3)
                     CellAddress1 = "N" & CStr(Fish + 3)
                     CellAddress2 = "O" & CStr(Fish + 3)
                  Case 5
                     FlagAddress = "P" & CStr(Fish + 3)
                     CellAddress1 = "Q" & CStr(Fish + 3)
                     CellAddress2 = "R" & CStr(Fish + 3)
               End Select
            End If

            If IsNumeric(xlWorkSheet.Range(FlagAddress).Value) Then
               FlagValue = xlWorkSheet.Range(FlagAddress).Value
               '- Note : This has changed from Old FRAM now that 1=scaler and 2=quota
               'If FlagValue = 2 Or FlagValue = 8 Then
               If IsNumeric(xlWorkSheet.Range(CellAddress2).Value) Then
                  If CInt(xlWorkSheet.Range(CellAddress2).Value) < 0 Or CInt(xlWorkSheet.Range(CellAddress2).Value) > 999999 Then GoTo NextBFCatch
                        FisheryQuota(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress2).Value).ToString("######0.0000")
               End If
               'ElseIf FlagValue = 1 Or FlagValue = 7 Then
               If IsNumeric(xlWorkSheet.Range(CellAddress1).Value) Then
                  If CInt(xlWorkSheet.Range(CellAddress1).Value) < 0 Or CInt(xlWorkSheet.Range(CellAddress1).Value) > 9999 Then GoTo NextBFCatch
                  FisheryScaler(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
               End If
               'End If
               FisheryFlag(Fish, TStep) = FlagValue
            End If
NextBFCatch:
         Next
      Next

      '- Check if Selective Fishery Parameters should be Updated
      Dim Result
      Result = MsgBox("Update Selective Fishery parameters from SpreadSheet?", MsgBoxStyle.YesNo)
      If Result = vbNo Then GoTo SkipMSFUpdate

      '- Find WorkSheets with FRAM_MSF numbers
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name = "FRAM_MSF" Then Exit For
         End If
      Next
      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAM_MSF" Then
         MsgBox("Can't Find 'FRAM_MSF' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate Spreadsheet with FRAM_MSF WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If
      Dim RowNum As Integer
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If FisheryFlag(Fish, TStep) = 7 Or FisheryFlag(Fish, TStep) = 8 Or FisheryFlag(Fish, TStep) = 17 Or FisheryFlag(Fish, TStep) = 18 Or FisheryFlag(Fish, TStep) = 27 Or FisheryFlag(Fish, TStep) = 28 Then
               For RowNum = 4 To 134
                  CellAddress1 = "B" & CStr(RowNum)
                  FlagAddress = "C" & CStr(RowNum)
                  If Fish = CInt(xlWorkSheet.Range(CellAddress1).Value) And TStep = CInt(xlWorkSheet.Range(FlagAddress).Value) Then
                     '- Found Correct Row with MSF Parameters
                     CellAddress1 = "E" & CStr(RowNum)
                     MSFFisheryScaler(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
                     CellAddress1 = "F" & CStr(RowNum)
                            MSFFisheryQuota(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value).ToString("######0.0000")
                     CellAddress1 = "G" & CStr(RowNum)
                     MarkSelectiveMortRate(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
                     CellAddress1 = "H" & CStr(RowNum)
                     MarkSelectiveMarkMisID(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
                     CellAddress1 = "I" & CStr(RowNum)
                     MarkSelectiveUnMarkMisID(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
                     CellAddress1 = "J" & CStr(RowNum)
                     MarkSelectiveIncRate(Fish, TStep) = CDbl(xlWorkSheet.Range(CellAddress1).Value)
                     Exit For
                  End If
               Next
            End If
         Next
      Next
SkipMSFUpdate:

      '- ReLoad DataGridView with New Values
      FVS_FisheryScalerEdit_Load(sender, e)
      ChangeFishScalers = True

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


   Private Sub LoadSheetButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadSheetButton.Click

      Dim OpenFRAMCatchSpreadsheet As New OpenFileDialog()
      Dim FRAMCatchSpreadSheet, FRAMCatchSpreadSheetPath As String

      '- Test if Excel was Running
      ExcelWasNotRunning = True
      Try
         xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
         ExcelWasNotRunning = False
      Catch ex As Exception
         xlApp = New Microsoft.Office.Interop.Excel.Application()
      End Try

      OpenFRAMCatchSpreadsheet.Filter = "FRAM-Catch Spreadsheets (*.xls)|*.xls|All files (*.*)|*.*"
      OpenFRAMCatchSpreadsheet.FilterIndex = 1
      OpenFRAMCatchSpreadsheet.RestoreDirectory = True

      If OpenFRAMCatchSpreadsheet.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         FRAMCatchSpreadSheet = OpenFRAMCatchSpreadsheet.FileName
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
            If xlWorkSheet.Name = "FRAMInput" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAMInput Worksheet
      If xlWorkSheet.Name <> "FRAMInput" Then
         MsgBox("Can't Find 'FRAMInput' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Check first Fishery Name for correct Species Spreadsheet
      Dim testname As String
      testname = xlWorkSheet.Range("A4").Value
      If SpeciesName = "CHINOOK" Then
         If Trim(xlWorkSheet.Range("A4").Value) <> "SE Alaska Troll" Then
            MsgBox("Can't Find 'SE Alaska Troll' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate CHINOOK DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      ElseIf SpeciesName = "COHO" Then
         If xlWorkSheet.Range("A4").Value <> "No Cal Trm" Then
            MsgBox("Can't Find 'No Cal Trm' as first Fishery your Spreadsheet Selection" & vbCrLf & _
                   "Please Choose appropriate COHO DataBase with FRAM Catch WorkSheet!", MsgBoxStyle.OkOnly)
            GoTo CloseExcelWorkBook
         End If
      End If

      '- Load WorkSheet Catch into Quota Array (Change Flag)
      Me.Cursor = Cursors.WaitCursor
      Dim CellAddress1, CellAddress2 As String
      Dim FlagAddress As String
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            CellAddress1 = ""
            CellAddress2 = ""
            FlagAddress = ""
            If SpeciesName = "CHINOOK" Then
               Select Case TStep
                  Case 1
                     FlagAddress = "C" & CStr(Fish + 3)
                     CellAddress1 = "D" & CStr(Fish + 3)
                     CellAddress2 = "E" & CStr(Fish + 3)
                  Case 2
                     FlagAddress = "F" & CStr(Fish + 3)
                     CellAddress1 = "G" & CStr(Fish + 3)
                     CellAddress2 = "H" & CStr(Fish + 3)
                  Case 3
                     FlagAddress = "I" & CStr(Fish + 3)
                     CellAddress1 = "J" & CStr(Fish + 3)
                     CellAddress2 = "K" & CStr(Fish + 3)
                  Case 4
                     FlagAddress = "L" & CStr(Fish + 3)
                     CellAddress1 = "M" & CStr(Fish + 3)
                     CellAddress2 = "N" & CStr(Fish + 3)
                  Case 5
                     FlagAddress = "O" & CStr(Fish + 3)
                     CellAddress1 = "P" & CStr(Fish + 3)
                     CellAddress2 = "Q" & CStr(Fish + 3)
               End Select
            Else
               Select Case TStep
                  Case 1
                     FlagAddress = "D" & CStr(Fish + 3)
                     CellAddress1 = "E" & CStr(Fish + 3)
                     CellAddress2 = "F" & CStr(Fish + 3)
                  Case 2
                     FlagAddress = "G" & CStr(Fish + 3)
                     CellAddress1 = "H" & CStr(Fish + 3)
                     CellAddress2 = "I" & CStr(Fish + 3)
                  Case 3
                     FlagAddress = "J" & CStr(Fish + 3)
                     CellAddress1 = "K" & CStr(Fish + 3)
                     CellAddress2 = "L" & CStr(Fish + 3)
                  Case 4
                     FlagAddress = "M" & CStr(Fish + 3)
                     CellAddress1 = "N" & CStr(Fish + 3)
                     CellAddress2 = "O" & CStr(Fish + 3)
                  Case 5
                     FlagAddress = "P" & CStr(Fish + 3)
                     CellAddress1 = "Q" & CStr(Fish + 3)
                     CellAddress2 = "R" & CStr(Fish + 3)
               End Select
            End If
            If AnyBaseRate(Fish, TStep) = 0 Then
               xlWorkSheet.Range(CellAddress1).Value = "*******"
               xlWorkSheet.Range(CellAddress1).Interior.Color = RGB(148, 150, 232)
               xlWorkSheet.Range(CellAddress2).Value = "*******"
               xlWorkSheet.Range(CellAddress2).Interior.Color = RGB(148, 150, 232)
               xlWorkSheet.Range(FlagAddress).Value = "*"
               xlWorkSheet.Range(FlagAddress).Interior.Color = RGB(148, 150, 232)
               GoTo NextBFCatch
            End If
            xlWorkSheet.Range(FlagAddress).Value = FisheryFlag(Fish, TStep).ToString
                'xlWorkSheet.Range(FlagAddress).Interior.Color = RGB(255, 255, 255)
            xlWorkSheet.Range(CellAddress1).Value = FisheryScaler(Fish, TStep).ToString("#####0.0000")
                'xlWorkSheet.Range(CellAddress1).Interior.Color = RGB(255, 255, 255)
                xlWorkSheet.Range(CellAddress2).Value = FisheryQuota(Fish, TStep).ToString("######0.00000")
                ' xlWorkSheet.Range(CellAddress2).Interior.Color = RGB(255, 255, 255)
NextBFCatch:
         Next
      Next

      '- Find WorkSheets with FRAM MSF Numbers
      For Each xlWorkSheet In xlWorkBook.Worksheets
         If xlWorkSheet.Name.Length > 7 Then
            If xlWorkSheet.Name = "FRAM_MSF" Then Exit For
         End If
      Next

      '- Check if DataBase contains FRAM MSF Worksheet
      If xlWorkSheet.Name <> "FRAM_MSF" Then
         MsgBox("Can't Find 'FRAM_MSF' WorkSheet in your Spreadsheet Selection" & vbCrLf & _
                "Please Choose appropriate DataBase with FRAM MSF WorkSheet!", MsgBoxStyle.OkOnly)
         GoTo CloseExcelWorkBook
      End If

      '- Load WorkSheet with MSF Arrays
      Dim NumMSF As Integer
      NumMSF = 0
      Me.Cursor = Cursors.WaitCursor
      xlWorkSheet.Range("A4:J591").ClearContents()
      For Fish As Integer = 1 To NumFish
         For TStep As Integer = 1 To NumSteps
            If (FisheryFlag(Fish, TStep) <= 2) Then GoTo NextMSFCatch
            NumMSF += 1
            CellAddress1 = "A" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = FisheryTitle(Fish)
            CellAddress1 = "B" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = Fish.ToString
            CellAddress1 = "C" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = TStep.ToString
            CellAddress1 = "D" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = FisheryFlag(Fish, TStep)
            CellAddress1 = "E" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MSFFisheryScaler(Fish, TStep).ToString("####0.0000")
            CellAddress1 = "F" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MSFFisheryQuota(Fish, TStep).ToString("######0.00000")
            CellAddress1 = "G" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MarkSelectiveMortRate(Fish, TStep).ToString("0.0000")
            CellAddress1 = "H" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MarkSelectiveMarkMisID(Fish, TStep).ToString("0.0000")
            CellAddress1 = "I" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MarkSelectiveUnMarkMisID(Fish, TStep).ToString("0.0000")
            CellAddress1 = "J" & CStr(NumMSF + 3)
            xlWorkSheet.Range(CellAddress1).Value = MarkSelectiveIncRate(Fish, TStep).ToString("0.0000")
NextMSFCatch:
         Next
      Next

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

   'Private Sub FVS_FisheryScalerEdit_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LostFocus
   '   Dim jim1 As Integer
   '   jim1 = FisheryScalerGrid.CurrentCell.ColumnIndex
   'End Sub

   Private Sub FisheryScalerGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles FisheryScalerGrid.CellContentClick

   End Sub

Private Sub ClipboardCopyToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ClipboardCopyToolStripMenuItem.Click

End Sub
End Class