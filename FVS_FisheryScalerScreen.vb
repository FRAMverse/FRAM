Public Class FVS_FisheryScalerScreen

   Private Sub FVS_FisheryScalerScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      'FormHeight = 898
      FormHeight = 950
      FormWidth = 934
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
         If FVS_FisheryScalerScreen_ReSize = False Then
            Resize_Form(Me)
            FVS_FisheryScalerScreen_ReSize = True
         End If
      End If

      FisheryScalerGrid.Columns.Clear()
      FisheryScalerGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      FisheryScalerGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

      If SpeciesName = "CHINOOK" Then
         FisheryScalerGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         FisheryScalerGrid.Columns.Add("Name", "FisheryName")
         FisheryScalerGrid.Columns(0).Width = 175
         FisheryScalerGrid.Columns(0).ReadOnly = True
         FisheryScalerGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         FisheryScalerGrid.Columns.Add("T1", "Oct-Apr1")
         FisheryScalerGrid.Columns(1).Width = 100
         FisheryScalerGrid.Columns(1).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T2", "May-June")
         FisheryScalerGrid.Columns(2).Width = 100
         FisheryScalerGrid.Columns(2).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T3", "July-Sept")
         FisheryScalerGrid.Columns(3).Width = 100
         FisheryScalerGrid.Columns(3).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T4", "Oct-Apr2")
         FisheryScalerGrid.Columns(4).Width = 100
         FisheryScalerGrid.Columns(4).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.RowCount = NumFish
      ElseIf SpeciesName = "COHO" Then
         FisheryScalerGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         FisheryScalerGrid.Columns.Add("Name", "FisheryName")
         FisheryScalerGrid.Columns(0).Width = 175
         FisheryScalerGrid.Columns(0).ReadOnly = True
         FisheryScalerGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         FisheryScalerGrid.Columns.Add("T1", "Jan-June")
         FisheryScalerGrid.Columns(1).Width = 100
         FisheryScalerGrid.Columns(1).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T2", "July")
         FisheryScalerGrid.Columns(2).Width = 100
         FisheryScalerGrid.Columns(2).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T3", "August")
         FisheryScalerGrid.Columns(3).Width = 100
         FisheryScalerGrid.Columns(3).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T4", "September")
         FisheryScalerGrid.Columns(4).Width = 100
         FisheryScalerGrid.Columns(4).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.Columns.Add("T5", "Oct-Dec")
         FisheryScalerGrid.Columns(5).Width = 100
         FisheryScalerGrid.Columns(5).DefaultCellStyle.Format = ("###0.0000")
         FisheryScalerGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         FisheryScalerGrid.RowCount = NumFish
      End If
      '- Put Fishery Scaler Array into Grid
      For Fish As Integer = 1 To NumFish
         FisheryScalerGrid.Item(0, Fish - 1).Value = FisheryName(Fish)
         For TStep As Integer = 1 To NumSteps
            If AnyBaseRate(Fish, TStep) = 1 Then
               If FisheryScaler(Fish, TStep) = 0 Then
                  FisheryScalerGrid.Item(TStep, Fish - 1).Value = "0"
               Else
                  FisheryScalerGrid.Item(TStep, Fish - 1).Value = FisheryScaler(Fish, TStep).ToString("###0.0000")
               End If
            Else
               FisheryScalerGrid.Item(TStep, Fish - 1).Value = "****"
               FisheryScalerGrid.Item(TStep, Fish - 1).Style.BackColor = Color.LightBlue
            End If
         Next
      Next

   End Sub

   Private Sub FSExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FSExitButton.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
   End Sub

   Private Sub ClipBoardCopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClipBoardCopyToolStripMenuItem.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      If SpeciesName = "CHINOOK" Then
         ClipStr = "CHINOOK "
      ElseIf SpeciesName = "COHO" Then
         ClipStr = "COHO "
      End If
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr
      If SpeciesName = "CHINOOK" Then
         ClipStr &= "FisheryName" & vbTab & "Oct-Apr1" & vbTab & "May-June" & vbTab & "July-Sept" & vbTab & "Oct-Apr2" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To NumSteps
               If ColNum = 0 Then
                  ClipStr &= FisheryScalerGrid.Item(ColNum, RecNum).Value
               Else
                  If FisheryScalerGrid.Item(ColNum, RecNum).Value = "****" Then
                     ClipStr &= vbTab & "****"
                  Else
                     ClipStr &= vbTab & CDbl(FisheryScalerGrid.Item(ColNum, RecNum).Value)
                  End If
               End If
            Next
            ClipStr &= vbCr
         Next
      ElseIf SpeciesName = "COHO" Then
         ClipStr &= "FisheryName" & vbTab & "Jan-June" & vbTab & "July" & vbTab & "August" & vbTab & "September" & vbTab & "Oct-Dec" & vbCr
         For RecNum = 0 To NumFish - 1
            For ColNum = 0 To NumSteps
               If ColNum = 0 Then
                  ClipStr = ClipStr & FisheryScalerGrid.Item(ColNum, RecNum).Value
               Else
                  If FisheryScalerGrid.Item(ColNum, RecNum).Value = "****" Then
                     ClipStr &= vbTab & "****"
                  Else
                     ClipStr &= vbTab & CDbl(FisheryScalerGrid.Item(ColNum, RecNum).Value)
                  End If
                  'ClipStr = ClipStr & vbTab & CDbl(FisheryScalerGrid.Item(ColNum, RecNum).Value)
               End If
            Next
            ClipStr &= vbCr
         Next
      End If
      Clipboard.SetDataObject(ClipStr)

   End Sub
End Class