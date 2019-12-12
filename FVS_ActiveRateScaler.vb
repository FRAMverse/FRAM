Public Class FVS_ActiveRateScaler
    Public GridLoading As Boolean

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Visible = False
        FVS_InputMenu.Visible = True
        FVS_InputMenu.Refresh()
    End Sub

    Private Sub FVS_ActiveRateScaler_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        FormHeight = 847
        FormWidth = 1049
        Dim fish As Integer

        ' - Check if Form fits within Screen Dimensions
        If (FormHeight > My.Computer.Screen.Bounds.Height Or _
            FormWidth > My.Computer.Screen.Bounds.Width) Then
            Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
            Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
            If FVS_ActiveRateScaler_ReSize = False Then
                Resize_Form(Me)
                FVS_ActiveRateScaler_ReSize = True
            End If
        End If

        GridLoading = True
        StkFishRateScalerGrid.Columns.Clear()
        StkFishRateScalerGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
        StkFishRateScalerGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        StkFishRateScalerGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)

        StkFishRateScalerGrid.Columns.Add("Fishery", "Fishery")
        StkFishRateScalerGrid.Columns("Fishery").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Fishery").ReadOnly = True
        StkFishRateScalerGrid.Columns("Fishery").DefaultCellStyle.BackColor = Color.Aquamarine

        StkFishRateScalerGrid.Columns.Add("FishNum", "Fish#")
        StkFishRateScalerGrid.Columns("FishNum").Width = 50 / FormWidthScaler
        StkFishRateScalerGrid.Columns("FishNum").ReadOnly = True
        StkFishRateScalerGrid.Columns("FishNum").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        StkFishRateScalerGrid.Columns("FishNum").DefaultCellStyle.BackColor = Color.Azure

        StkFishRateScalerGrid.Columns.Add("Stock", "Stock")
        StkFishRateScalerGrid.Columns("Stock").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Stock").ReadOnly = True
        StkFishRateScalerGrid.Columns("Stock").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        StkFishRateScalerGrid.Columns("Stock").DefaultCellStyle.BackColor = Color.Azure

        StkFishRateScalerGrid.Columns.Add("StkNum", "Stk#")
        StkFishRateScalerGrid.Columns("StkNum").Width = 50 / FormWidthScaler
        StkFishRateScalerGrid.Columns("StkNum").ReadOnly = True
        StkFishRateScalerGrid.Columns("StkNum").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        StkFishRateScalerGrid.Columns("StkNum").DefaultCellStyle.BackColor = Color.Azure


        StkFishRateScalerGrid.Columns.Add("Time1", "Time1")
        StkFishRateScalerGrid.Columns("Time1").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Time1").DefaultCellStyle.Format = ("###0.0000")
        StkFishRateScalerGrid.Columns("Time1").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        StkFishRateScalerGrid.Columns.Add("Time2", "Time2")
        StkFishRateScalerGrid.Columns("Time2").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Time2").DefaultCellStyle.Format = ("###0.0000")
        StkFishRateScalerGrid.Columns("Time2").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        StkFishRateScalerGrid.Columns.Add("Time3", "Time3")
        StkFishRateScalerGrid.Columns("Time3").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Time3").DefaultCellStyle.Format = ("###0.0000")
        StkFishRateScalerGrid.Columns("Time3").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        StkFishRateScalerGrid.Columns.Add("Time4", "Time4")
        StkFishRateScalerGrid.Columns("Time4").Width = 100 / FormWidthScaler
        StkFishRateScalerGrid.Columns("Time4").DefaultCellStyle.Format = ("###0.0000")
        StkFishRateScalerGrid.Columns("Time4").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter

        If SpeciesName = "COHO" Then
            StkFishRateScalerGrid.Columns.Add("Time5", "Time5")
            StkFishRateScalerGrid.Columns("Time5").Width = 100 / FormWidthScaler
            StkFishRateScalerGrid.Columns("Time5").DefaultCellStyle.Format = ("###0.0000")
            StkFishRateScalerGrid.Columns("Time5").DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
        End If



        Dim rowcount As Integer = 0
        If SpeciesName = "COHO" Then
            For fish = 1 To NumFish
                For Stk As Integer = 1 To NumStk
                    If StockFishRateScalers(Stk, fish, 1) <> 1.0 Or StockFishRateScalers(Stk, fish, 2) <> 1.0 Or StockFishRateScalers(Stk, fish, 3) <> 1.0 _
                    Or StockFishRateScalers(Stk, fish, 4) <> 1.0 Or StockFishRateScalers(Stk, fish, 5) <> 1 Then
                        rowcount = rowcount + 1
                    End If
                Next
            Next

        Else
            For fish = 1 To NumFish
                For Stk As Integer = 1 To NumStk
                    If StockFishRateScalers(Stk, fish, 1) <> 1.0 Or StockFishRateScalers(Stk, fish, 2) <> 1.0 Or StockFishRateScalers(Stk, fish, 3) <> 1.0 Or StockFishRateScalers(Stk, fish, 4) <> 1.0 Then
                        rowcount = rowcount + 1
                    End If
                Next
            Next
        End If

        StkFishRateScalerGrid.RowCount = rowcount + 1

        rowcount = 0
        If SpeciesName = "COHO" Then
            For fish = 1 To NumFish
                For Stk As Integer = 1 To NumStk
                    If StockFishRateScalers(Stk, fish, 1) <> 1.0 Or StockFishRateScalers(Stk, fish, 2) <> 1.0 Or StockFishRateScalers(Stk, fish, 3) <> 1.0 _
                    Or StockFishRateScalers(Stk, fish, 4) <> 1.0 Or StockFishRateScalers(Stk, fish, 5) <> 1 Then
                        For TStep = 1 To NumSteps
                            If AnyBaseRate(fish, TStep) = 0 Then
                                StkFishRateScalerGrid.Item(0, rowcount).Value = FisheryName(fish)
                                StkFishRateScalerGrid.Item(1, rowcount).Value = FisheryID(fish)
                                StkFishRateScalerGrid.Item(2, rowcount).Value = StockName(Stk)
                                StkFishRateScalerGrid.Item(3, rowcount).Value = StockID(Stk)
                                StkFishRateScalerGrid.Item(TStep + 3, rowcount).Value = "****"
                            Else
                                StkFishRateScalerGrid.Item(0, rowcount).Value = FisheryName(fish)
                                StkFishRateScalerGrid.Item(1, rowcount).Value = FisheryID(fish)
                                StkFishRateScalerGrid.Item(2, rowcount).Value = StockName(Stk)
                                StkFishRateScalerGrid.Item(3, rowcount).Value = StockID(Stk)
                                StkFishRateScalerGrid.Item(TStep + 3, rowcount).Value = StockFishRateScalers(Stk, fish, TStep).ToString("###0.0000")
                            End If

                        Next TStep
                        rowcount = rowcount + 1
                    End If
                Next
            Next

        Else
            For fish = 1 To NumFish
                For Stk As Integer = 1 To NumStk
                    If StockFishRateScalers(Stk, fish, 1) <> 1.0 Or StockFishRateScalers(Stk, fish, 2) <> 1.0 Or StockFishRateScalers(Stk, fish, 3) <> 1.0 Or StockFishRateScalers(Stk, fish, 4) <> 1.0 Then
                        For TStep = 1 To NumSteps
                            If AnyBaseRate(fish, TStep) = 0 Then
                                StkFishRateScalerGrid.Item(0, rowcount).Value = FisheryName(fish)
                                StkFishRateScalerGrid.Item(1, rowcount).Value = FisheryID(fish)
                                StkFishRateScalerGrid.Item(2, rowcount).Value = StockName(Stk)
                                StkFishRateScalerGrid.Item(3, rowcount).Value = StockID(Stk)
                                StkFishRateScalerGrid.Item(TStep + 3, rowcount).Value = "****"
                               
                            Else
                                StkFishRateScalerGrid.Item(0, rowcount).Value = FisheryName(fish)
                                StkFishRateScalerGrid.Item(1, rowcount).Value = FisheryID(fish)
                                StkFishRateScalerGrid.Item(2, rowcount).Value = StockName(Stk)
                                StkFishRateScalerGrid.Item(3, rowcount).Value = StockID(Stk)
                                StkFishRateScalerGrid.Item(TStep + 3, rowcount).Value = StockFishRateScalers(Stk, fish, TStep).ToString("###0.0000")
                            End If

                        Next TStep
                        rowcount = rowcount + 1

                    End If
                Next
            Next
        End If

        GridLoading = False
    End Sub
End Class