Public Class FVS_StockSelect

   Private Sub FVS_StockSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 874
      FormWidth = 590
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_StockSelect_ReSize = False Then
            Resize_Form(Me)
            FVS_StockSelect_ReSize = True
         End If
      End If

      StockListBox.Items.Clear()
      StockGroupLabel.Text = ""
      For Stk As Integer = 1 To NumStk
         StockListBox.Items.Add(StockTitle(Stk))
      Next
      StockGroupNameTextBox.Text = ""
      If StockSelectionType = 1 Then
         StockSelectionTitle.Text = "Single Stock Selection"
         StockGroupNameTextBox.Visible = False
         StockGroupLabel.Visible = False
      ElseIf StockSelectionType = 2 Then
         StockSelectionTitle.Text = "Single/Multi Stock Selection"
         StockGroupNameTextBox.Visible = True
         StockGroupLabel.Visible = True
      ElseIf StockSelectionType = 3 Then
         StockSelectionTitle.Text = "Stock Group Selection"
         StockGroupNameTextBox.Visible = True
         StockGroupLabel.Visible = True
      End If
   End Sub

   'Private Sub StockListBox_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles StockListBox.Click
   '   StockSelection = StockListBox.SelectedIndex + 1
   '   Me.Visible = False
   '   FVS_ScreenReports.Enabled = True
   'End Sub

   Private Sub SSDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SSDoneButton.Click
      NumSelectedStocks = StockListBox.SelectedItems.Count
      ReDim StockSelection(NumSelectedStocks)
      For Stk As Integer = 1 To NumSelectedStocks
         StockSelection(Stk) = StockListBox.SelectedIndices.Item(Stk - 1) + 1
      Next
      If NumSelectedStocks > 1 And StockSelectionType > 1 Then
         If StockGroupNameTextBox.Text <> "" Then
            StockGroupName = StockGroupNameTextBox.Text
         Else
            MsgBox("Stock Group Name Required for this Selection!!" & vbCrLf & _
            "Please Enter Name in TextBox Below", MsgBoxStyle.OkOnly)
            Exit Sub
         End If
      ElseIf NumSelectedStocks = 1 And StockSelectionType > 1 Then
         StockGroupName = StockTitle(StockSelection(1))
      End If
      Me.Close()

      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FVS_StockFisheryScalerEdit.Visible = True
      End If

   End Sub

   Private Sub SSCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SSCancelButton.Click
      NumSelectedStocks = 0
      Me.Close()
      '- Return Point for calling SubRoutines 
      If CallingRoutine = 1 Then
         FVS_ReportSelection.Visible = True
      ElseIf CallingRoutine = 2 Then
         FVS_ScreenReports.Visible = True
      ElseIf CallingRoutine = 3 Then
         FVS_StockFisheryScalerEdit.Visible = True
      End If
   End Sub
End Class