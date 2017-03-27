Public Class FVS_InputMenu

   Private Sub CmdStockRecruits_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdStockRecruits.Click
      Me.Visible = False
      FVS_StockRecruitEdit.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub CmdExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdExit.Click
      Me.Close()
      FVS_MainMenu.Visible = True
   End Sub

   Private Sub CmdFishery_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdFishery.Click
      Me.Visible = False
      FVS_FisheryScalerEdit.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub CmdNonRetention_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdNonRetention.Click
      Me.Visible = False
      FVS_NonRetentionEdit.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub CmdStkFish_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdStkFish.Click
      Me.Visible = False
      FVS_StockFisheryScalerEdit.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub FVS_InputMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      FormHeight = 706
      FormWidth = 850
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
         If FVS_InputMenu_ReSize = False Then
            Resize_Form(Me)
            FVS_InputMenu_ReSize = True
         End If
      End If
      Me.BringToFront()
   End Sub

   Private Sub CmdSizeLimits_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdSizeLimits.Click
      If SpeciesName = "COHO" Then
         MsgBox("No Size-Limit Evaluation for COHO", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False
      FVS_SizeLimitEdit.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub CmdPSCMaxER_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmdPSCMaxER.Click
      If SpeciesName = "CHINOOK" Then
         MsgBox("No PSC-ER Evaluation for CHINOOK", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False
      FVS_PSCMaxER.ShowDialog()
      Me.BringToFront()
   End Sub
End Class