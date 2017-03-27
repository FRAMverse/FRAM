Public Class FVS_VersionChanges

   Private Sub FVS_Done_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FVS_Done.Click
      Me.Close()
      FVS_MainMenu.Visible = True
      Exit Sub
   End Sub

   Private Sub FVS_VersionChanges_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      Dim LastChange As Integer
      VersionChangeListBox.Items.Clear()
      For Stk As Integer = 0 To 100
         If VersionNumberChanges(Stk) = "" Then
            LastChange = Stk - 1
            Exit For
         End If
      Next
      For Stk As Integer = LastChange To 0 Step -1
         VersionChangeListBox.Items.Add(VersionNumberChanges(Stk))
      Next
   End Sub

   Private Sub VersionChangeListBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles VersionChangeListBox.SelectedIndexChanged

   End Sub
End Class