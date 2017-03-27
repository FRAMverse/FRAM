Imports System.Data.OleDb
Public Class FVS_BasePeriodSelect

   Private Sub FVS_BasePeriodSelect_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 827
      FormWidth = 1014
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
         If FVS_BasePeriodSelect_ReSize = False Then
            Resize_Form(Me)
            FVS_BasePeriodSelect_ReSize = True
         End If
      End If

      FramDB.Open()
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      cmd1.Connection = FramDB
      If ModelRunBPSelect = True Then
         cmd1.CommandText = "SELECT * FROM BaseID WHERE SpeciesName = " & Chr(34) & SelectSpeciesName.ToString & Chr(34) & " ORDER BY BasePeriodID ASC"
      Else
         cmd1.CommandText = "SELECT * FROM BaseID ORDER BY BasePeriodID ASC"
      End If
      drd1 = cmd1.ExecuteReader
      Dim str1 As String
      Dim int1 As Integer
      int1 = 0
      CheckedListBox1.Items.Clear()
      Do While drd1.Read
         '- Fill CheckedListBox Items
         str1 = String.Format("{0,4}-", drd1.GetInt32(1).ToString("###0"))
         str1 &= String.Format("{0,-7}-", drd1.GetString(3).ToString)
         str1 &= String.Format("{0,-20}-", Mid(drd1.GetString(2).ToString, 1, 20))
         str1 &= String.Format(" Stks={0,4}", drd1.GetInt32(4).ToString)
         str1 &= String.Format(" Fish={0,4}", drd1.GetInt32(5).ToString)
         str1 &= String.Format(" TStp={0,1}", drd1.GetInt32(6).ToString)
         str1 &= String.Format(" Date={0,10}", drd1.GetDateTime(10).ToString)
         CheckedListBox1.Items.Add(str1)
      Loop
      FramDB.Close()

   End Sub

   Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.Click
      BasePeriodIDSelect = CInt(Mid(CheckedListBox1.Items(CheckedListBox1.SelectedIndex), 1, 4))
      'return point check
      If ModelRunBPSelect = True Then
         Me.Close()
         FVS_ModelRunSelection.Visible = True
      Else
         If BasePeriodIDSelect = BasePeriodID Then
            MsgBox("ERROR- Can't DELETE BasePeriod when it is CURRENTLY in use!!" & vbCrLf & "Try SELECTING another CMD RecordSet" & vbCrLf & "that doesn't use the SELECTED BASE" & vbCrLf & "and then DELETE the Selection!", MsgBoxStyle.OkOnly)
            BasePeriodIDSelect = 0
         End If
         Me.Close()
         FVS_FramUtils.Visible = True
      End If
   End Sub

   Private Sub BPCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BPCancelButton.Click
      Me.Close()
      BasePeriodIDSelect = 0
      'return point check
      If ModelRunBPSelect = True Then
         FVS_ModelRunSelection.Visible = True
      Else
         FVS_FramUtils.Visible = True
      End If
   End Sub

End Class