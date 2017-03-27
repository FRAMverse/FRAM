Public Class FVS_OutputDriverSelection

   Private Sub FVS_OutputDriverSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 766
      FormWidth = 553
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
         If FVS_OutputDriverSelection_ReSize = False Then
            Resize_Form(Me)
            FVS_OutputDriverSelection_ReSize = True
         End If
      End If

      FramDB.Open()
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT DISTINCT DriverName FROM ReportDriver WHERE SpeciesName = " & Chr(34) & SpeciesName & Chr(34)
      drd1 = cmd1.ExecuteReader
      Dim str1 As String
      CheckedListBox1.Items.Clear()
      If drd1.HasRows = False Then
         '- No Report Driver in Table .. Must Read Old DRV File
         MsgBox("NO Report DRIVERS Exist in Your Database Table" & vbCrLf & "You Must READ or CREATE Report DRIVER First", MsgBoxStyle.OkOnly)
         FVS_OutputDriver.ReportDriverSelectionLabel.Text = "No Report Drivers Available"
         ReportDriverName = ""
         Me.Close()
         FVS_OutputDriver.Visible = True
      End If
      Do While drd1.Read
         '- Fill CheckedListBox Items
         str1 = drd1.GetString(0).ToString
         CheckedListBox1.Items.Add(str1)
         '- Set RunID Array Values
      Loop
      FramDB.Close()
      If DriverSelectionType = 1 Then
         DriverSelectionLabel.Text = "Report Driver Selection"
      ElseIf DriverSelectionType = 2 Then
         DriverSelectionLabel.Text = "Delete Report Driver"
      End If
      Me.BringToFront()

   End Sub

   Private Sub CheckedListBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.Click
      Dim result As Integer
      ReportDriverName = CheckedListBox1.SelectedItem.ToString
      '- Delete Report Driver for this Selection Type
      If DriverSelectionType = 2 Then
         result = MsgBox("DELETE this Report Driver ?? = " & ReportDriverName, MsgBoxStyle.YesNo)
         If result = vbYes Then
            '- Report Driver SELECT Statement
            Dim CmdStr As String
            CmdStr = "SELECT * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName & Chr(34) & ";"
            Dim RDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            Dim RDDA As New System.Data.OleDb.OleDbDataAdapter
            RDDA.SelectCommand = RDcm
            '- RunID DELETE Statement
            CmdStr = "DELETE * FROM ReportDriver WHERE DriverName = " & Chr(34) & ReportDriverName & Chr(34) & ";"
            Dim RDDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
            RDDA.DeleteCommand = RDDcm
            '- Command Builder
            Dim RDcb As New OleDb.OleDbCommandBuilder
            RDcb = New OleDb.OleDbCommandBuilder(RDDA)
            FramDB.Open()
            RDDA.DeleteCommand.ExecuteScalar()
            FramDB.Close()
            ReportDriverName = ""
         Else
            ReportDriverName = ""
         End If
      End If
      FVS_OutputDriver.ReportDriverSelectionLabel.Text = ReportDriverName
      Me.Close()
      FVS_OutputDriver.Visible = True
   End Sub

   Private Sub RDSCancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RDSCancelButton.Click
      FVS_OutputDriver.ReportDriverSelectionLabel.Text = "No Report Driver Selected"
      ReportDriverName = ""
      Me.Close()
      FVS_OutputDriver.Visible = True
   End Sub

End Class