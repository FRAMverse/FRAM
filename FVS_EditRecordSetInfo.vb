Imports System.Data.OleDb

Public Class FVS_EditRecordSetInfo

   Private Sub FVS_EditRecordsetInfo_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 720
      FormWidth = 892
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_EditRecordSetInfo_ReSize = False Then
            Resize_Form(Me)
            FVS_EditRecordSetInfo_ReSize = True
         End If
      End If

      If RecordsetSelectionType = 3 Then
         '- EDIT RunID Header Information
         RSETitle.Text = "Recordset Information"
         RunIDLabel.Text = RunIDSelect.ToString
         SpeciesNameLabel.Text = SpeciesName.ToString
         BasePeriodIDLabel.Text = BasePeriodID.ToString
         BasePeriodNameLabel.Text = BasePeriodName.ToString
            RunNameTextBox.Text = RunIDNameSelect.ToString
         RunTitleTextBox.Text = RunIDTitleSelect.ToString
         CommentsRichTextBox.Text = RunIDCommentsSelect
         CreationDateLabel.Text = RunIDCreationDateSelect.ToString
         ModifyInputDateLabel.Text = RunIDModifyInputDateSelect.ToString
            RunTimeDateLabel.Text = RunIDRunTimeDateSelect.ToString
            RunYearTextBox.Text = RunIDYearSelect.ToString
      ElseIf RecordsetSelectionType = 4 Or RecordsetSelectionType = 5 Then
         '- EDIT New RunID Header Information (from Copied Recordset)
         Dim drd1 As OleDb.OleDbDataReader
         Dim cmd1 As New OleDb.OleDbCommand()
         Dim MaxOldID As Integer
         '- Get Current Max RunID Value, Add One for New Recordset RunID Value
         cmd1.Connection = FramDB
         cmd1.CommandText = "SELECT * FROM RunID ORDER BY RunID DESC"
         FramDB.Open()
         drd1 = cmd1.ExecuteReader
         drd1.Read()
         MaxOldID = drd1.GetInt32(1)
         cmd1.Dispose()
         drd1.Dispose()
         FramDB.Close()
         '- Value used in CopyNewRecordset Routine
         NewRunID = MaxOldID + 1
         RSETitle.Text = "NEW Copied Recordset Information"
         RunIDLabel.Text = NewRunID.ToString
         SpeciesNameLabel.Text = SpeciesName.ToString
         BasePeriodIDLabel.Text = BasePeriodID.ToString
         BasePeriodNameLabel.Text = BasePeriodName.ToString
         RunNameTextBox.Text = "COPY OF " & RunIDNameSelect.ToString
            RunTitleTextBox.Text = RunIDTitleSelect.ToString
            RunYearTextBox.Text = RunIDYearSelect.ToString
         CommentsRichTextBox.Text = RunIDCommentsSelect
         CreationDateLabel.Text = Now.ToString
         ModifyInputDateLabel.Text = ""
         RunTimeDateLabel.Text = ""
      End If

   End Sub

   Private Sub RECancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RECancelButton.Click
      If RecordsetSelectionType <= 4 Then
         '- Cancel Recordset Copy
         RecordsetSelectionType = -4
         Me.Close()
         FVS_FramUtils.Visible = True
      End If
      If RecordsetSelectionType = 5 Then
         '- Cancel Recordset Copy
         RecordsetSelectionType = -5
         Me.Close()
         FVS_SaveModelRunInputs.Visible = True
      End If
   End Sub

   Private Sub REDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles REDoneButton.Click

        'Dim AnyChange As Boolean
      Dim Result As Integer
      AnyChange = False

      If RunNameTextBox.Text <> RunIDNameSelect.ToString Then
         AnyChange = True
        End If
        
        If (RunTitleTextBox.Text <> RunIDTitleSelect.ToString) Then
            AnyChange = True
        End If
        If (RunTitleTextBox.Text <> RunIDTitleSelect.ToString) Then
            AnyChange = True
        End If
        If CommentsRichTextBox.Text <> RunIDCommentsSelect Then
            AnyChange = True
        End If
        If RunYearTextBox.Text <> RunIDYearSelect.ToString Then
            AnyChange = True
        End If

        If RecordsetSelectionType = 4 Or RecordsetSelectionType = 5 Then
            If AnyChange = False Then
                Result = MsgBox("RunName, Title, or Comments should be changed for NEW Recordset" & vbCrLf & _
                "Do You Want to Re-Edit these for this NEW Recordset ???", MsgBoxStyle.YesNo)
                If Result = vbYes Then Exit Sub
            End If
            ''Result = MsgBox("Please Select OPEN DATABASE after Recordset COPY" & vbCrLf & _
            ''   "to SELECT the CURRENT Recordset for use", MsgBoxStyle.OkOnly)
            ''RunIDSelect = 0
            ''FVSdatabasename = ""
            'Me.Close()
            'If RecordsetSelectionType = 4 Then
            '   Me.Close()
            '   FVS_MainMenu.Visible = True
            '   Exit Sub
            'ElseIf RecordsetSelectionType = 5 Then
            '   Me.Close()
            '   FVS_SaveModelRunInputs.Visible = True
            '   Exit Sub
            '   'RunIDSelect = NewRunID
            '   'Me.Visible = False
            '   'FVS_FramUtils.Visible = True
            '   'Exit Sub
            'End If
            RunIDSelect = NewRunID
        End If

        RunIDNameSelect = RunNameTextBox.Text
        RunIDTitleSelect = RunTitleTextBox.Text
        RunIDCommentsSelect = CommentsRichTextBox.Text
        RunIDYearSelect = RunYearTextBox.Text
        Dim CmdStr As String
        '- DataApapter SELECT Statement
        CmdStr = "SELECT * FROM RunID WHERE RunID = " & RunIDSelect.ToString & " ORDER BY StockID, Age, TimeStep"
        Dim RIcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        Dim RunDA As New System.Data.OleDb.OleDbDataAdapter
        RunDA.SelectCommand = RIcm
        '- DataApapter DELETE Statement
        CmdStr = "DELETE * FROM RunID WHERE RunID = " & RunIDSelect.ToString & ";"
        Dim RIDcm As New OleDb.OleDbCommand(CmdStr, FramDB)
        RunDA.DeleteCommand = RIDcm
        '- Command Builder
        Dim RIcb As New OleDb.OleDbCommandBuilder
        RIcb = New OleDb.OleDbCommandBuilder(RunDA)
        '- Set Up DataBase Transaction
        Dim RITrans As OleDb.OleDbTransaction
        Dim RIC As New OleDbCommand
        FramDB.Open()
        RunDA.DeleteCommand.ExecuteScalar()
        RITrans = FramDB.BeginTransaction
        RIC.Connection = FramDB
        RIC.Transaction = RITrans
        '- INSERT Record into DataBase Table
        RIC.CommandText = "INSERT INTO RunID (RunID,SpeciesName,RunName,RunTitle,BasePeriodID,RunComments,CreationDate,ModifyInputDate,RunTimeDate,RunYear) " & _
              "VALUES(" & RunIDSelect.ToString & "," & _
              Chr(34) & SpeciesName.ToString & Chr(34) & "," & _
            Chr(34) & RunIDNameSelect.ToString & Chr(34) & "," & _
            Chr(34) & RunIDTitleSelect.ToString & Chr(34) & "," & _
            BasePeriodID.ToString & "," & _
            Chr(34) & RunIDCommentsSelect.ToString & Chr(34) & "," & _
            Chr(35) & RunIDCreationDateSelect.ToString & Chr(35) & "," & _
            Chr(35) & Now().ToString & Chr(35) & "," & _
            Chr(35) & RunIDRunTimeDateSelect.ToString & Chr(35) & "," & _
            Chr(34) & RunIDYearSelect & Chr(34) & ");"

        'Chr(35) & "1/1/1" & Chr(35) & "," & _
        'Chr(35) & "1/1/1" & Chr(35) & "," & _
        'Chr(35) & RunIDCreationDateSelect.ToString & Chr(35) & "," & _
        'Chr(35) & RunIDModifyInputDateSelect.ToString & Chr(35) & "," & _
        'Chr(35) & RunIDRunTimeDateSelect.ToString & Chr(35) & ");"
        RIC.ExecuteNonQuery()
        RITrans.Commit()
        FramDB.Close()
        RunDA = Nothing

        'If RecordsetSelectionType = 4 Then
        '   Result = MsgBox("Please Select OPEN DATABASE after Recordset COPY" & vbCrLf & _
        '      "to SELECT the CURRENT Recordset for use", MsgBoxStyle.OkOnly)
        '   RunIDSelect = 0
        '   FVSdatabasename = ""
        '   Me.Visible = False
        '   FVS_MainMenu.Visible = True
        '   Exit Sub
        'End If

        Me.Close()

        If RecordsetSelectionType = 5 Then
            FVS_SaveModelRunInputs.Visible = True
        Else
            FVS_FramUtils.Visible = True
        End If

    End Sub

    Private Sub RunNameTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunNameTextBox.TextChanged

    End Sub

    Private Sub RunTitleTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunTitleTextBox.TextChanged

    End Sub

    Private Sub CommentsRichTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CommentsRichTextBox.TextChanged

    End Sub

    Private Sub RunYearTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RunYearTextBox.TextChanged

    End Sub
End Class