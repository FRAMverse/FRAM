Imports System.Data.OleDb
Imports System.Data
Imports System.IO
Imports System.Windows
Imports System.Text
Imports System.IO.File

Public Class FVS_MainMenu

   Private Sub FVS_MainMenu_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      ReadOldCmd = False
      FramVersionLabel.Text = "Version " & FramVersion
      FormHeight = 762
      FormWidth = 832
      If DevWidth > My.Computer.Screen.Bounds.Width Then
         FormWidthScaler = DevWidth / My.Computer.Screen.Bounds.Width
      Else
         FormWidthScaler = 1
      End If
      'FormWidthScaler = 1280 / My.Computer.Screen.Bounds.Width
      '- Check if Form fits within Screen Dimensions
      If (FormHeight > My.Computer.Screen.Bounds.Height Or _
          FormWidth > My.Computer.Screen.Bounds.Width) Then
         Me.Height = FormHeight / (DevHeight / My.Computer.Screen.Bounds.Height)
         Me.Width = FormWidth / (DevWidth / My.Computer.Screen.Bounds.Width)
         If FVS_MainMenu_ReSize = False Then
            Resize_Form(Me)
            FVS_MainMenu_ReSize = True
         End If
      End If
      BackFramSave = False
   End Sub

   Private Sub FVS_Exit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FVS_Exit.Click
      Dim Result As Integer
      If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
         ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
         ChangeStockFishScaler = True Or ChangeStockRecruit = True Then
         ChangeAnyInput = True
         Result = MsgBox("Input Values have been Changed!" & vbCrLf & "Save Current Model Run ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
                Call SaveModelRunInputs()
            
            End If
      End If

        End

   End Sub

   Private Sub OpenDB_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OpenDB.Click

      Dim OpenFVSdatabase As New OpenFileDialog()
      'Dim FVSdatabaseBackupName As String
      'Dim DBNameLen As Integer
      Dim Result As Integer

      If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
         ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
         ChangeStockFishScaler = True Or ChangeStockRecruit = True Then
         ChangeAnyInput = True
         Result = MsgBox("Input Values have been Changed!" & vbCrLf & "Save Current Model Run ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            Call SaveModelRunInputs()
         End If
      End If

TryDBAgain:
      FVSdatabasename = ""
      OpenFVSdatabase.Filter = "DataBase Files (*.mdb)|*.mdb|All files (*.*)|*.*"
      OpenFVSdatabase.FilterIndex = 1
      OpenFVSdatabase.RestoreDirectory = True
      If OpenFVSdatabase.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            FVSdatabasename = OpenFVSdatabase.FileName
            FVSshortname = My.Computer.FileSystem.GetFileInfo(FVSdatabasename).Name
            FVSdatabasepath = My.Computer.FileSystem.GetFileInfo(FVSdatabasename).DirectoryName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If InStr(FVSdatabasename, "NewTransferModelRun.mdb") > 0 Or _
         InStr(FVSdatabasename, "ModelRunTransfer.mdb") > 0 Then
         MsgBox("You cannot use the NewTransferModelRun or ModelRunTransfer" & vbCrLf & _
         "Databases because they are reserved for Run Transfer Operations" & vbCrLf & _
         "Please chose another Database", MsgBoxStyle.OkOnly)
         GoTo TryDBAgain
      End If
      If FVSdatabasename.Length > 50 Then
         DatabaseNameLabel.Text = FVSshortname
      Else
         DatabaseNameLabel.Text = FVSdatabasename
      End If
      If FVSdatabasename = "" Then Exit Sub

      ''- Auto Save Backup-Copy of Database File
      'Me.Cursor = Cursors.WaitCursor
      'DBNameLen = InStr(FVSdatabasename, "mdb")
      'FVSdatabaseBackupName = FVSdatabasename.Substring(0, DBNameLen - 2) & "_AutoBackup.mdb"
      'If Exists(FVSdatabaseBackupName) Then Delete(FVSdatabaseBackupName)
      'File.Copy(FVSdatabasename, FVSdatabaseBackupName, True)
      'Me.Cursor = Cursors.Default

      '- DB Connection String
      FramDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FVSdatabasename
      Me.Visible = False



      '==============================================================================================
      '- (Pete 12/13) Code that checks for the existence of the Target Sublegal:Legal Ratio and 
      '  and RunEncounterRateAdjustment table (SLRatio) 
      '- needed to use external sublegals; works to make things functional retroactively
      Dim sql As String       'SQL Query text string
      Dim oledbAdapter As OleDb.OleDbDataAdapter

      'First check the FRAM database for the SLRatio and RunEncounterRateAdjustment tables
      FramDB.Open()
      Dim restrictions1(3) As String
      Dim restrictions2(3) As String
      Dim DoesTableExist1 As Boolean
      Dim DoesTableExist2 As Boolean
      restrictions1(2) = "SLRatio"

      Dim dbTbl As DataTable = FramDB.GetSchema("Tables", restrictions1)
      If dbTbl.Rows.Count = 0 Then
         'Table does not exist
         DoesTableExist1 = False
      Else
         'Table exists
         DoesTableExist1 = True
      End If

      dbTbl.Dispose()
      FramDB.Close()

      'If SLRatio doesn't exist, create it.
      If DoesTableExist1 = False Then
         sql = "CREATE TABLE SLRatio (RunID INTEGER,FisheryID INTEGER,Age INTEGER,TimeStep INTEGER,TargetRatio DOUBLE, RunEncounterRateAdjustment DOUBLE, UpdateWhen DATETIME, UpdateBy VARCHAR(255))"
         'Now connect to the database and make the table...
         'create a command
         Dim my_Command As New OleDbCommand(sql, FramDB)
         FramDB.Open()
         'command execute
         my_Command.ExecuteNonQuery()
         FramDB.Close()
      End If

      '==============================================================================================

      '- Recordset (Model Run) Selection
      RecordsetSelectionType = 1
      FVS_ModelRunSelection.ShowDialog()
      Me.Visible = True
      Me.BringToFront()
      RecordSetNameLabel.Text = RunIDNameSelect
      If RunIDSelect = 0 Then
         FVSdatabasename = ""
         'MsgBox("NO Recordsets Available in this Database File" & vbCrLf & _
         '"You must Read Old CMD File to continue", MsgBoxStyle.OkOnly)
      End If

   End Sub

   Private Sub InputOptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles InputOptions.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False
      FVS_InputMenu.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub ModelRun_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ModelRun.Click
      Dim Result As Integer
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
          ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
          ChangeStockFishScaler = True Or ChangeStockRecruit = True Then
         ChangeAnyInput = True
         Result = MsgBox("Input Values have been Changed!" & vbCrLf & "Changes Must be Saved before Running Model!!!" & vbCrLf & "Save Current Model Run ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            'Call SaveModelRunInputs()
            Me.Visible = False
            FVS_SaveModelRunInputs.ShowDialog()
            Me.Visible = True
            RecordSetNameLabel.Text = RunIDNameSelect
            Me.BringToFront()
         Else
            MsgBox("Please be aware that the OUTPUT for this run" & vbCrLf & "cannot be duplicated without saving your INPUT values", MsgBoxStyle.OkOnly)
         End If
      End If
      Me.Visible = False
      FVS_RunModel.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub FramUtilButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FramUtilButton.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False

      FVS_FramUtils.ShowDialog()
      RecordSetNameLabel.Text = RunIDNameSelect
      Me.BringToFront()
   End Sub

   Private Sub OutputResults_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles OutputResults.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False
      FVS_Output.ShowDialog()
      Me.Refresh()
      Me.BringToFront()
   End Sub

   Private Sub PostSeason_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PostSeason.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      Me.Visible = False
      FVS_BackwardsFram.ShowDialog()
      Me.BringToFront()
   End Sub

   Private Sub SelectRecordset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SelectRecordset.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      RecordsetSelectionType = 1
      Me.Visible = False
      FVS_ModelRunSelection.ShowDialog()
      Me.Visible = True
      Me.BringToFront()
      RecordSetNameLabel.Text = RunIDNameSelect
      If RunIDSelect = 0 Then
         FVSdatabasename = ""
      End If
   End Sub

   Private Sub SaveInputButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles SaveInputButton.Click
      If FVSdatabasename = "" Then
         MsgBox("Database and Model Run Must be Selected First !!!", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
        If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
           ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
           ChangeStockFishScaler = True Or ChangeStockRecruit = True Or AnyChange = True Then
            ChangeAnyInput = True
        Else
            MsgBox("No Input Values have been Changed!" & vbCrLf & "No Action Taken", MsgBoxStyle.OkOnly)
            Exit Sub
        End If

      Me.Visible = False
      FVS_SaveModelRunInputs.ShowDialog()
      Me.Visible = True
      RecordSetNameLabel.Text = RunIDNameSelect
      Me.BringToFront()

   End Sub

   Private Sub VersionChangesButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles VersionChangesButton.Click
      Me.Visible = False
      FVS_VersionChanges.ShowDialog()
      Me.Visible = True
      Me.BringToFront()
   End Sub

End Class