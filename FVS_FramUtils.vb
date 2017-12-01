Imports System.IO
Imports System.IO.File
Public Class FVS_FramUtils

   Private Sub FVS_FramUtils_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      FormHeight = 778
      FormWidth = 892
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
         If FVS_FramUtils_ReSize = False Then
            Resize_Form(Me)
            FVS_FramUtils_ReSize = True
         End If
        End If
        'If SpeciesName = "COHO" Then
        '    GetBPTransferBtn.Visible = False
        '    TransferBPBtn.Visible = False
        'ElseIf SpeciesName = "CHINOOK" Then
        GetBPTransferBtn.Visible = True
        TransferBPBtn.Visible = True
        'End If
   End Sub

   Private Sub FUExitButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles FUExitButton.Click
      Me.Visible = False
      FVS_MainMenu.Visible = True
   End Sub

   Private Sub ReadCmdButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadCmdButton.Click

      Dim Result As Integer
      '- First Get CMD File Name
      OldCMDFile = ""
      CMDFileDialog.Filter = "Command Files (*.CMD)|*.CMD|All files (*.*)|*.*"
      CMDFileDialog.FilterIndex = 1
      CMDFileDialog.RestoreDirectory = True
      If CMDFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            OldCMDFile = CMDFileDialog.FileName
            OldCMDFilePath = My.Computer.FileSystem.GetFileInfo(OldCMDFile).DirectoryName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If OldCMDFile = "" Then Exit Sub
      If Exists(OldCMDFile) Then
         Jim = 1
         ReadOldCmd = True
      End If


      If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
          ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
          ChangeStockFishScaler = True Or ChangeStockRecruit = True Then
         ChangeAnyInput = True
         Result = MsgBox("Input Values have been Changed!" & vbCrLf & "Save Current Model Run ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            'Call SaveModelRunInputs()
            Me.Visible = False
            FVS_SaveModelRunInputs.ShowDialog()
            Me.Visible = True
            RecordSetNameLabel.Text = RunIDNameSelect
            Me.BringToFront()
         End If
      End If

      Me.Enabled = False
      '- Call FramUtils Module Routine
      Me.Cursor = Cursors.WaitCursor
      ReadOldCommandFile()
      Me.Cursor = Cursors.Default

      Me.Enabled = True
      Me.BringToFront()

      ReadOldCmd = False

   End Sub

   Private Sub DelRecSetButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DelRecSetButton.Click

      Me.Visible = False
      RecordsetSelectionType = 2
      FVS_ModelRunSelection.btn_DeleteMulti.Visible = True
      FVS_ModelRunSelection.ShowDialog()
      If RecordsetSelectionType = 9 Then
         RecordsetSelectionType = 0
         Exit Sub
      End If
      Cursor.Current = Cursors.WaitCursor
      Call DeleteRecordset()
      Me.Cursor = Cursors.Default
      Me.BringToFront()

   End Sub

   Private Sub ReadOUTFileButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadOUTFileButton.Click

      '- First Get OUT File Name
      OldOUTFile = ""
      CMDFileDialog.Filter = "BasePeriod Files (*.OUT)|*.OUT|All files (*.*)|*.*"
      CMDFileDialog.FilterIndex = 1
      CMDFileDialog.RestoreDirectory = True
      If CMDFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            OldOUTFile = CMDFileDialog.FileName
            OldOUTFilePath = My.Computer.FileSystem.GetFileInfo(OldOUTFile).DirectoryName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If OldOUTFile = "" Then Exit Sub
      If Exists(OldOUTFile) Then Jim = 1
      Me.Enabled = False
      '- Call FramUtils Module Routine
      Me.Cursor = Cursors.WaitCursor
      ReadOldBasePeriodOUTFile()
      Me.Cursor = Cursors.Default

      Me.Enabled = True
      Me.BringToFront()

   End Sub

   Private Sub RecSetInfoButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles RecSetInfoButton.Click
      Me.Visible = False
      RecordsetSelectionType = 3
      FVS_EditRecordSetInfo.ShowDialog()
      RecordsetSelectionType = 0
      Me.BringToFront()
   End Sub

   Private Sub DeleteBPButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles DeleteBPButton.Click
      Me.Visible = False
      FVS_BasePeriodSelect.ShowDialog()
      Me.BringToFront()
      If BasePeriodIDSelect = 0 Then Exit Sub
      DeleteBasePeriodRecordset()
   End Sub

   Private Sub CopyRecordsetButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CopyRecordsetButton.Click
      Dim Result As Integer
      Me.Visible = False
      RecordsetSelectionType = 4
      FVS_EditRecordSetInfo.ShowDialog()
      If RecordsetSelectionType = -4 Then
         MsgBox("Recordset COPY Cancelled", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      If ChangeAnyInput = True Or ChangeBackFram = True Or ChangeFishScalers = True Or _
          ChangeNonRetention = True Or ChangePSCMaxER = True Or ChangeSizeLimit = True Or _
          ChangeStockFishScaler = True Or ChangeStockRecruit = True Then
         ChangeAnyInput = True
         Result = MsgBox("Input Values have been Changed!" & vbCrLf & "Save Current Model Run ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
            'Call SaveModelRunInputs()
            Me.Visible = False
            FVS_SaveModelRunInputs.ShowDialog()
            Me.Visible = True
            RecordSetNameLabel.Text = RunIDNameSelect
            Me.BringToFront()
         End If
      End If
      Me.Cursor = Cursors.WaitCursor
      Call CopyNewRecordset()
      Me.Cursor = Cursors.Default
      RecordsetSelectionType = 0
      Me.BringToFront()
   End Sub

   Private Sub ReadTaaEtrsButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadTaaEtrsButton.Click
      Dim result
      result = MsgBox("The Old Version of the 'TaaETRSList' will be Deleted" & vbCrLf & "Do you want to replace it with the new 'TaaETRSNum.Txt' file ???", MsgBoxStyle.YesNo)
      If result = vbNo Then Exit Sub
      Call ReadTaaEtrsFile()
   End Sub

   Private Sub TransferModelRunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TransferModelRunButton.Click
        Dim TransferDBName, NewTransferDB, TransferDBNameShort As String

        TransferDBName = ""
        TransferDBNameShort = ""

        MsgBox("Please select the Transfer Database.")
        OpenTransferModelRunFileDialog.Filter = "Model Run Transfer Files (*.MDB)|*.MDB|All files (*.*)|*.*"
        OpenTransferModelRunFileDialog.FilterIndex = 1
        OpenTransferModelRunFileDialog.RestoreDirectory = True
        If OpenTransferModelRunFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                TransferDBName = OpenTransferModelRunFileDialog.FileName
                TransferDBNameShort = System.IO.Path.GetFileName(OpenTransferModelRunFileDialog.FileName)
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
        End If



        RecordsetSelectionType = 3
        If Exists(TransferDBName) Then
            Me.Visible = False
            FVS_ModelRunSelection.ShowDialog()
            If RecordsetSelectionType = 9 Then
                MsgBox("Model Run Transfer Cancelled", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            Me.Refresh()
            Me.Cursor = Cursors.WaitCursor
            '- Create Copy of Transfer Database File


            MDBSaveFileDialog.Filter = "*.mdb|*.mdb"

NewName:
            NewTransferDB = ""
            MDBSaveFileDialog.Filter = "Transfer File Name (*.MDB)|*.MDB|All files (*.*)|*.*"
            MDBSaveFileDialog.FilterIndex = 1
            MDBSaveFileDialog.FileName = TransferDBNameShort
            MDBSaveFileDialog.RestoreDirectory = True
            If MDBSaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    NewTransferDB = MDBSaveFileDialog.FileName
                Catch Ex As Exception
                    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                End Try
            End If
            If NewTransferDB = "" Then Exit Sub
            If NewTransferDB = "NewModelRunTransfer5.Mdb" Then
                MsgBox("The file 'NewModelRunTransfer5.Mdb' is Reserved" & vbCrLf & _
                       "Please Choose Different Name for Transfer DataBase" & vbCrLf & _
                       "Prevents Corruption of Database Structure!", MsgBoxStyle.OkOnly)
                GoTo NewName
            End If

            'If Exists(FVSdatabasepath & "\" & NewTransferDB) Then Delete(FVSdatabasepath & "\" & NewTransferDB)
            'File.Copy(FVSdatabasepath & "\" & TransferDBName, NewTransferDB, True)
            If Exists(NewTransferDB) Then Delete(NewTransferDB)
            File.Copy(TransferDBName, NewTransferDB, True)



            '- TransferDB Connection String
            TransDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NewTransferDB

            '==============================================================================================
            '- (Pete 12/13) Part I.  Code that corrects an error in older versions of the Transfer Database
            '- (Basically an older version included the MSF flag in the FisheryScaler table;
            '- this presence of this field messed things up severely (alignment wise) when runs were imported/rerun

            ' Open connection to the database
            TransDB.Open()
            Dim dbTbl As New DataTable
            Dim DoesFieldExist As Boolean
            Dim tblName, fldName As String
            tblName = "FisheryScalers"
            fldName = "MarkSelectiveFlag"


            ' Get the table definition loaded in a table adapter
            Dim strSql As String = "Select TOP 1 * from " & tblName
            Dim dbAdapater As New System.Data.OleDb.OleDbDataAdapter(strSql, TransDB)
            dbAdapater.Fill(dbTbl)

            ' Get the index of the field name
            Dim i As Integer = dbTbl.Columns.IndexOf(fldName)

            If i = -1 Then
                'Field is missing
                DoesFieldExist = False
            Else
                'Field is there
                DoesFieldExist = True
            End If

            dbTbl.Dispose()
            TransDB.Close()

            If DoesFieldExist = True Then
                TransDB.Open()
                Dim transDBCommand As System.Data.OleDb.OleDbCommand
                transDBCommand = TransDB.CreateCommand()
                transDBCommand.CommandText = "ALTER TABLE " & tblName & " DROP COLUMN " & fldName
                transDBCommand.ExecuteNonQuery()
                TransDB.Close()
            End If
            '==============================================================================================

            '==============================================================================================
            '- (Pete 12/13) Part II.  Code that creates the Target Sublegal:Legal Ratio (SLRatio) 
            '- and run-specific sublegal encounter rate adjustment (RunEncounterRateAdjustment) tables
            '- needed to use external sublegals in the transfer database

            Dim sql As String       'SQL Query text string
            sql = "CREATE TABLE SLRatio (RunID INTEGER,FisheryID INTEGER,Age INTEGER,TimeStep INTEGER,TargetRatio DOUBLE,RunEncounterRateAdjustment DOUBLE, UpdateWhen DATETIME, UpdateBy VARCHAR(255))"
            'create a command
            Dim my_Command1 As New System.Data.OleDb.OleDbCommand(sql, TransDB)
            TransDB.Open()
            'command execute
            my_Command1.ExecuteNonQuery()
            TransDB.Close()

            '==============================================================================================



            Me.Cursor = Cursors.WaitCursor
            Call TransferModelRunTables()
            Me.Cursor = Cursors.Default
            Me.Visible = True
        Else
            MsgBox("Can't find NewModelRunTransfer.MDB file in Current Directory!" & vbCrLf & "Please Make Copy before Model Run Transfer", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Me.Cursor = Cursors.Default

    End Sub

   Private Sub GetModelRunButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles GetModelRunButton.Click
      Dim NewTransferDB As String

      NewTransferDB = ""
      CMDFileDialog.Filter = "Model Run Transfer Files (*.MDB)|*.MDB|All files (*.*)|*.*"
      CMDFileDialog.FilterIndex = 1
      CMDFileDialog.RestoreDirectory = True
      If CMDFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
         Try
            NewTransferDB = CMDFileDialog.FileName
         Catch Ex As Exception
            MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
         End Try
      End If
      If NewTransferDB = "" Then Exit Sub

      RecordsetSelectionType = 3
      If Exists(NewTransferDB) Then
         Me.Cursor = Cursors.WaitCursor
         '- TransferDB Connection String
         TransDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NewTransferDB
         Call GetTransferModelRunTables()
         Me.Visible = True
      End If
      Me.Cursor = Cursors.Default

   End Sub

   Private Sub CoweemanButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CoweemanButton.Click
      Me.Visible = False
      If NumStk < 67 Then
         MsgBox("ERROR- Coweeman Transfer is for current CHINOOK Base Period Runs" & vbCrLf & "You are using and older base period which do not have the new CR Stocks", MsgBoxStyle.OkOnly)
         Exit Sub
      End If
      FVS_Coweeman.ShowDialog()
      Me.BringToFront()
      Exit Sub
   End Sub

Private Sub btn_Chin2s3s_Click(sender As System.Object, e As System.EventArgs) Handles btn_Chin2s3s.Click

        'First load up the table with the goods
        'Dim dsNewTwos As New DataSet
        Dim tempRecruts(NumStk) As Double
        ChangeStockRecruit = True
        'Dim CWTdb As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FVSdatabasename)
        'Dim sql As String       'SQL Query text string
        'Dim oledbAdapter As OleDb.OleDbDataAdapter

        'sql = "SELECT * FROM ChinookTwoThreeMultipliers"

        'Try
        '    CWTdb.Open()
        '    oledbAdapter = New OleDb.OleDbDataAdapter(sql, CWTdb)
        '    oledbAdapter.Fill(dsNewTwos, "ChinookTwoThreeMultipliers")
        '    oledbAdapter.Dispose()
        '    CWTdb.Close()

        'Catch ex As Exception
        '    MsgBox("Failed to load Twos from Threes Multiplier table!" & vbCr & "Verify that your database contains this table and try again.")
        'End Try

        'Alternatively, define the values for adjustment in code to keep things simple on the database versioning end of things...


        'Dim NewTwos(,) As Object = {{1, "U-NkSm FF", "UnMarked Nooksack/Samish Fall", 0.961535},
        '                               {2, "M-NkSm FF", "Marked Nooksack/Samish Fall", 0.961456},
        '                               {3, "U-NFNK Sp", "UnMarked NF Nooksack Spr", 0.979531},
        '                               {4, "M-NFNK Sp", "Marked NF Nooksack Spr", 0.979533},
        '                               {5, "U-SFNK Sp", "UnMarked SF Nooksack Spr", 0.979539},
        '                               {6, "M-SFNK Sp", "Marked SF Nooksack Spr", 0.979539},
        '                               {7, "U-Skag FF", "UnMarked Skagit Summer/Fall Fing", 0.931295},
        '                               {8, "M-Skag FF", "Marked Skagit Summer/Fall Fing", 0.930408},
        '                               {9, "U-SkagFYr", "UnMarked Skagit Summer/Fall Year", 0.963309},
        '                               {10, "M-SkagFYr", "Marked Skagit Summer/Fall Year", 0.963309},
        '                               {11, "U-SkagSpY", "UnMarked Skagit Spring Year", 0.978404},
        '                               {12, "M-SkagSpY", "Marked Skagit Spring Year", 0.978334},
        '                               {13, "U-Snoh FF", "UnMarked Snohomish Fall Fing", 0.934515},
        '                               {14, "M-Snoh FF", "Marked Snohomish Fall Fing", 0.934488},
        '                               {15, "U-SnohFYr", "UnMarked Snohomish Fall Year", 0.948666},
        '                               {16, "M-SnohFYr", "Marked Snohomish Fall Year", 0.948623},
        '                               {17, "U-Stil FF", "UnMarked Stillaguamish Fall Fing", 0.932873},
        '                               {18, "M-Stil FF", "Marked Stillaguamish Fall Fing", 0.931865},
        '                               {19, "U-Tula FF", "UnMarked Tulalip Fall Fing", 0.97014},
        '                               {20, "M-Tula FF", "Marked Tulalip Fall Fing", 0.970262},
        '                               {21, "U-MidPSFF", "UnMarked Mid PS Fall Fing", 0.951612},
        '                               {22, "M-MidPSFF", "Marked Mid PS Fall Fing", 0.951606},
        '                               {23, "U-UWAc FF", "UnMarked UW Accelerated", 0.875365},
        '                               {24, "M-UWAc FF", "Marked UW Accelerated", 0.875365},
        '                               {25, "U-SPSd FF", "UnMarked South Puget Sound Fall Fing", 0.956446},
        '                               {26, "M-SPSd FF", "Marked South Puget Sound Fall Fing", 0.956421},
        '                               {27, "U-SPS Fyr", "UnMarked South Puget Sound Fall Year", 0.924627},
        '                               {28, "M-SPS Fyr", "Marked South Puget Sound Fall Year", 0.92434},
        '                               {29, "U-WhiteSp", "UnMarked White River Spring Fing", 0.956444},
        '                               {30, "M-WhiteSp", "Marked White River Spring Fing", 0.956444},
        '                               {31, "U-HdCl FF", "UnMarked Hood Canal Fall Fing", 0.938545},
        '                               {32, "M-HdCl FF", "Marked Hood Canal Fall Fing", 0.938637},
        '                               {33, "U-HdCl FY", "UnMarked Hood Canal Fall Year", 0.951887},
        '                               {34, "M-HdCl FY", "Marked Hood Canal Fall Year", 0.953089},
        '                               {35, "U-SJDF FF", "UnMarked JDF Tribs. Fall", 0.957195},
        '                               {36, "M-SJDF FF", "Marked JDF Tribs. Fall", 0.959135},
        '                               {37, "U-OR Tule", "UnMarked CR Oregon Hatchery Tule", 0.944107},
        '                               {38, "M-OR Tule", "Marked CR Oregon Hatchery Tule", 0.943285},
        '                               {39, "U-WA Tule", "UnMarked CR Washington Hatchery Tule", 0.973106},
        '                               {40, "M-WA Tule", "Marked CR Washington Hatchery Tule", 0.973184},
        '                               {41, "U-LCRWild", "UnMarked Lower Columbia River Wild", 0.983265},
        '                               {42, "M-LCRWild", "Marked Lower Columbia River Wild", 0.982486},
        '                               {43, "U-BPHTule", "UnMarked CR Bonneville Pool Hatchery", 0.922715},
        '                               {44, "M-BPHTule", "Marked CR Bonneville Pool Hatchery", 0.926528},
        '                               {45, "U-UpCR Su", "UnMarked Columbia R Upriver Summer", 0.981848},
        '                               {46, "M-UpCR Su", "Marked Columbia R Upriver Summer", 0.981849},
        '                               {47, "U-UpCR Br", "UnMarked Columbia R Upriver Bright", 0.968566},
        '                               {48, "M-UpCR Br", "Marked Columbia R Upriver Bright", 0.968613},
        '                               {49, "U-Cowl Sp", "UnMarked Cowlitz River Spring", -99.0},
        '                               {50, "M-Cowl Sp", "Marked Cowlitz River Spring", -99.0},
        '                               {51, "U-Will Sp", "UnMarked Willamette River Spring", 0.985291},
        '                               {52, "M-Will Sp", "Marked Willamette River Spring", 0.98522},
        '                               {53, "U-Snake F", "UnMarked Snake River Fall", 0.967637},
        '                               {54, "M-Snake F", "Marked Snake River Fall", 0.967652},
        '                               {55, "U-OR No F", "UnMarked Oregon North Coast Fall", 0.977251},
        '                               {56, "M-OR No F", "Marked Oregon North Coast Fall", 0.978782},
        '                               {57, "U-WCVI Tl", "UnMarked WCVI Total Fall", 0.977586},
        '                               {58, "M-WCVI Tl", "Marked WCVI Total Fall", 0.978774},
        '                               {59, "U-FrasRLt", "UnMarked Fraser River Late", 0.945673},
        '                               {60, "M-FrasRLt", "Marked Fraser River Late", 0.945376},
        '                               {61, "U-FrasREr", "UnMarked Fraser River Early", 0.986253},
        '                               {62, "M-FrasREr", "Marked Fraser River Early", 0.986363},
        '                               {63, "U-LwGeo S", "UnMarked Lower Georgia Strait", 0.806815},
        '                               {64, "M-LwGeo S", "Marked Lower Georgia Strait", 0.806292},
        '                               {65, "U-WhtSpYr", "UnMarked White River Spring Year", 0.972191},
        '                               {66, "M-WhtSpYr", "Marked White River Spring Year", 0.972191},
        '                               {67, "U-LColNat", "UnMarked Lower Columbia Naturals", 0.95597},
        '                               {68, "M-LColNat", "Marked Lower Columbia Naturals", 0.95597},
        '                               {69, "U-CentVal", "UnMarked Central Valley Fall", 0.966988},
        '                               {70, "M-CentVal", "Marked Central Valley Fall", 0.966831},
        '                               {71, "U-WA NCst", "UnMarked WA North Coast Fall", 0.980832},
        '                               {72, "M-WA NCst", "Marked WA North Coast Fall", 0.98092},
        '                               {73, "U-Willapa", "UnMarked Willapa Bay", 0.980222},
        '                               {74, "M-Willapa", "Marked Willapa Bay", 0.980282},
        '                               {75, "U-Hoko Rv", "UnMarked Hoko River", 0.978399},
        '                               {76, "M-Hoko Rv", "Marked Hoko River", 0.978319}}

        Dim NewTwos(NumStk + 10) As Double

        'NewTwos(0) = 0.961536
        'NewTwos(1) = 0.961456
        'NewTwos(2) = 0.979531
        'NewTwos(3) = 0.979533
        'NewTwos(4) = 0.979539
        'NewTwos(5) = 0.979539
        'NewTwos(6) = 0.931295
        'NewTwos(7) = 0.930408
        'NewTwos(8) = 0.963309
        'NewTwos(9) = 0.963309
        'NewTwos(10) = 0.978404
        'NewTwos(11) = 0.978334
        'NewTwos(12) = 0.934515
        'NewTwos(13) = 0.934488
        'NewTwos(14) = 0.948666
        'NewTwos(15) = 0.948623
        'NewTwos(16) = 0.932873
        'NewTwos(17) = 0.931865
        'NewTwos(18) = 0.97014
        'NewTwos(19) = 0.970262
        'NewTwos(20) = 0.951612
        'NewTwos(21) = 0.951606
        'NewTwos(22) = 0.875365
        'NewTwos(23) = 0.875365
        'NewTwos(24) = 0.956446
        'NewTwos(25) = 0.956421
        'NewTwos(26) = 0.924627
        'NewTwos(27) = 0.92434
        'NewTwos(28) = 0.956444
        'NewTwos(29) = 0.956444
        'NewTwos(30) = 0.938545
        'NewTwos(31) = 0.938637
        'NewTwos(32) = 0.951887
        'NewTwos(33) = 0.953089
        'NewTwos(34) = 0.957195
        'NewTwos(35) = 0.959135
        'NewTwos(36) = 0.944107
        'NewTwos(37) = 0.943285
        'NewTwos(38) = 0.973106
        'NewTwos(39) = 0.973184
        'NewTwos(40) = 0.983265
        'NewTwos(41) = 0.982486
        'NewTwos(42) = 0.922715
        'NewTwos(43) = 0.926528
        'NewTwos(44) = 0.981848
        'NewTwos(45) = 0.981849
        'NewTwos(46) = 0.968566
        'NewTwos(47) = 0.968613
        'NewTwos(48) = -99.0
        'NewTwos(49) = -99.0
        'NewTwos(50) = 0.985291
        'NewTwos(51) = 0.98522
        'NewTwos(52) = 0.967637
        'NewTwos(53) = 0.967652
        'NewTwos(54) = 0.977251
        'NewTwos(55) = 0.978782
        'NewTwos(56) = 0.977586
        'NewTwos(57) = 0.978774
        'NewTwos(58) = 0.945673
        'NewTwos(59) = 0.945376
        'NewTwos(60) = 0.986253
        'NewTwos(61) = 0.986363
        'NewTwos(62) = 0.806815
        'NewTwos(63) = 0.806292
        'NewTwos(64) = 0.972191
        'NewTwos(65) = 0.972191
        'NewTwos(66) = 0.95597
        'NewTwos(67) = 0.95597
        'NewTwos(68) = 0.966988
        'NewTwos(69) = 0.966831
        'NewTwos(70) = 0.980832
        'NewTwos(71) = 0.98092
        'NewTwos(72) = 0.980222
        'NewTwos(73) = 0.980282
        'NewTwos(74) = 0.978399
        'NewTwos(75) = 0.978319

        NewTwos(0) = 1.0082
        NewTwos(1) = 1.0094
        NewTwos(2) = 1.0189
        NewTwos(3) = 1.0196
        NewTwos(4) = 1.0181
        NewTwos(5) = 1.0181
        NewTwos(6) = 1.0062
        NewTwos(7) = 1.0058
        NewTwos(8) = 1.0202
        NewTwos(9) = 1.0202
        NewTwos(10) = 1.0018
        NewTwos(11) = 1.0015
        NewTwos(12) = 1.0036
        NewTwos(13) = 1.0034
        NewTwos(14) = 1.0388
        NewTwos(15) = 1.0388
        NewTwos(16) = 1.0103
        NewTwos(17) = 1.0109
        NewTwos(18) = 1.009
        NewTwos(19) = 1.0105
        NewTwos(20) = 1.0057
        NewTwos(21) = 1.0062
        NewTwos(22) = 1
        NewTwos(23) = 1
        NewTwos(24) = 0.9957
        NewTwos(25) = 0.9981
        NewTwos(26) = 1.0088
        NewTwos(27) = 0.9949
        NewTwos(28) = 1.0018
        NewTwos(29) = 1.0018
        NewTwos(30) = 1.0037
        NewTwos(31) = 1.0038
        NewTwos(32) = 1.0436
        NewTwos(33) = 1.0436
        NewTwos(34) = 1.0076
        NewTwos(35) = 1.0076
        NewTwos(36) = 1.0072
        NewTwos(37) = 1.008
        NewTwos(38) = 1.0018
        NewTwos(39) = 1.002
        NewTwos(40) = 1.0165
        NewTwos(41) = 1.0166
        NewTwos(42) = 1.0098
        NewTwos(43) = 1.0115
        NewTwos(44) = 1.0094
        NewTwos(45) = 1.0094
        NewTwos(46) = 1.0097
        NewTwos(47) = 1.0099
        NewTwos(48) = 1.014
        NewTwos(49) = 1.0209
        NewTwos(50) = 1.0036
        NewTwos(51) = 1.0039
        NewTwos(52) = 1.0097
        NewTwos(53) = 1.0098
        NewTwos(54) = 1.0027
        NewTwos(55) = 1.0027
        NewTwos(56) = 1.0018
        NewTwos(57) = 1.0019
        NewTwos(58) = 1.0054
        NewTwos(59) = 1.0061
        NewTwos(60) = 1.0029
        NewTwos(61) = 1.0029
        NewTwos(62) = 1.0628
        NewTwos(63) = 1.0629
        NewTwos(64) = 1.0051
        NewTwos(65) = 1.0051
        NewTwos(66) = 1.006
        NewTwos(67) = 1.006
        NewTwos(68) = 1.015
        NewTwos(69) = 1.015
        NewTwos(70) = 1.0019
        NewTwos(71) = 1.0019
        NewTwos(72) = 1.001
        NewTwos(73) = 1.001
        NewTwos(74) = 1.0002
        NewTwos(75) = 1.0004
        NewTwos(76) = 1.0022
        NewTwos(77) = 1.0021





        'Now do the recalculation
        'For s = 1 To NumStk
        '    If s < 49 Or s > 52 Then 'Skip past Cowlitz and Willamette
        '    End If
        'Next


      Dim Result As Integer
         Result = MsgBox("Age 2 scalars are about to be changed." & vbCrLf & "Changes Must be Saved before Running Model!!!" & vbCrLf & "Save Changes ???", MsgBoxStyle.YesNo)
         If Result = vbYes Then
                For s = 1 To NumStk
                    If s < 49 Or s > 52 Then 'Skip past Cowlitz and Willamette
                        Dim adjust As New Double
                        'adjust = dsNewTwos.Tables(0).Select("StockID = " & s)(0)("K_TwoThree")
                    adjust = NewTwos(s - 1)
                        tempRecruts(s) = StockRecruit(s, 3, 1) * adjust
                        StockRecruit(s, 2, 1) = tempRecruts(s)
                    End If
                Next
            Me.Visible = False
            FVS_SaveModelRunInputs.ShowDialog()
            Me.Visible = True
            RecordSetNameLabel.Text = RunIDNameSelect
            Me.BringToFront()
         Else
            MsgBox("Age 2 scalars haven't been changed", MsgBoxStyle.OkOnly)
         End If

End Sub


    Private Sub TransferBPBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TransferBPBtn.Click


        TransferBPName = ""
        CMDFileDialog.Filter = "Base Period Transfer Files (*.MDB)|*.MDB|All files (*.*)|*.*"
        CMDFileDialog.FilterIndex = 1
        CMDFileDialog.RestoreDirectory = True
        If CMDFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                TransferBPName = CMDFileDialog.FileName
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
        End If
        If TransferBPName = "" Then Exit Sub

        'TransferBPName = "NewFRAMBasePeriodTransfer.mdb"
        RecordsetSelectionType = 11
        If Exists(TransferBPName) Then
            Me.Visible = False
            FVS_ModelRunSelection.ShowDialog()
            If RecordsetSelectionType = 9 Then
                MsgBox("Model Run Transfer Cancelled", MsgBoxStyle.OkOnly)
                Exit Sub
            End If
            Me.Refresh()
            Me.Cursor = Cursors.WaitCursor
            '- Create Copy of Transfer Database File
            MDBSaveFileDialog.Filter = "*.mdb|*.mdb"

NewName:
            NewTransferBP = ""
            MDBSaveFileDialog.Filter = "Transfer File Name (*.MDB)|*.MDB|All files (*.*)|*.*"
            MDBSaveFileDialog.FilterIndex = 1
            MDBSaveFileDialog.FileName = "NewFRAMBasePeriodTransfer.Mdb"
            MDBSaveFileDialog.RestoreDirectory = True
            If MDBSaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Try
                    NewTransferBP = MDBSaveFileDialog.FileName
                Catch Ex As Exception
                    MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
                End Try
            End If
            If NewTransferBP = "" Then Exit Sub
            If NewTransferBP = "NewFRAMBasePeriodTransfer.Mdb" Then
                MsgBox("The file 'NewFRAMBasePeriodTransfeer.Mdb' is Reserved" & vbCrLf & _
                       "Please Choose Different Name for Transfer DataBase" & vbCrLf & _
                       "Prevents Corruption of Database Structure!", MsgBoxStyle.OkOnly)
                GoTo NewName
            End If

            'If Exists(FVSdatabasepath & "\" & NewTransferDB) Then Delete(FVSdatabasepath & "\" & NewTransferDB)
            'File.Copy(FVSdatabasepath & "\" & TransferDBName, NewTransferDB, True)
            If Exists(NewTransferBP) Then Delete(NewTransferBP)
            File.Copy(TransferBPName, NewTransferBP, True)



            '- TransferDB Connection String
            TransDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NewTransferBP

            Me.Cursor = Cursors.WaitCursor
            Call TransferBasePeriodTables()
            Me.Cursor = Cursors.Default
            Me.Visible = True
        Else
            MsgBox("Can't find BasePeriodTransfer.MDB file in the same directory as the FRAM database!" & vbCrLf & "Please Make Copy before Model Run Transfer", MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub GetBPTransferBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GetBPTransferBtn.Click
        Dim NewTransferBP As String

        NewTransferBP = ""
        CMDFileDialog.Filter = "Base Period Transfer Files (*.MDB)|*.MDB|All files (*.*)|*.*"
        CMDFileDialog.FilterIndex = 1
        CMDFileDialog.RestoreDirectory = True
        If CMDFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Try
                NewTransferBP = CMDFileDialog.FileName
            Catch Ex As Exception
                MessageBox.Show("Cannot read file from disk. Original error: " & Ex.Message)
            End Try
        End If
        If NewTransferBP = "" Then Exit Sub

        RecordsetSelectionType = 11
        If Exists(NewTransferBP) Then
            Me.Cursor = Cursors.WaitCursor
            '- TransferDB Connection String
            TransBP.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NewTransferBP
            Me.Visible = False
            FVS_BPTransfer.ShowDialog()

            Call GetTransferBasePeriodTables()
            
        End If
        Me.Visible = True
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenTransferModelRunFileDialog.FileOk

    End Sub
End Class