﻿Public Class FVS_PopStatScreen

   Private Sub FVS_PopStatScreen_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      Dim Col As Integer

      'FormHeight = 936
      FormHeight = 956
      FormWidth = 1215
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
         If FVS_PopStatScreen_ReSize = False Then
            Resize_Form(Me)
            FVS_PopStatScreen_ReSize = True
         End If
      End If

      PopStatGrid.Columns.Clear()
      PopStatGrid.ColumnHeadersDefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
      PopStatGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter
      If SpeciesName = "CHINOOK" Then

         'Old Chinook PopStat Bypass Code
         'Me.Close()
         'FVS_ScreenReports.Visible = True
         'Exit Sub

         '#####################################################################################
         'BEGIN PETE'S FEB 2013 CHINOOK POPSTAT CODE, CHUNK I...see also clipboard copy code below.
         '#####################################################################################
         PopStatGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         PopStatGrid.Columns.Add("Name", "StockName")
         PopStatGrid.Columns(0).Width = 110 / FormWidthScaler
         PopStatGrid.Columns(0).ReadOnly = True
         PopStatGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         PopStatGrid.Columns.Add("T0", "Age")
         PopStatGrid.Columns(1).Width = 30 / FormWidthScaler
         PopStatGrid.Columns(1).ReadOnly = True
         PopStatGrid.Columns(1).DefaultCellStyle.BackColor = Color.Aquamarine
         PopStatGrid.Columns.Add("T1", "T1-Coh")
         PopStatGrid.Columns(2).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T2", "T1-postNM")
         PopStatGrid.Columns(3).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T3", "T1-postPT")
         PopStatGrid.Columns(4).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T4", "T1-Mat")
         PopStatGrid.Columns(5).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T5", "T1-Esc")
         PopStatGrid.Columns(6).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T6", "T2-Coh")
         PopStatGrid.Columns(7).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(7).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T7", "T2-postNM")
         PopStatGrid.Columns(8).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(8).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T8", "T2-postPT")
         PopStatGrid.Columns(9).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(9).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T9", "T2-Mat")
         PopStatGrid.Columns(10).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(10).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T10", "T2-Esc")
         PopStatGrid.Columns(11).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(11).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T11", "T3-Coh")
         PopStatGrid.Columns(12).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(12).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T12", "T3-postNM")
         PopStatGrid.Columns(13).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(13).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T13", "T3-postPT")
         PopStatGrid.Columns(14).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(14).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(14).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T14", "T3-Mat")
         PopStatGrid.Columns(15).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(15).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(15).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T15", "T3-Esc")
         PopStatGrid.Columns(16).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(16).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(16).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T16", "T4-Coh")
         PopStatGrid.Columns(17).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(17).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T17", "T4-postNM")
         PopStatGrid.Columns(18).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(18).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T18", "T4-postPT")
         PopStatGrid.Columns(19).Width = 80 / FormWidthScaler
         PopStatGrid.Columns(19).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(19).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T19", "T4-Mat")
         PopStatGrid.Columns(20).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(20).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(20).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T20", "T4-Esc")
         PopStatGrid.Columns(21).Width = 60 / FormWidthScaler
         PopStatGrid.Columns(21).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(21).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight

         PopStatGrid.RowCount = NumStk * 4
         '#####################################################################################
         'END PETE'S FEB 2013 CHINOOK POPSTAT CODE, CHUNK I...see also clipboard copy code below.
         '#####################################################################################


      ElseIf SpeciesName = "COHO" Then
         PopStatGrid.DefaultCellStyle.Font = New Font("Microsoft San Serif", CInt(10 / FormWidthScaler), FontStyle.Bold)
         PopStatGrid.Columns.Add("Name", "StockName")
         PopStatGrid.Columns(0).Width = 125 / FormWidthScaler
         PopStatGrid.Columns(0).ReadOnly = True
         PopStatGrid.Columns(0).DefaultCellStyle.BackColor = Color.Aquamarine
         PopStatGrid.Columns.Add("T1", "StartCoh")
         PopStatGrid.Columns(1).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(1).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T2", "T1-Coht")
         PopStatGrid.Columns(2).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(2).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T3", "T1-Rem")
         PopStatGrid.Columns(3).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(3).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T4", "T2-Coht")
         PopStatGrid.Columns(4).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(4).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T5", "T2-Rem")
         PopStatGrid.Columns(5).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(5).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T6", "T3-Coht")
         PopStatGrid.Columns(6).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(6).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T7", "T3-Rem")
         PopStatGrid.Columns(7).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(7).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(7).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T8", "T4-Coht")
         PopStatGrid.Columns(8).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(8).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T9", "T4-Rem")
         PopStatGrid.Columns(9).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(9).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T10", "T5-Coht")
         PopStatGrid.Columns(10).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(10).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(10).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T11", "T5-Rem")
         PopStatGrid.Columns(11).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(11).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(11).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T12", "Mature")
         PopStatGrid.Columns(12).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(12).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(12).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.Columns.Add("T13", "Escape")
         PopStatGrid.Columns(13).Width = 65 / FormWidthScaler
         PopStatGrid.Columns(13).DefaultCellStyle.Format = ("########0")
         PopStatGrid.Columns(13).DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight
         PopStatGrid.RowCount = NumStk
      End If

      '- Put ERs into Grid
      If SpeciesName = "COHO" Then
         Age = 3
         For Stk As Integer = 1 To NumStk
            PopStatGrid.Item(0, Stk - 1).Value = StockName(Stk)
            '- Starting Cohort Time Step 1
            PopStatGrid.Item(1, Stk - 1).Value = Cohort(Stk, Age, 4, 1).ToString("########0")
            '- Cohort Sizes for each Time Step
            For TStep As Integer = 1 To NumSteps
               Col = TStep * 2
               PopStatGrid.Item(Col, Stk - 1).Value = Cohort(Stk, Age, 3, TStep).ToString("########0")
               PopStatGrid.Item(Col + 1, Stk - 1).Value = Cohort(Stk, Age, 2, TStep).ToString("########0")
            Next
            '- Mature Cohort
            PopStatGrid.Item(12, Stk - 1).Value = Cohort(Stk, Age, 1, 5).ToString("########0")
            '- Escapement
            PopStatGrid.Item(13, Stk - 1).Value = Escape(Stk, Age, 5).ToString("########0")
         Next


         '#####################################################################################
         'BEGIN PETE'S FEB 2013 CHINOOK POPSTAT CODE, CHUNK II...see also clipboard copy code below.
         '#####################################################################################
      ElseIf SpeciesName = "CHINOOK" Then
         Dim l As Integer 'Indexing variable to allow multipe rows per stock (age)
         l = 0
         For Stk As Integer = 1 To NumStk
            For Age As Integer = 2 To MaxAge
                    For TStep As Integer = 1 To NumSteps
                        
                        PopStatGrid.Item(0, l).Value = StockName(Stk)
                        PopStatGrid.Item(1, l).Value = Age
                        '- Starting Cohort Time Step 1
                        Col = 5 * (TStep - 1)
                        PopStatGrid.Item(Col + 2, l).Value = Cohort(Stk, Age, 4, TStep).ToString("########0")
                        '-After Nat Mort
                        PopStatGrid.Item(Col + 3, l).Value = Cohort(Stk, Age, 3, TStep).ToString("########0")
                        '-After PT Fishing
                        PopStatGrid.Item(Col + 4, l).Value = Cohort(Stk, Age, 2, TStep).ToString("########0")
                        '- Mature Cohort
                        PopStatGrid.Item(Col + 5, l).Value = Cohort(Stk, Age, 1, TStep).ToString("########0")
                        '- Escapement
                        PopStatGrid.Item(Col + 6, l).Value = Escape(Stk, Age, TStep).ToString("########0")
                    Next TStep
               l = l + 1
            Next Age
         Next Stk
      End If
      '#####################################################################################
      'END PETE'S FEB 2013 CHINOOK POPSTAT CODE, CHUNK II...see also clipboard copy code below.
      '#####################################################################################


   End Sub

   Private Sub PSExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles PSExit.Click
      Me.Close()
      FVS_ScreenReports.Visible = True
   End Sub

   Private Sub ClipBoardCopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClipBoardCopyToolStripMenuItem.Click
      '- Load String for Copy/Paste Report Output
      Dim ClipStr As String
      Dim RecNum, ColNum As Integer

      ClipStr = ""
      Clipboard.Clear()
      ClipStr = SpeciesName
      ClipStr &= "  {" & RunIDNameSelect & "}  " & RunIDRunTimeDateSelect.Date & vbCr

      If SpeciesName = "CHINOOK" Then
         '#####################################################################################
         'BEGIN PETE'S FEB 2013 CHINOOK POPSTAT CODE...CLIPBOARD COPY FUNCTIONALITY.
         '#####################################################################################
         ClipStr &= "StockName" & vbTab & "Age" & vbTab & "T1-StartCoh" & vbTab & "T1-postNM" & vbTab & "T1-postPT" & vbTab & "T1-Mat" & vbTab & "T1-Esc" & vbTab & "T2-StartCoh" & vbTab & "T2-postNM" & vbTab & "T2-postPT" & vbTab & "T2-Mat" & vbTab & "T2-Esc" & vbTab & "T3-StartCoh" & vbTab & "T3-postNM" & vbTab & "T3-postPT" & vbTab & "T3-Mat" & vbTab & "T3-Esc" & vbTab & "T4-StartCoh" & vbTab & "T4-postNM" & vbTab & "T4-postPT" & vbTab & "T4-Mat" & vbTab & "T4-Esc" & vbCr
         For RecNum = 0 To (NumStk * 4) - 1
            For ColNum = 0 To 21
               If ColNum = 0 Then
                  ClipStr = ClipStr & PopStatGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & CLng(PopStatGrid.Item(ColNum, RecNum).Value)
               End If
            Next
            ClipStr &= vbCr
         Next
         Clipboard.SetDataObject(ClipStr)

         '#####################################################################################
         'END PETE'S FEB 2013 CHINOOK POPSTAT CODE...CLIPBOARD COPY FUNCTIONALITY.
         '#####################################################################################

      ElseIf SpeciesName = "COHO" Then
            ClipStr &= "StockName" & vbTab & "StartCoh" & vbTab & "T1-Coht" & vbTab & "T1-Rem" & vbTab & "T2-Coht" & vbTab & "T2-Rem" & vbTab & "T3-Coht" & vbTab & "T3-Rem" & vbTab & "T4-Coht" & vbTab & "T4-Rem" & vbTab & "T5-Coht" & vbTab & "T5-Rem" & vbTab & "Mature" & vbTab & "Escape" & vbCr
         For RecNum = 0 To NumStk - 1
            For ColNum = 0 To 13
               If ColNum = 0 Then
                  ClipStr = ClipStr & PopStatGrid.Item(ColNum, RecNum).Value
               Else
                  ClipStr = ClipStr & vbTab & CLng(PopStatGrid.Item(ColNum, RecNum).Value)
               End If
            Next
            ClipStr &= vbCr
         Next
         Clipboard.SetDataObject(ClipStr)
      End If

   End Sub

   Private Sub PopStatGrid_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles PopStatGrid.CellContentClick

   End Sub
End Class