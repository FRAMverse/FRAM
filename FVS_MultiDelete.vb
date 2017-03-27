Imports System
Imports System.IO
Imports System.Math
Imports System.Text
Imports System.IO.File
Imports System.Data.OleDb
Imports System.Data

Public Class FVS_MultipleRunDeletion

   '- Run List Selection Variables
   Public Shared RunID(150) As Integer
   Public Shared RunIDName(150) As String
   Public Shared RunBasePeriodID(150) As Integer
   Public MultiDeleteList() As String

   'This subroutine gets the complete list of runs contained in the current FRAM database and preps it for populating the checkbox list
   Public Sub FillRunList()
      Dim FramDB As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & FVSdatabasename)
      FramDB.Open()
      Dim drd1 As OleDb.OleDbDataReader
      Dim cmd1 As New OleDb.OleDbCommand()
      cmd1.Connection = FramDB
      cmd1.CommandText = "SELECT * FROM RunID ORDER BY RunID"
      drd1 = cmd1.ExecuteReader
      Dim str1 As String
      Dim int1 As Integer
      int1 = 0
      list_MultiDelete.Items.Clear()
      If drd1.HasRows = False Then
         '- No RunID Recordsets .. Must Read Old CMD File
         RunIDSelect = 0
         RunIDNameSelect = "No Recordsets Available"
         Me.Close()
         FVS_MainMenu.Visible = True
      End If
      Do While drd1.Read
         '- Fill CheckedListBox Items
         str1 = String.Format("{0,5}-", drd1.GetInt32(1).ToString("####0"))
         str1 &= String.Format("{0,-7}-", drd1.GetString(2).ToString)
         str1 &= String.Format(" {0,-25} -", Mid(drd1.GetString(3).ToString, 1, 25))
         str1 &= String.Format("{0,-65}", Mid(drd1.GetString(4).ToString, 1, 65))
         list_MultiDelete.Items.Add(str1)
         '- Set RunID Array Values
         RunID(int1) = drd1.GetInt32(1)
         RunBasePeriodID(int1) = drd1.GetInt32(5)
         RunIDName(int1) = drd1.GetString(3)
         int1 = int1 + 1
      Loop
      FramDB.Close()

   End Sub

   'This loads the form and populates the check box list
   Private Sub FVS_MultipleRunDeletion_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
      FillRunList()
   End Sub

   'This takes the selection and passes it to a vector for passing to the delete subroutine
   Private Sub DeleteSelection_Click(sender As System.Object, e As System.EventArgs) Handles DeleteSelection.Click

      Me.Cursor = Cursors.WaitCursor

      Dim NumDeleteID As Integer
      Dim CheckForIt As Boolean
      CheckForIt = False
      NumDeleteID = list_MultiDelete.CheckedItems.Count
      multiRunDeleteMode = True

      'First make sure the run in use wasn't selected; if so, purge it.
      For Num = 0 To NumDeleteID - 1
         If CInt(list_MultiDelete.CheckedItems(Num).ToString.Substring(0, 5)) = RunIDSelect Then
            CheckForIt = True
            MsgBox("You selected the RunID that is CURRENTLY in use!!" & vbCrLf & "It will NOT be deleted", MsgBoxStyle.OkOnly)
            If NumDeleteID - 1 = 0 Then
               Exit Sub
               Me.Close()
               FVS_FramUtils.Visible = True
            End If
         End If
      Next

      'Shorten the list by one if the run in use was selected
      If CheckForIt = True Then
         NumDeleteID = NumDeleteID - 1
      End If

      'Now resize the vector of RunIDs to be deleted to whatever the final selection will be
      ReDim RunIDmultiDelete(NumDeleteID - 1)

      'Now populate the list for passing back and forth to delete mode
      For Num = 0 To NumDeleteID - 1
         RunIDmultiDelete(Num) = CInt(list_MultiDelete.CheckedItems(Num).ToString.Substring(0, 5))
      Next

      'Finally, iteratively call the delete recordset subroutine
      For i = 0 To RunIDmultiDelete.Length - 1
         multiRunPass = RunIDmultiDelete(i).ToString()
         Call DeleteRecordset()
      Next

      Me.Cursor = Cursors.Default

      'Done, back to the utilities menu.
      Me.Close()
      FVS_FramUtils.Visible = True
      FVS_ModelRunSelection.btn_DeleteMulti.Visible = False 'Make the multi-button invisible again

   End Sub

   'Exit back to utilities if you click CANCEL

   Private Sub CmdCancel_Click(sender As System.Object, e As System.EventArgs) Handles CmdCancel.Click

      Me.Close()
      FVS_FramUtils.Visible = True
      FVS_ModelRunSelection.btn_DeleteMulti.Visible = False 'Make the multi-button invisible again

   End Sub
End Class