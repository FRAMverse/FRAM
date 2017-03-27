Public Class FVS_TRunTimeSteps

   Private Sub FVS_TRunTimeSteps_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
      If SpeciesName = "COHO" Then
         ComboBox1.Items.Clear()
         ComboBox1.Items.Add("Time-1 Jan-June")
         ComboBox1.Items.Add("Time-2 July")
         ComboBox1.Items.Add("Time-3 August")
         ComboBox1.Items.Add("Time-4 September")
         ComboBox1.Items.Add("Time-5 Oct-Dec")
         ComboBox1.SelectedIndex = 3
         ComboBox2.Items.Clear()
         ComboBox2.Items.Add("Time-1 Jan-June")
         ComboBox2.Items.Add("Time-2 July")
         ComboBox2.Items.Add("Time-3 August")
         ComboBox2.Items.Add("Time-4 September")
         ComboBox2.Items.Add("Time-5 Oct-Dec")
         ComboBox2.SelectedIndex = 4
      ElseIf SpeciesName = "CHINOOK" Then
         ComboBox1.Items.Clear()
         ComboBox1.Items.Add("Time-1 Oct-Apr")
         ComboBox1.Items.Add("Time-2 May-June")
         ComboBox1.Items.Add("Time-3 July-Sept")
         ComboBox1.Items.Add("Time-4 Oct-Apr")
         ComboBox1.SelectedIndex = 2
         ComboBox2.Items.Clear()
         ComboBox2.Items.Add("Time-1 Oct-Apr")
         ComboBox2.Items.Add("Time-2 May-June")
         ComboBox2.Items.Add("Time-3 July-Sept")
         ComboBox2.Items.Add("Time-4 Oct-Apr")
         ComboBox2.SelectedIndex = 2
      End If
      ComboBox3.Items.Clear()
      ComboBox3.Items.Add("ETRS-ExtremeTerm")
      ComboBox3.Items.Add("TAA-TermAreaAbun")
      ComboBox3.SelectedIndex = 0

   End Sub

   Private Sub TRTSDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles TRTSDoneButton.Click
      TimeStepSelection1 = ComboBox1.SelectedIndex + 1
      TimeStepSelection2 = ComboBox2.SelectedIndex + 1
      TermRunTypeSelection = ComboBox3.SelectedIndex
      Me.Close()
      FVS_ReportSelection.Visible = True
      Exit Sub
   End Sub

End Class