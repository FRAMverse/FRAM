Public Class FVS_PSCMaxER

   Private Sub FVS_PSCMaxER_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

      FormHeight = 767
      FormWidth = 961
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
         If FVS_PSCMaxER_ReSize = False Then
            Resize_Form(Me)
            FVS_PSCMaxER_ReSize = True
         End If
      End If

      For Stk As Integer = 1 To 13
         Select Case Stk
            Case 1
               PSCInput1.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 2
               PSCInput2.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 3
               PSCInput3.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 4
               PSCInput4.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 5
               PSCInput5.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 6
               PSCInput6.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 7
               PSCInput7.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 8
               PSCInput8.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 9
               PSCInput9.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 10
               PSCInput10.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 11
               PSCInput11.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 12
               PSCInput12.Text = PSCMaxER(Stk).ToString("0.0000")
            Case 13
               PSCInput13.Text = PSCMaxER(Stk).ToString("0.0000")
         End Select
      Next
   End Sub

   Private Sub MERDoneButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MERDoneButton.Click
      For Stk As Integer = 1 To 13
         Select Case Stk
            Case 1
               If PSCMaxER(Stk) <> CDbl(PSCInput1.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput1.Text)
                  ChangePSCMaxER = True
               End If
            Case 2
               If PSCMaxER(Stk) <> CDbl(PSCInput2.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput2.Text)
                  ChangePSCMaxER = True
               End If
            Case 3
               If PSCMaxER(Stk) <> CDbl(PSCInput3.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput3.Text)
                  ChangePSCMaxER = True
               End If
            Case 4
               If PSCMaxER(Stk) <> CDbl(PSCInput4.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput4.Text)
                  ChangePSCMaxER = True
               End If
            Case 5
               If PSCMaxER(Stk) <> CDbl(PSCInput5.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput5.Text)
                  ChangePSCMaxER = True
               End If
            Case 6
               If PSCMaxER(Stk) <> CDbl(PSCInput6.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput6.Text)
                  ChangePSCMaxER = True
               End If
            Case 7
               If PSCMaxER(Stk) <> CDbl(PSCInput7.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput7.Text)
                  ChangePSCMaxER = True
               End If
            Case 8
               If PSCMaxER(Stk) <> CDbl(PSCInput8.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput8.Text)
                  ChangePSCMaxER = True
               End If
            Case 9
               If PSCMaxER(Stk) <> CDbl(PSCInput9.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput9.Text)
                  ChangePSCMaxER = True
               End If
            Case 10
               If PSCMaxER(Stk) <> CDbl(PSCInput10.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput10.Text)
                  ChangePSCMaxER = True
               End If
            Case 11
               If PSCMaxER(Stk) <> CDbl(PSCInput11.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput11.Text)
                  ChangePSCMaxER = True
               End If
            Case 12
               If PSCMaxER(Stk) <> CDbl(PSCInput12.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput12.Text)
                  ChangePSCMaxER = True
               End If
            Case 13
               If PSCMaxER(Stk) <> CDbl(PSCInput13.Text) Then
                  PSCMaxER(Stk) = CDbl(PSCInput13.Text)
                  ChangePSCMaxER = True
               End If
         End Select
      Next
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub

   Private Sub MER_CancelButton_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MER_CancelButton.Click
      Me.Close()
      FVS_InputMenu.Visible = True
   End Sub
End Class