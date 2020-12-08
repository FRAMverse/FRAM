Public Class FVS_BPTransfer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    

    Private Sub FVS_BPTransfer_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'BPSizeLimitchk.Checked = False
        Fisherieschk.Checked = False
        Stockschk.Checked = False
        TimeStepschk.Checked = False
    End Sub

    Private Sub Stockschk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Stockschk.CheckedChanged
        If Stockschk.Checked = True Then
            ImportStock = True
        End If
    End Sub

    Private Sub Fisherieschk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Fisherieschk.CheckedChanged
        If Fisherieschk.Checked = True Then
            ImportFish = True
        End If
    End Sub

    Private Sub TimeStepschk_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeStepschk.CheckedChanged
        If TimeStepschk.Checked = True Then
            ImportTS = True
        End If
    End Sub
End Class