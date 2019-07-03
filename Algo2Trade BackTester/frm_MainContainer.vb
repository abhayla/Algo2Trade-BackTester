Public Class frm_MainContainer
    Private Sub IndicatorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IndicatorToolStripMenuItem.Click
        Dim frmToShow As New frm_Indicator

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Algo2Trade BackTester v" + My.Application.Info.Version.ToString + " - " + My.Application.Info.Description.ToString
        'Me.Size = Screen.PrimaryScreen.WorkingArea.Size
    End Sub

    Private Sub BackTestToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BackTestToolStripMenuItem.Click
        Dim frmToShow As New frm_BackTest

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub SignalCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SignalCheckToolStripMenuItem.Click
        Dim frmToShow As New frm_signalCheck

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub
End Class