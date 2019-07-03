Imports System.Threading
Public Class frm_signalCheck
#Region "Common Delegates"
    Delegate Sub SetObjectEnableDisable_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectEnableDisable_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectEnableDisable_Delegate(AddressOf SetObjectEnableDisable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Enabled = [value]
        End If
    End Sub

    Delegate Sub SetLabelText_Delegate(ByVal [label] As Label, ByVal [text] As String)
    Public Sub SetLabelText_ThreadSafe(ByVal [label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetLabelText_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelText_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelText_Delegate(AddressOf GetLabelText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetLabelTag_Delegate(ByVal [label] As Label, ByVal [tag] As String)
    Public Sub SetLabelTag_ThreadSafe(ByVal [label] As Label, ByVal [tag] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelTag_Delegate(AddressOf SetLabelTag_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [tag]})
        Else
            [label].Tag = [tag]
        End If
    End Sub
    Delegate Function GetLabelTag_Delegate(ByVal [label] As Label) As String

    Public Function GetLabelTag_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelTag_Delegate(AddressOf GetLabelTag_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Tag
        End If
    End Function
    Delegate Sub SetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
    Public Sub SetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New SetToolStripLabel_Delegate(AddressOf SetToolStripLabel_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[toolStrip], [label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
    Public Function GetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New GetToolStripLabel_Delegate(AddressOf GetToolStripLabel_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[toolStrip], [label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetDatagridBindDatatable_Delegate(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
    Public Sub SetDatagridBindDatatable_ThreadSafe(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [datagrid].InvokeRequired Then
            Dim MyDelegate As New SetDatagridBindDatatable_Delegate(AddressOf SetDatagridBindDatatable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[datagrid], [table]})
        Else
            [datagrid].DataSource = [table]
        End If
    End Sub
#End Region
    Dim dt As DataTable = Nothing
    Dim cts As CancellationTokenSource
    Private Sub test_Heartbeat(msg As String)
        SetLabelText_ThreadSafe(lbl_progress, msg)
    End Sub
    Private Sub frm_signalCheck_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmb_rule.SelectedIndex = My.Settings.SignalCheck
        nmrc_signalTimeFrame.Value = My.Settings.SignalTimeFrame
    End Sub
    Private Async Sub btn_view_Click(sender As Object, e As EventArgs) Handles btn_view.Click
        My.Settings.SignalCheck = cmb_rule.SelectedIndex
        My.Settings.SignalTimeFrame = nmrc_signalTimeFrame.Value
        My.Settings.Save()
        btn_view.Enabled = False
        btn_export.Enabled = False
        Dim startDate As Date = uc_date.dtpckr_from.Value.Date
        Dim endDate As Date = uc_date.dtpckr_to.Value.Date
        Dim signalTimeFrame As Integer = nmrc_signalTimeFrame.Value
        Select Case cmb_rule.SelectedIndex
            Case 0
                Await VWALLRuleAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
            Case 1
                Await ConstrictedRangeAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
            Case 2
                Await GapAndGoAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
            Case 3
                Await GapAndGoAfterGapFillAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
        End Select
        test_Heartbeat("Process Complete")
        SetObjectEnableDisable_ThreadSafe(btn_view, True)
        SetObjectEnableDisable_ThreadSafe(btn_export, True)
    End Sub
    Private Sub btn_export_Click(sender As Object, e As EventArgs) Handles btn_export.Click
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.FileName = "Output.csv"
        SaveFileDialog1.Filter = "CSV (*.csv)|*.csv"
        SaveFileDialog1.ShowDialog()
    End Sub
    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        Using export As New Utilities.DAL.CSVHelper(SaveFileDialog1.FileName, ",", cts)
            export.GetCSVFromDataGrid(dgrv_Signal)
        End Using
        If MessageBox.Show("Do you want to open file?", "Signal Check CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Process.Start(SaveFileDialog1.FileName)
        End If
    End Sub
    Private Sub cmb_rule_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_rule.SelectedIndexChanged
        Select Case cmb_rule.SelectedIndex
            Case 2
                Label2.Visible = False
                nmrc_signalTimeFrame.Visible = False
            Case 3
                Label2.Visible = False
                nmrc_signalTimeFrame.Visible = False
            Case Else
                Label2.Visible = True
                nmrc_signalTimeFrame.Visible = True
        End Select
    End Sub
    Private Async Function VWALLRuleAsync(ByVal startDate As Date, ByVal endDate As Date, ByVal signalTimeFrame As Integer) As Task
        Dim rule As VWALLRule = New VWALLRule(cts)
        AddHandler rule.Heartbeat, AddressOf test_Heartbeat
        Dim dt As DataTable = Await rule.TestDataAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
        SetDatagridBindDatatable_ThreadSafe(dgrv_Signal, dt)
    End Function
    Private Async Function ConstrictedRangeAsync(ByVal startDate As Date, ByVal endDate As Date, ByVal signalTimeFrame As Integer) As Task
        Dim rule As ConstrictedRange = New ConstrictedRange(cts)
        AddHandler rule.Heartbeat, AddressOf test_Heartbeat
        Dim dt As DataTable = Await rule.TestDataAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
        SetDatagridBindDatatable_ThreadSafe(dgrv_Signal, dt)
    End Function
    Private Async Function GapAndGoAsync(ByVal startDate As Date, ByVal endDate As Date, ByVal signalTimeFrame As Integer) As Task
        Dim rule As GapAndGo = New GapAndGo(cts)
        AddHandler rule.Heartbeat, AddressOf test_Heartbeat
        Dim dt As DataTable = Await rule.TestDataAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
        SetDatagridBindDatatable_ThreadSafe(dgrv_Signal, dt)
    End Function
    Private Async Function GapAndGoAfterGapFillAsync(ByVal startDate As Date, ByVal endDate As Date, ByVal signalTimeFrame As Integer) As Task
        Dim rule As GapAndGoAfterGapFill = New GapAndGoAfterGapFill(cts)
        AddHandler rule.Heartbeat, AddressOf test_Heartbeat
        Dim dt As DataTable = Await rule.TestDataAsync(startDate, endDate, signalTimeFrame).ConfigureAwait(False)
        SetDatagridBindDatatable_ThreadSafe(dgrv_Signal, dt)
    End Function
End Class