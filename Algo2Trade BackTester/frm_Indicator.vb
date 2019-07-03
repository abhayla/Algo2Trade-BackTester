Imports System.Threading
Public Class frm_Indicator
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
#End Region
    Dim dt As DataTable = Nothing
    Dim cts As CancellationTokenSource
    Private Sub frm_Indicator_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        pnl_ema.Visible = False
        pnl_fractals.Visible = False
        pnl_sma.Visible = False
        pnl_smi.Visible = False
        pnl_atr.Visible = False
        pnl_cci.Visible = False
        pnl_rsi.Visible = False
        pnl_macd.Visible = False

        btn_export.Enabled = False
        cmb_indicator.SelectedIndex = My.Settings.Indicator
        uc_Indicator.cmb_inscategory.SelectedIndex = My.Settings.IndicatorStockType
        uc_Indicator.txt_instrument.Text = My.Settings.IndicatorStockName
        Control.CheckForIllegalCrossThreadCalls = False
        ProgressBar1.Visible = False
    End Sub

    Private Sub test_Heartbeat(msg As String)
        SetLabelText_ThreadSafe(lblProgressStatus, msg)
    End Sub

    Private Sub btn_view_Click(sender As Object, e As EventArgs) Handles btn_view.Click
        My.Settings.Indicator = cmb_indicator.SelectedIndex
        My.Settings.IndicatorStockType = uc_Indicator.cmb_inscategory.SelectedIndex
        My.Settings.IndicatorStockName = uc_Indicator.txt_instrument.Text
        My.Settings.Save()

        btn_export.Enabled = False
        btn_view.Enabled = False
        dt = New DataTable
        DataGridView1.Visible = False
        ProgressBar1.Visible = True
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub cmb_indicator_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmb_indicator.SelectedIndexChanged
        Dim i As Integer = cmb_indicator.SelectedIndex
        Select Case i
            Case 0
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 1
                pnl_ema.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
                pnl_fractals.Visible = True
            Case 2
                pnl_fractals.Visible = False
                pnl_sma.Visible = True
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 3
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_ema.Visible = True
            Case 4
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
                pnl_smi.Visible = True
            Case 5
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
                pnl_atr.Visible = True
            Case 6
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
                pnl_cci.Visible = True
            Case 7
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_macd.Visible = False
                pnl_rsi.Visible = True
            Case 8
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = True
            Case 9
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 10
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 11
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 12
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 13
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 14
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_ema.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 15
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 16
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 17
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 18
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 19
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 20
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 21
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 22
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 23
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False
            Case 24
                pnl_ema.Visible = False
                pnl_fractals.Visible = False
                pnl_sma.Visible = False
                pnl_smi.Visible = False
                pnl_atr.Visible = False
                pnl_cci.Visible = False
                pnl_rsi.Visible = False
                pnl_macd.Visible = False

        End Select
    End Sub

    Private Sub btn_export_Click(sender As Object, e As EventArgs) Handles btn_export.Click
        Dim i As Integer = cmb_indicator.SelectedIndex
        Dim cancel As Threading.CancellationTokenSource = Nothing
        Select Case i
            Case 0
                Using export As New Utilities.DAL.CSVHelper("VWAP Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("VWAP Output.csv")
                End If
            Case 1
                Using export As New Utilities.DAL.CSVHelper("Fractals Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Fractals Output.csv")
                End If
            Case 2
                Using export As New Utilities.DAL.CSVHelper("SMA Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("SMA Output.csv")
                End If
            Case 3
                Using export As New Utilities.DAL.CSVHelper("EMA Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("EMA Output.csv")
                End If
            Case 4
                Using export As New Utilities.DAL.CSVHelper("SMI Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("SMI Output.csv")
                End If
            Case 5
                Using export As New Utilities.DAL.CSVHelper("ATR Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("ATR Output.csv")
                End If
            Case 6
                Using export As New Utilities.DAL.CSVHelper("CCI Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("CCI Output.csv")
                End If
            Case 7
                Using export As New Utilities.DAL.CSVHelper("RSI Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("RSI Output.csv")
                End If
            Case 8
                Using export As New Utilities.DAL.CSVHelper("MACD Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("MACD Output.csv")
                End If
            Case 9
                Using export As New Utilities.DAL.CSVHelper("Heiken Ashi Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Heiken Ashi Output.csv")
                End If
            Case 10
                Using export As New Utilities.DAL.CSVHelper("Supertrend Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Supertrend Output.csv")
                End If
            Case 11
                Using export As New Utilities.DAL.CSVHelper("ATR Trailing Stop Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("ATR Trailing Stop Output.csv")
                End If
            Case 12
                Using export As New Utilities.DAL.CSVHelper("Fractal1 Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Fractal1 Output.csv")
                End If
            Case 13
                Using export As New Utilities.DAL.CSVHelper("Fractal With SMA Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Fractal With SMA Output.csv")
                End If
            Case 14
                Using export As New Utilities.DAL.CSVHelper("Fractal Retracement Entry Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Fractal Retracement Entry Rule Output.csv")
                End If
            Case 15
                Using export As New Utilities.DAL.CSVHelper("ATR Bands Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("ATR Bands Output.csv")
                End If
            Case 16
                Using export As New Utilities.DAL.CSVHelper("Naughty Boy VWAP Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Naughty Boy VWAP Rule Output.csv")
                End If
            Case 17
                Using export As New Utilities.DAL.CSVHelper("VWAP Double Confirmation Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("VWAP Double Confirmation Rule Output.csv")
                End If
            Case 18
                Using export As New Utilities.DAL.CSVHelper("In The Trend Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("In The Trend Rule Output.csv")
                End If
            Case 19
                Using export As New Utilities.DAL.CSVHelper("In The Trend Pearsing Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                    'export.GetCSVFromDataGrid(DataGridView1)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("In The Trend Pearsing Rule Output.csv")
                End If
            Case 20
                Using export As New Utilities.DAL.CSVHelper("Swing High Low Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Swing High Low Output.csv")
                End If
            Case 21
                Using export As New Utilities.DAL.CSVHelper("Double U Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Double U Rule Output.csv")
                End If
            Case 22
                Using export As New Utilities.DAL.CSVHelper("MR With One Fractal And MA Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("MR With One Fractal And MA Rule Output.csv")
                End If
            Case 23
                Using export As New Utilities.DAL.CSVHelper("JOYMA4 Now I Should Go Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("JOYMA4 Now I Should Go Rule Output.csv")
                End If
            Case 24
                Using export As New Utilities.DAL.CSVHelper("Higher High Higher Low Pattern Rule Output.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Higher High Higher Low Pattern Rule Output.csv")
                End If
            Case 25
                Using export As New Utilities.DAL.CSVHelper("Bollinger Bands.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Bollinger Bands.csv")
                End If
            Case 26
                Using export As New Utilities.DAL.CSVHelper("Algo2Trade Renko.csv", ",", cancel)
                    export.GetCSVFromDataTable(dt)
                End Using
                If MessageBox.Show("Do you want to open file?", "Indicator CSV File", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    Process.Start("Algo2Trade Renko.csv")
                End If
        End Select
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Dim worker As System.ComponentModel.BackgroundWorker = DirectCast(sender, System.ComponentModel.BackgroundWorker)

        worker.ReportProgress(1)
        Dim indicator As New IndicatorHelper(uc_Indicator.txt_instrument.Text, uc_Indicator.cmb_inscategory.SelectedIndex, uc_Indicator.chkb_open.Checked, uc_Indicator.chkb_close.Checked, uc_Indicator.chkb_high.Checked, uc_Indicator.chkb_low.Checked, uc_Indicator.chkb_volume.Checked, cts)
        AddHandler indicator.Heartbeat, AddressOf test_Heartbeat
        Dim outputdt As DataTable = Nothing
        worker.ReportProgress(3)
        worker.ReportProgress(10)
        Dim chk_date As DateTime = Convert.ToDateTime(uc_Indicator.dtpckr_from.Value.ToString)
        Dim trgt_date As DateTime = Convert.ToDateTime(uc_Indicator.dtpckr_to.Value.ToString)
        Dim totalDays As TimeSpan = trgt_date.Subtract(chk_date)
        Dim count As Integer = Math.Floor(85 / (totalDays.Days + 1))
        Dim progresscounter = 11
        Try
            While Not chk_date.ToString("yyyy-MM-dd") > trgt_date.ToString("yyyy-MM-dd")
                progresscounter = progresscounter + (count)
                Dim i As Integer = cmb_indicator.SelectedIndex
                Select Case i
                    Case 0
                        outputdt = indicator.VWAP_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 1
                        outputdt = indicator.Fractals_Calculate(nmrc_fractalsPeriod.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 2
                        outputdt = indicator.SMA_Calculate(nmrc_smaperiod.Value, cmb_smafields.SelectedIndex, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 3
                        outputdt = indicator.EMA_Calculate(nmrc_emaPeriod.Value, cmb_emafield.SelectedIndex, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 4
                        outputdt = indicator.SMI_Calculate(nmrc_K.Value, nmrc_Ksmoothing.Value, nmrc_KdoubleSmoothing.Value, nmrc_D.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 5
                        outputdt = indicator.ATR_Calculate(nmrc_atr.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 6
                        outputdt = indicator.CCI_Calculate(nmrc_cciPeriod.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 7
                        outputdt = indicator.RSI_Calculate(nmrc_rsiPeriod.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 8
                        outputdt = indicator.MACD_Calculate(nmrc_macdFast.Value, nmrc_macdSlow.Value, nmrc_macdSignal.Value, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 9
                        outputdt = indicator.HikenAshiConversion(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 10
                        outputdt = indicator.Supertrend_Calculate(10, 3, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 11
                        outputdt = indicator.ATRTrailingStop_Calculate(5, 3, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 12
                        outputdt = indicator.Fractal1_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 13
                        outputdt = indicator.FractalWithSMA_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 14
                        outputdt = indicator.FractalRetracementEntry_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 15
                        outputdt = indicator.ATRBands_Calculate(3, 14, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 16
                        outputdt = indicator.NaughtyBoyVWAP_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 17
                        outputdt = indicator.VWAPDoubleConfirmation_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 18
                        outputdt = indicator.InTheTrend_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 19
                        outputdt = indicator.InTheTrendPearsing_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 20
                        outputdt = indicator.SwingHighLow_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 21
                        outputdt = indicator.DoubleU_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 22
                        outputdt = indicator.MRWithOneFractalAndMA_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 23
                        outputdt = indicator.JOYMA4_NowIShouldGo_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 24
                        outputdt = indicator.HigherHighHigherLowPattern_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 25
                        outputdt = indicator.BollingerBands_Calculate(20, 2, chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 26
                        outputdt = indicator.Renko_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                    Case 27
                        outputdt = indicator.TII_Calculate(chk_date.ToString("yyyy-MM-dd"))
                        If outputdt IsNot Nothing Then
                            dt.Merge(outputdt)
                        End If
                End Select
                worker.ReportProgress(progresscounter)
                chk_date = chk_date.AddDays(1)
            End While
        Catch ex As Exception
            MsgBox(ex.ToString())
        End Try
        worker.ReportProgress(99)
        System.Threading.Thread.Sleep(1000)
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        ProgressBar1.Visible = False
        DataGridView1.Visible = True
        DataGridView1.AutoGenerateColumns = True
        DataGridView1.DataSource = dt
        DataGridView1.Refresh()
        btn_view.Enabled = True
        btn_export.Enabled = True
        test_Heartbeat("Execution Complete")
        BackgroundWorker1.Dispose()
    End Sub

    Private Sub BackgroundWorker1_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles BackgroundWorker1.ProgressChanged
        Me.ProgressBar1.Value = e.ProgressPercentage
    End Sub
End Class