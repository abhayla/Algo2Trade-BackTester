Imports MySql.Data.MySqlClient
Imports System.Threading
Imports Algo2TradeBLL
Imports System.IO

Public Class frm_BackTest

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

    Delegate Function GetDateTimePickerValue_Delegate(ByVal [dateTimePicker] As DateTimePicker) As Date
    Public Function GetDateTimePickerValue_ThreadSafe(ByVal [dateTimePicker] As DateTimePicker) As Date
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [dateTimePicker].InvokeRequired Then
            Dim MyDelegate As New GetDateTimePickerValue_Delegate(AddressOf GetDateTimePickerValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New DateTimePicker() {[dateTimePicker]})
        Else
            Return [dateTimePicker].Value
        End If
    End Function

    Delegate Function GetNumericUpDownValue_Delegate(ByVal [numericUpDown] As NumericUpDown) As Integer
    Public Function GetNumericUpDownValue_ThreadSafe(ByVal [numericUpDown] As NumericUpDown) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [numericUpDown].InvokeRequired Then
            Dim MyDelegate As New GetNumericUpDownValue_Delegate(AddressOf GetNumericUpDownValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New NumericUpDown() {[numericUpDown]})
        Else
            Return [numericUpDown].Value
        End If
    End Function

    Delegate Function GetComboBoxIndex_Delegate(ByVal [combobox] As ComboBox) As Integer
    Public Function GetComboBoxIndex_ThreadSafe(ByVal [combobox] As ComboBox) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [combobox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxIndex_Delegate(AddressOf GetComboBoxIndex_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[combobox]})
        Else
            Return [combobox].SelectedIndex
        End If
    End Function

    Delegate Function GetComboBoxItem_Delegate(ByVal [ComboBox] As ComboBox) As String
    Public Function GetComboBoxItem_ThreadSafe(ByVal [ComboBox] As ComboBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [ComboBox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxItem_Delegate(AddressOf GetComboBoxItem_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[ComboBox]})
        Else
            Return [ComboBox].SelectedItem.ToString
        End If
    End Function

    Delegate Function GetTextBoxText_Delegate(ByVal [textBox] As TextBox) As String
    Public Function GetTextBoxText_ThreadSafe(ByVal [textBox] As TextBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [textBox].InvokeRequired Then
            Dim MyDelegate As New GetTextBoxText_Delegate(AddressOf GetTextBoxText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[textBox]})
        Else
            Return [textBox].Text
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

    Private _cts As CancellationTokenSource
    Private _cmn As Common
    Private _dataSource As Strategy.SourceOfData = Strategy.SourceOfData.None
    Private _includeSlippage As Boolean = False
    Private _slippageMultiplier As Decimal = 1
    Private Sub frm_BackTest_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        btn_cancel.Visible = False
        cmb_strategy.SelectedIndex = My.Settings.BackTest
        rdbDatabase.Checked = My.Settings.DataSourceDatabase
        rdbLive.Checked = My.Settings.DataSourceLive
        chkbIncludeSlippage.Checked = My.Settings.IncludeSlippage
        txtSlippageMultiplier.Text = My.Settings.SlippageMultiplier
    End Sub
    Private Sub OnHeartbeat(msg As String)
        SetLabelText_ThreadSafe(lblProgressStatus, msg)
    End Sub
    Private Async Sub btn_start_Click(sender As Object, e As EventArgs) Handles btn_start.Click
        My.Settings.BackTest = cmb_strategy.SelectedIndex
        My.Settings.DataSourceDatabase = rdbDatabase.Checked
        My.Settings.DataSourceLive = rdbLive.Checked
        My.Settings.IncludeSlippage = chkbIncludeSlippage.Checked
        My.Settings.SlippageMultiplier = txtSlippageMultiplier.Text
        My.Settings.Save()
        If rdbDatabase.Checked Then
            _dataSource = Strategy.SourceOfData.Database
        ElseIf rdbLive.Checked Then
            _dataSource = Strategy.SourceOfData.Live
        End If
        _includeSlippage = chkbIncludeSlippage.Checked
        _slippageMultiplier = txtSlippageMultiplier.Text
        _cts = New CancellationTokenSource
        _cmn = New Common(_cts)
        btn_start.Enabled = False
        Await Task.Run(AddressOf ViewDataAsync).ConfigureAwait(False)
    End Sub
    Private Async Function ViewDataAsync() As Task
        Try
            Dim startDate As Date = GetDateTimePickerValue_ThreadSafe(uc_BackTest.dtpckr_from)
            Dim endDate As Date = GetDateTimePickerValue_ThreadSafe(uc_BackTest.dtpckr_to)
            Dim rule As Integer = GetComboBoxIndex_ThreadSafe(cmb_strategy)
            Select Case rule
                Case 0
                'Await FractalMACandleRetracementStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 1
                    'Await NaughtyBoyStockVWAPStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 2
                    'Await InTheTrendStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 3
                    'Await GapFillStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 4
                    'Await GapFillWithoutFillingPCStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 5
                    'Await GapFillWithoutFillingNeglectGnGStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 6
                    'Await GapFillWithoutFillingPHLCStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 7
                    Await GenericStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 8
                    'Await MarutiStrategyAsync(startDate, endDate).ConfigureAwait(False)
                Case 9
                    'Await PairHazingStrategyAsync(startDate, endDate).ConfigureAwait(False)
            End Select
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            OnHeartbeat("Process Complete")
            SetObjectEnableDisable_ThreadSafe(btn_start, True)
            SetObjectEnableDisable_ThreadSafe(btn_cancel, False)
        End Try
    End Function

    Private Async Function GenericStrategyAsync(startDate As Date, endDate As Date) As Task
        Dim strategyStockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures
        Dim tradeStartDate As Date = startDate.Date
        While tradeStartDate.Date <= endDate.Date
            Dim tradeEndDate As Date '= tradeStartDate.AddMonths(12)
            tradeEndDate = endDate.Date

            If endDate.Date < tradeEndDate Then tradeEndDate = endDate
            For timeframe As Integer = 1 To 1 Step 2
                For tradeMultiplier As Double = 4 To 4 Step 1
                    For trlng As Integer = 0 To 0 Step 1
                        For smdirectiocEntry As Integer = 1 To 1 Step 1
                            'If trlng = 0 And smdirectiocEntry = 1 Then Continue For
                            For countBreakevenTrades As Integer = 1 To 1 Step 1
                                'If trlng = 0 And countBreakevenTrades = 1 Then Continue For
                                For overAllLoss As Decimal = 40000 To 40000 Step 10000
                                    For nmbrOfTrade As Integer = 100 To 100 Step 1
                                        Using backtestStrategy As New GenericStrategy(canceller:=_cts,
                                                                                        tickSize:=0.05,
                                                                                        eodExitTime:=TimeSpan.Parse("15:15:00"),
                                                                                        lastTradeEntryTime:=TimeSpan.Parse("14:30:00"),
                                                                                        exchangeStartTime:=TimeSpan.Parse("09:15:00"),
                                                                                        exchangeEndTime:=TimeSpan.Parse("15:29:00"),
                                                                                        signalTimeFrame:=timeframe,
                                                                                        UseHeikenAshi:=False,
                                                                                        associatedStockData:=Nothing,
                                                                                        stockType:=strategyStockType)
                                            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat

                                            With backtestStrategy
                                                .BothSideSlippage = True
                                                .DataSource = _dataSource
                                                .IncludeSlippage = _includeSlippage
                                                .SlippageMultiplier = _slippageMultiplier

                                                .InitialCapital = 500000
                                                .CapitalForPumpIn = 400000
                                                .MinimumEarnedCapitalToWithdraw = 600000
                                                .AmountToBeWithdrawn = 100000

                                                .StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Future Stock List ATR Based.csv")

                                                '1 from investment, 2 from SL, 3 from futures lot
                                                .QuantityFlag = 3
                                                .MaxStoplossAmount = 1000
                                                .FirstTradeTargetMultiplier = tradeMultiplier
                                                .EarlyStoploss = False
                                                .ForwardTradeTargetMultiplier = tradeMultiplier
                                                .CapitalToBeUsed = 20000
                                                .CandleBasedEntry = True

                                                Select Case strategyStockType
                                                    Case Trade.TypeOfStock.Cash
                                                        Strategy.MarginMultiplier = 13
                                                    Case Trade.TypeOfStock.Commodity
                                                        Strategy.MarginMultiplier = 70
                                                    Case Trade.TypeOfStock.Currency
                                                        Strategy.MarginMultiplier = 70
                                                    Case Trade.TypeOfStock.Futures
                                                        Strategy.MarginMultiplier = 30
                                                End Select
                                                .NumberOfTradeableStockPerDay = 5
                                                .NumberOfTradePerDay = Integer.MaxValue
                                                .NumberOfTradePerStockPerDay = nmbrOfTrade
                                                .CountTradesWithBreakevenMovement = countBreakevenTrades
                                                .TrailingSL = trlng
                                                .SameDirectionTrade = smdirectiocEntry
                                                .ReverseSignalTrade = False
                                                .ModifyTarget = False
                                                .ModifyStoploss = False
                                                .StopAtTargetReach = False
                                                .ExitOnStockFixedTargetStoploss = False
                                                .StockMaxProfitPerDay = Decimal.MaxValue
                                                .StockMaxLossPerDay = Decimal.MinValue
                                                .ExitOnOverAllFixedTargetStoploss = True
                                                .OverAllProfitPerDay = 7500
                                                .OverAllLossPerDay = -5000
                                            End With
                                            Await backtestStrategy.TestStrategyAsync(tradeStartDate, tradeEndDate).ConfigureAwait(False)
                                        End Using
                                    Next
                                Next
                            Next
                        Next
                    Next
                Next
            Next
            tradeStartDate = tradeEndDate.AddDays(1)
        End While
    End Function

#Region "Generic Strategy Dump"
    'Private Async Function GenericStrategyAsync(startDate As Date, endDate As Date) As Task
    '    'Dim stockData As Dictionary(Of String, Decimal()) = GetStockData(New Date(2019, 2, 15))
    '    Dim stockData As Dictionary(Of String, Decimal()) = Nothing
    '    'Dim stockData As Dictionary(Of String, Decimal()) = New Dictionary(Of String, Decimal())
    '    'stockData.Add("JINDALSTEL", {0, 0})
    '    'Dim stockData As Dictionary(Of String, Decimal()) = Nothing
    '    If stockData IsNot Nothing AndAlso stockData.Count > 0 Then
    '        For Each stock In stockData.Keys
    '            Dim dummyStockData As Dictionary(Of String, Decimal()) = New Dictionary(Of String, Decimal())
    '            dummyStockData.Add(stock, stockData(stock))
    '            Dim backtestStrategy As GenericStrategy = New GenericStrategy(canceller:=cts,
    '                                                                  tickSize:=0.05,
    '                                                                  eodExitTime:=TimeSpan.Parse("12:30:00"),
    '                                                                  lastTradeEntryTime:=TimeSpan.Parse("9:21:00"),
    '                                                                  exchangeStartTime:=TimeSpan.Parse("09:15:00"),
    '                                                                  exchangeEndTime:=TimeSpan.Parse("15:29:00"),
    '                                                                  signalTimeFrame:=1,
    '                                                                  UseHeikenAshi:=False,
    '                                                                  associatedStockData:=dummyStockData,
    '                                                                  stockType:=Trade.TypeOfStock.Cash)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            'For nmbrTrd As Integer = 2 To 2
    '            'For trl As Integer = 0 To 0
    '            'For rvrs As Integer = 0 To 0
    '            'For mtm As Integer = 0 To 0
    '            'For tmt = 2 To 2
    '            'For slmt = 1 To 1
    '            'If slmt <= tmt Then
    '            'For mp = 8000 To 8000 Step 10000
    '            'For ml = 8000 To 8000 Step 5000
    '            'If ml <= mp Then
    '            Strategy.MarginMultiplier = 30
    '            backtestStrategy.NumberOfTradePerDay = 15
    '            backtestStrategy.NumberOfTradePerStockPerDay = 1
    '            backtestStrategy.TrailingSL = False
    '            backtestStrategy.ReverseSignalTrade = False
    '            backtestStrategy.ModifyTarget = False
    '            backtestStrategy.ModifyStoploss = False
    '            backtestStrategy.ExitOnStockFixedTargetStoploss = True
    '            backtestStrategy.StockMaxProfitPerDay = 8500
    '            backtestStrategy.StockMaxLossPerDay = -5000
    '            backtestStrategy.ExitOnOverAllFixedTargetStoploss = True
    '            backtestStrategy.OverAllProfitPerDay = 8500
    '            backtestStrategy.OverAllLossPerDay = -5000
    '            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '            'End If
    '            'Next
    '            'Next
    '            'End If
    '            'Next
    '            'Next
    '            'Next
    '            'Next
    '            'Next
    '            'Next
    '        Next
    '    Else
    '        If MessageBox.Show("Stockdata did not return anything. Now it will run on all internal screener stock at a time. Do you want to continue?", "Backtest", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
    '            'For timeframe As Integer = 1 To 5
    '            Dim strategyStockType As Trade.TypeOfStock = Trade.TypeOfStock.Futures

    '            Dim backtestStrategy As GenericStrategy = New GenericStrategy(canceller:=cts,
    '                                                                          tickSize:=0.05,
    '                                                                          eodExitTime:=TimeSpan.Parse("15:15:00"),
    '                                                                          lastTradeEntryTime:=TimeSpan.Parse("14:30:00"),
    '                                                                          exchangeStartTime:=TimeSpan.Parse("09:15:00"),
    '                                                                          exchangeEndTime:=TimeSpan.Parse("15:29:00"),
    '                                                                          signalTimeFrame:=5,
    '                                                                          UseHeikenAshi:=False,
    '                                                                          associatedStockData:=Nothing,
    '                                                                          stockType:=strategyStockType)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            For file As Integer = 1 To 1 Step 1
    '                For firstTradeMultiplier As Integer = 3 To 3 Step 1
    '                    For targetBuffer As Integer = 0 To 0 Step 1
    '                        If file = 1 Then
    '                            'backtestStrategy.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Cash Stock List ATR Based.csv")
    '                            backtestStrategy.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Future Stock List ATR Based.csv")
    '                            'backtestStrategy.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "BANKNIFTY Future Stock List.csv")
    '                        ElseIf file = 2 Then
    '                            'backtestStrategy.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Cash Stock List Cap Based.csv")
    '                            backtestStrategy.StockFileName = Path.Combine(My.Application.Info.DirectoryPath, "Future Stock List Cap Based.csv")
    '                        Else
    '                            Throw New ApplicationException("No input file found")
    '                        End If
    '                        backtestStrategy.QuantityFlag = 1
    '                        backtestStrategy.MaxStoplossAmount = 100
    '                        backtestStrategy.FirstTradeTargetMultiplier = firstTradeMultiplier
    '                        backtestStrategy.TargetMultiplierWithBuffer = targetBuffer
    '                        backtestStrategy.ForwardTradeTargetMultiplier = 3
    '                        backtestStrategy.CapitalToBeUsed = 20000

    '                        Select Case strategyStockType
    '                            Case Trade.TypeOfStock.Cash
    '                                Strategy.MarginMultiplier = 13
    '                            Case Trade.TypeOfStock.Commodity
    '                                Strategy.MarginMultiplier = 70
    '                            Case Trade.TypeOfStock.Currency
    '                                Strategy.MarginMultiplier = 70
    '                            Case Trade.TypeOfStock.Futures
    '                                Strategy.MarginMultiplier = 30
    '                        End Select
    '                        backtestStrategy.NumberOfTradeableStockPerDay = 5
    '                        backtestStrategy.NumberOfTradePerDay = Integer.MaxValue
    '                        backtestStrategy.NumberOfTradePerStockPerDay = 3
    '                        backtestStrategy.TrailingSL = False
    '                        backtestStrategy.ReverseSignalTrade = True
    '                        backtestStrategy.ModifyTarget = False
    '                        backtestStrategy.ModifyStoploss = False
    '                        backtestStrategy.ExitOnOverAllFixedTargetStoploss = False
    '                        backtestStrategy.OverAllProfitPerDay = 8500
    '                        backtestStrategy.OverAllLossPerDay = -800000
    '                        backtestStrategy.ExitOnStockFixedTargetStoploss = False
    '                        backtestStrategy.StockMaxProfitPerDay = 8500
    '                        backtestStrategy.StockMaxLossPerDay = -500000
    '                        Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '                    Next
    '                Next
    '            Next
    '        End If
    '    End If
    'End Function
#End Region

#Region "Other Strategy"
    ''Private Async Function FractalMACandleRetracementStrategyAsync(startDate As Date, endDate As Date) As Task
    ''    Dim backtestStrategy As FractalMACandleRetrace = New FractalMACandleRetrace(cts,
    ''                                                                                0.05,
    ''                                                                                TimeSpan.Parse("15:15:00"),
    ''                                                                                TimeSpan.Parse("15:00:00"),
    ''                                                                                TimeSpan.Parse("09:15:00"),
    ''                                                                                TimeSpan.Parse("15:30:00"),
    ''                                                                                50)
    ''    AddHandler backtestStrategy.Heartbeat, AddressOf test_Heartbeat
    ''    backtestStrategy.SpikeChangePercentageOfStockForPreMarketVolumeSpikeScreener = 100
    ''    backtestStrategy.NumberOfDaysForNRxScreener = 5
    ''    backtestStrategy.NumberOfTradeableStock = 1
    ''    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    ''End Function
    'Private Async Function NaughtyBoyStockVWAPStrategyAsync(startDate As Date, endDate As Date) As Task
    '    Dim backtestStrategy As NaughtyBoyVWAP = New NaughtyBoyVWAP(_cts,
    '                                                                0.05,
    '                                                                TimeSpan.Parse("15:15:00"),
    '                                                                TimeSpan.Parse("15:00:00"),
    '                                                                TimeSpan.Parse("09:15:00"),
    '                                                                TimeSpan.Parse("15:30:00"),
    '                                                                1,
    '                                                                2,
    '                                                                5)

    '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '    backtestStrategy.PerTradeMaxProfitPercentage = 1
    '    backtestStrategy.PerTradeMaxLossPercentage = -0.3
    '    backtestStrategy.PerStockMaxProfitPercentage = 1.5
    '    backtestStrategy.PerStockMaxLossPercentage = -0.5
    '    backtestStrategy.OverAllProfitPerDay = 16000
    '    backtestStrategy.OverAllLossPerDay = -7000
    '    backtestStrategy.NumberOfTradeableStockPerDay = 4
    '    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    'End Function
    'Private Async Function InTheTrendStrategyAsync(startDate As Date, endDate As Date) As Task
    '    Dim backtestStrategy As InTheTrend = New InTheTrend(_cts,
    '                                                        0.05,
    '                                                        TimeSpan.Parse("15:15:00"),
    '                                                        TimeSpan.Parse("15:00:00"),
    '                                                        TimeSpan.Parse("09:15:00"),
    '                                                        TimeSpan.Parse("15:30:00"),
    '                                                        1,
    '                                                        1,
    '                                                        5)

    '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '    backtestStrategy.PerTradeMaxProfitPercentage = 1
    '    backtestStrategy.PerTradeMaxLossPercentage = -1
    '    backtestStrategy.PerStockMaxProfitPercentage = 1
    '    backtestStrategy.PerStockMaxLossPercentage = -1
    '    backtestStrategy.OverAllProfitPerDay = 20000
    '    backtestStrategy.OverAllLossPerDay = -10000
    '    backtestStrategy.NumberOfTradeableStockPerDay = 6
    '    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    'End Function
    'Private Async Function GapFillStrategyAsync(startDate As Date, endDate As Date) As Task
    '    For numberOfStock As Integer = 5 To 8
    '        For useContinuousTrailingSL As Integer = 0 To 1
    '            Dim backtestStrategy As GapFill = New GapFill(_cts,
    '                                                    0.05,
    '                                                    TimeSpan.Parse("15:15:00"),
    '                                                    TimeSpan.Parse("15:00:00"),
    '                                                    TimeSpan.Parse("09:15:00"),
    '                                                    TimeSpan.Parse("15:30:00"),
    '                                                    0.001,
    '                                                    0.01,
    '                                                    useContinuousTrailingSL,
    '                                                    False)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            backtestStrategy.PerTradeMaxProfitPercentage = 20
    '            backtestStrategy.PerTradeMaxLossPercentage = -0.5
    '            backtestStrategy.PerStockMaxProfitPercentage = 20
    '            backtestStrategy.PerStockMaxLossPercentage = -0.5
    '            backtestStrategy.NumberOfTradeableStockPerDay = numberOfStock
    '            Strategy.MarginMultiplier = 30
    '            backtestStrategy.OverAllCapital = 150000
    '            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '        Next
    '    Next
    'End Function
    'Private Async Function GapFillWithoutFillingPCStrategyAsync(startDate As Date, endDate As Date) As Task
    '    For numberOfStock As Integer = 5 To 8
    '        For useContinuousTrailingSL As Integer = 0 To 1
    '            Dim backtestStrategy As GapFillWithoutFillingPC = New GapFillWithoutFillingPC(_cts,
    '                                                                                        0.05,
    '                                                                                        TimeSpan.Parse("15:15:00"),
    '                                                                                        TimeSpan.Parse("15:00:00"),
    '                                                                                        TimeSpan.Parse("09:15:00"),
    '                                                                                        TimeSpan.Parse("15:30:00"),
    '                                                                                        0.001,
    '                                                                                        0.01,
    '                                                                                        useContinuousTrailingSL,
    '                                                                                        False)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            backtestStrategy.PerTradeMaxProfitPercentage = 20
    '            backtestStrategy.PerTradeMaxLossPercentage = -0.5
    '            backtestStrategy.PerStockMaxProfitPercentage = 20
    '            backtestStrategy.PerStockMaxLossPercentage = -0.5
    '            backtestStrategy.NumberOfTradeableStockPerDay = numberOfStock
    '            Strategy.MarginMultiplier = 30
    '            backtestStrategy.OverAllCapital = 150000
    '            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '        Next
    '    Next
    'End Function
    'Private Async Function GapFillWithoutFillingNeglectGnGStrategyAsync(startDate As Date, endDate As Date) As Task
    '    For numberOfStock As Integer = 5 To 8
    '        For useContinuousTrailingSL As Integer = 0 To 1
    '            Dim backtestStrategy As GapFillWithoutFillingPCNeglectGnG = New GapFillWithoutFillingPCNeglectGnG(_cts,
    '                                                                                                    0.05,
    '                                                                                                    TimeSpan.Parse("15:15:00"),
    '                                                                                                    TimeSpan.Parse("15:00:00"),
    '                                                                                                    TimeSpan.Parse("09:15:00"),
    '                                                                                                    TimeSpan.Parse("15:30:00"),
    '                                                                                                    0.001,
    '                                                                                                    0.01,
    '                                                                                                    useContinuousTrailingSL,
    '                                                                                                    False)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            backtestStrategy.PerTradeMaxProfitPercentage = 20
    '            backtestStrategy.PerTradeMaxLossPercentage = -0.5
    '            backtestStrategy.PerStockMaxProfitPercentage = 20
    '            backtestStrategy.PerStockMaxLossPercentage = -0.5
    '            backtestStrategy.NumberOfTradeableStockPerDay = numberOfStock
    '            Strategy.MarginMultiplier = 30
    '            backtestStrategy.OverAllCapital = 150000
    '            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '        Next
    '    Next
    'End Function
    'Private Async Function GapFillWithoutFillingPHLCStrategyAsync(startDate As Date, endDate As Date) As Task
    '    For numberOfStock As Integer = 5 To 8
    '        For useContinuousTrailingSL As Integer = 0 To 1
    '            Dim backtestStrategy As GapFillWithoutFillingPHLC = New GapFillWithoutFillingPHLC(_cts,
    '                                                                                        0.05,
    '                                                                                        TimeSpan.Parse("15:15:00"),
    '                                                                                        TimeSpan.Parse("15:00:00"),
    '                                                                                        TimeSpan.Parse("09:15:00"),
    '                                                                                        TimeSpan.Parse("15:30:00"),
    '                                                                                        0.001,
    '                                                                                        0.01,
    '                                                                                        useContinuousTrailingSL,
    '                                                                                        False)

    '            AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '            backtestStrategy.PerTradeMaxProfitPercentage = 20
    '            backtestStrategy.PerTradeMaxLossPercentage = -0.5
    '            backtestStrategy.PerStockMaxProfitPercentage = 20
    '            backtestStrategy.PerStockMaxLossPercentage = -0.5
    '            backtestStrategy.NumberOfTradeableStockPerDay = numberOfStock
    '            Strategy.MarginMultiplier = 30
    '            backtestStrategy.OverAllCapital = 150000
    '            Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    '        Next
    '    Next
    'End Function
    'Private Async Function PairHazingStrategyAsync(startDate As Date, endDate As Date) As Task
    '    Dim backtestStrategy As SpreadStrategy = New SpreadStrategy(canceller:=_cts,
    '                                                                  tickSize:=0.05,
    '                                                                  eodExitTime:=TimeSpan.Parse("15:15:00"),
    '                                                                  lastTradeEntryTime:=TimeSpan.Parse("14:30:00"),
    '                                                                  exchangeStartTime:=TimeSpan.Parse("09:15:00"),
    '                                                                  exchangeEndTime:=TimeSpan.Parse("15:29:00"),
    '                                                                  signalTimeFrame:=1,
    '                                                                  UseHeikenAshi:=False,
    '                                                                  associatedStockData:=Nothing)

    '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '    Strategy.MarginMultiplier = 70
    '    backtestStrategy.NumberOfTradePerDay = Integer.MaxValue
    '    backtestStrategy.NumberOfTradePerStockPerDay = 2
    '    backtestStrategy.TrailingSL = False
    '    backtestStrategy.ReverseSignalTrade = True
    '    backtestStrategy.ModifyTargetStoploss = False
    '    backtestStrategy.ExitOnStockFixedTargetStoploss = False
    '    backtestStrategy.StockMaxProfitPerDay = 4500
    '    backtestStrategy.StockMaxLossPerDay = -9000
    '    backtestStrategy.ExitOnOverAllFixedTargetStoploss = False
    '    backtestStrategy.OverAllProfitPerDay = 3000
    '    backtestStrategy.OverAllLossPerDay = -10000
    '    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    'End Function
    'Private Async Function MarutiStrategyAsync(startDate As Date, endDate As Date) As Task
    '    Dim backtestStrategy As MarutiStrategy = New MarutiStrategy(canceller:=_cts,
    '                                                                  tickSize:=0.05,
    '                                                                  eodExitTime:=TimeSpan.Parse("15:15:00"),
    '                                                                  lastTradeEntryTime:=TimeSpan.Parse("15:09:00"),
    '                                                                  exchangeStartTime:=TimeSpan.Parse("9:15:00"),
    '                                                                  exchangeEndTime:=TimeSpan.Parse("15:29:00"))

    '    AddHandler backtestStrategy.Heartbeat, AddressOf OnHeartbeat
    '    Strategy.MarginMultiplier = 13
    '    Await backtestStrategy.TestStrategyAsync(startDate, endDate).ConfigureAwait(False)
    'End Function
#End Region

#Region "Procedure"
    Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
        AddHandler _cmn.Heartbeat, AddressOf OnHeartbeat
        Dim dt As DataTable = Nothing
        Dim conn As MySqlConnection = _cmn.OpenDBConnection
        Dim ret As Dictionary(Of String, Decimal()) = Nothing

        Dim signalCheckingDate As Date = _cmn.GetPreviousTradingDay(Common.DataBaseTable.EOD_Cash, tradingDate.Date)
        If conn.State = ConnectionState.Open Then
            OnHeartbeat("Fetching All Stock Data")
            Dim cmd As New MySqlCommand("GET_STOCK_CASH_DATA_ATR_VOLUME_ALL_DATES", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@startDate", signalCheckingDate)
            cmd.Parameters.AddWithValue("@endDate", signalCheckingDate)
            cmd.Parameters.AddWithValue("@numberOfRecords", 0)
            cmd.Parameters.AddWithValue("@minClose", 100)
            cmd.Parameters.AddWithValue("@maxClose", 1500)
            cmd.Parameters.AddWithValue("@atrPercentage", 2.5)
            cmd.Parameters.AddWithValue("@potentialAmount", 500000)

            Dim adapter As New MySqlDataAdapter(cmd)
            adapter.SelectCommand.CommandTimeout = 3000
            dt = New DataTable
            adapter.Fill(dt)
        End If
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                If Not IsDBNull(dt.Rows(i).Item(5)) Then
                    If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
                    ret.Add(dt.Rows(i).Item(1), {dt.Rows(i).Item(5), dt.Rows(i).Item(5)})
                End If
            Next
        End If
        Return ret
    End Function
#End Region

End Class