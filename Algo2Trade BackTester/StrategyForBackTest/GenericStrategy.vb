Imports Algo2TradeBLL
Imports Utilities.Numbers
Imports System.Threading
Imports Utilities.Time
Imports System.IO

Public Class GenericStrategy
    Inherits Strategy
    Implements IDisposable

    Private ReadOnly _SignalTimeFrame As Integer
    Private ReadOnly _UseHeikenAshi As Boolean
    Public Property PartialExit As Boolean = False
    Public Property NumberOfTradePerStockPerDay As Integer = Integer.MaxValue
    Public Property NumberOfTradePerDay As Integer = Integer.MaxValue
    Public Property CountTradesWithBreakevenMovement As Boolean = False
    Public Property TrailingSL As Boolean = False
    Public Property ReverseSignalTrade As Boolean = False
    Public Property ExitOnStockFixedTargetStoploss As Boolean = False
    Public Property StockMaxProfitPerDay As Double = 15000
    Public Property StockMaxLossPerDay As Double = -10000
    Public Property ExitOnOverAllFixedTargetStoploss As Boolean = False
    Public Property ModifyTarget As Boolean = False
    Public Property ModifyStoploss As Boolean = False
    Public Property SameDirectionTrade As Boolean = False
    Public Property NIFTY50Stocks As String()

    'For ATR Based Candle Range Strategy
    Public QuantityFlag As Integer = 1
    Public MaxStoplossAmount As Decimal = 100
    Public StockFileName As String = Nothing
    Public FirstTradeTargetMultiplier As Decimal = 2
    Public EarlyStoploss As Boolean = False
    Public ForwardTradeTargetMultiplier As Decimal = 3
    Public CapitalToBeUsed As Decimal = 20000
    Public CandleBasedEntry As Boolean = False
    'End
    ''For JOYMA_ATM
    'Public MinimumWickPercentageOfCandleRange = 50
    'Public MinimumWickPercentageOutsideATRBand = 50
    ''end

    Private ReadOnly _StockData As Dictionary(Of String, Decimal())
    Private ReadOnly _Nifty50FilePath As String = Path.Combine(My.Application.Info.DirectoryPath, "NIFTY50 Stocks.txt")
    Private ReadOnly _StockType As Trade.TypeOfStock
    Private ReadOnly _DatabaseTable As Common.DataBaseTable
    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal tickSize As Double,
                   ByVal eodExitTime As TimeSpan,
                   ByVal lastTradeEntryTime As TimeSpan,
                   ByVal exchangeStartTime As TimeSpan,
                   ByVal exchangeEndTime As TimeSpan,
                   ByVal signalTimeFrame As Integer,
                   ByVal UseHeikenAshi As Boolean,
                   ByVal associatedStockData As Dictionary(Of String, Decimal()),
                   ByVal stockType As Trade.TypeOfStock)
        MyBase.New(canceller, tickSize, eodExitTime, lastTradeEntryTime, exchangeStartTime, exchangeEndTime)
        Me._SignalTimeFrame = signalTimeFrame
        Me._UseHeikenAshi = UseHeikenAshi
        Me._StockData = associatedStockData
        Me.NIFTY50Stocks = File.ReadAllLines(_Nifty50FilePath)
        Me._StockType = stockType
        Select Case Me._StockType
            Case Trade.TypeOfStock.Cash
                Me._DatabaseTable = Common.DataBaseTable.Intraday_Cash
            Case Trade.TypeOfStock.Commodity
                Me._DatabaseTable = Common.DataBaseTable.Intraday_Commodity
            Case Trade.TypeOfStock.Currency
                Me._DatabaseTable = Common.DataBaseTable.Intraday_Currency
            Case Trade.TypeOfStock.Futures
                Me._DatabaseTable = Common.DataBaseTable.Intraday_Futures
        End Select
    End Sub

#Region "Test Strategy Normal"
    Public Overrides Async Function TestStrategyAsync(startDate As Date, endDate As Date) As Task
        Await Task.Delay(0).ConfigureAwait(False)
        Dim filename As String = String.Format("TF {0},Trlg {1},Samesd {2},CtBrkevn {3},ML {4},Nmb {5},1TgMul {6}",
                                               Me._SignalTimeFrame,
                                               Me.TrailingSL,
                                               Me.SameDirectionTrade,
                                               Me.CountTradesWithBreakevenMovement,
                                               Me.OverAllLossPerDay,
                                               Me.NumberOfTradePerStockPerDay,
                                               Me.FirstTradeTargetMultiplier)

        Dim tradesFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Trades.a2t", filename))
        Dim capitalFileName As String = Path.Combine(My.Application.Info.DirectoryPath, String.Format("{0}.Capital.a2t", filename))

        If File.Exists(tradesFileName) AndAlso File.Exists(capitalFileName) Then
            Dim folderpath As String = Path.Combine(My.Application.Info.DirectoryPath, "BackTest Output")
            Dim files() = Directory.GetFiles(folderpath, "*.xlsx")
            For Each file In files
                If file.ToUpper.Contains(filename.ToUpper) Then
                    Exit Function
                End If
            Next
            PrintArrayToExcel(filename, tradesFileName, capitalFileName)
        Else
            If File.Exists(tradesFileName) Then File.Delete(tradesFileName)
            If File.Exists(capitalFileName) Then File.Delete(capitalFileName)
            'TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

            Dim totalPL As Decimal = 0
            Dim tradeCheckingDate As Date = startDate.Date
            While tradeCheckingDate <= endDate.Date
                TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))
                Dim stockList As Dictionary(Of String, Decimal()) = Nothing
                'Dim tempStockList As IEnumerable(Of KeyValuePair(Of String, Integer)) = GetStockData(tradeCheckingDate).Take(1)
                'stockList = tempStockList.ToDictionary(Of String, Integer)(Function(x)
                '                                                               Return x.Key
                '                                                           End Function, Function(y)
                '                                                                             Return y.Value
                '                                                                         End Function)
                If Me._StockData IsNot Nothing AndAlso Me._StockData.Count > 0 Then
                    stockList = New Dictionary(Of String, Decimal())
                    Dim lotSize As Integer = _common.GetAppropiateLotSize(_DatabaseTable, _StockData.FirstOrDefault.Key, tradeCheckingDate.Date)
                    stockList.Add(_StockData.FirstOrDefault.Key, {lotSize, lotSize})
                Else
                    stockList = GetStockData(tradeCheckingDate)
                End If
                'stockList = New Dictionary(Of String, Decimal())
                'stockList.Add("StockName", {Quantity, LotSize})
                'stockList.Add("CRUDEOIL", {1000, 100})

                If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                    Dim currentDayOneMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                    Dim XDayXMinuteStocksPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                    Dim XDayRuleSignalStocksPayload As Dictionary(Of String, Dictionary(Of Date, EntryDetails)) = Nothing
                    'Dim XDayRuleEntryStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    'Dim XDayRuleTargetStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    'Dim XDayRuleStoplossStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    'Dim XDayRuleQuantityStocksPayload As Dictionary(Of String, Dictionary(Of Date, Integer)) = Nothing

                    Dim XDayRuleModifyStoplossStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim XDayRuleModifyTargetStocksPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                    Dim XDayRuleSupporting1StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting2StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting3StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting4StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting5StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting6StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting7StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting8StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting9StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing
                    Dim XDayRuleSupporting10StocksPayload As Dictionary(Of String, Dictionary(Of Date, String)) = Nothing

                    'First lets build the payload for all the stocks
                    Dim stockCount As Integer = 0
                    'Dim stockData As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                    For Each stock In stockList.Keys
                        stockCount += 1
                        Dim XDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        Dim XDayXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        Dim XDayXMinuteHAPayload As Dictionary(Of Date, Payload) = Nothing
                        Dim currentDayOneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                        Dim XDayRuleSignalPayload As Dictionary(Of Date, EntryDetails) = Nothing
                        'Dim XDayRuleEntryPayload As Dictionary(Of Date, Decimal) = Nothing
                        'Dim XDayRuleTargetPayload As Dictionary(Of Date, Decimal) = Nothing
                        'Dim XDayRuleStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
                        'Dim XDayRuleQuantityPayload As Dictionary(Of Date, Integer) = Nothing

                        Dim XDayRuleModifyStoplossPayload As Dictionary(Of Date, Decimal) = Nothing
                        Dim XDayRuleModifyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
                        Dim XDayRuleSupporting1Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting2Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting3Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting4Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting5Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting6Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting7Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting8Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting9Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleSupporting10Payload As Dictionary(Of Date, String) = Nothing
                        Dim XDayRuleOutputPayload As Dictionary(Of String, Object) = Nothing
                        'Get payload
                        If Data.PastIntradayData IsNot Nothing AndAlso Data.PastIntradayData.Count > 0 AndAlso
                            Data.PastIntradayData.ContainsKey(tradeCheckingDate) AndAlso Data.PastIntradayData(tradeCheckingDate).ContainsKey(stock) Then
                            XDayOneMinutePayload = Data.PastIntradayData(tradeCheckingDate)(stock)
                        Else
                            If Me.DataSource = SourceOfData.Database Then
                                XDayOneMinutePayload = _common.GetRawPayload(_DatabaseTable, stock, tradeCheckingDate.AddDays(-7), tradeCheckingDate)
                            ElseIf Me.DataSource = SourceOfData.Live Then
                                XDayOneMinutePayload = Await _common.GetHistoricalData(_DatabaseTable, stock, tradeCheckingDate).ConfigureAwait(False)
                            End If
                            'If XDayOneMinutePayload IsNot Nothing Then
                            '    If stockData Is Nothing Then stockData = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            '    stockData.Add(stock, XDayOneMinutePayload)
                            'End If
                        End If

                        'Now transfer only the current date payload into the workable payload (this will be used for the main loop and checking if the date is a valid date)
                        If XDayOneMinutePayload IsNot Nothing AndAlso XDayOneMinutePayload.Count > 0 Then
                            OnHeartbeat(String.Format("Processing for {0} on {1}. Stock Counter: [ {2}/{3} ]", stock, tradeCheckingDate.ToShortDateString, stockCount, stockList.Count))
                            For Each runningPayload In XDayOneMinutePayload.Keys
                                If runningPayload.Date = tradeCheckingDate.Date Then
                                    If currentDayOneMinutePayload Is Nothing Then currentDayOneMinutePayload = New Dictionary(Of Date, Payload)
                                    currentDayOneMinutePayload.Add(runningPayload, XDayOneMinutePayload(runningPayload))
                                End If
                            Next
                            'Add all these payloads into the stock collections
                            If currentDayOneMinutePayload IsNot Nothing AndAlso currentDayOneMinutePayload.Count > 0 Then
                                If _SignalTimeFrame > 1 Then
                                    XDayXMinutePayload = _common.ConvertPayloadsToXMinutes(XDayOneMinutePayload, _SignalTimeFrame)
                                Else
                                    XDayXMinutePayload = XDayOneMinutePayload
                                End If
                                If _UseHeikenAshi Then
                                    Indicator.HeikenAshi.ConvertToHeikenAshi(XDayXMinutePayload, XDayXMinuteHAPayload)
                                Else
                                    XDayXMinuteHAPayload = XDayXMinutePayload
                                End If

                                If XDayXMinuteHAPayload IsNot Nothing AndAlso XDayXMinuteHAPayload.Count > 0 Then
                                    'TODO: Change
                                    Using strategyBaseRule As New ATRBasedCandleRangeStrategyRule(XDayXMinuteHAPayload, TickSize, stockList(stock)(0), _canceller, _common, tradeCheckingDate, _SignalTimeFrame, stockList(stock)(2), stockList(stock)(3), _StockType)
                                        strategyBaseRule.CandleBasedEntry = Me.CandleBasedEntry
                                        strategyBaseRule.QuantityFlag = Me.QuantityFlag
                                        strategyBaseRule.MaxStoplossAmount = Me.MaxStoplossAmount
                                        strategyBaseRule.FirstTradeTargetMultiplier = Me.FirstTradeTargetMultiplier
                                        strategyBaseRule.EarlyStoploss = Me.EarlyStoploss
                                        strategyBaseRule.ForwardTradeTargetMultiplier = Me.ForwardTradeTargetMultiplier
                                        strategyBaseRule.CapitalToBeUsed = Me.CapitalToBeUsed
                                        strategyBaseRule.CalculateRule(XDayRuleOutputPayload)
                                    End Using
                                    'Using strategyBaseRule As New OpenATRStrategyRule(XDayXMinuteHAPayload, TickSize, stockList(stock)(0), Canceller, Cmn, tradeCheckingDate, _SignalTimeFrame, stockList(stock)(2), stockList(stock)(3), _StockType)
                                    '    strategyBaseRule.CandleBasedEntry = Me.CandleBasedEntry
                                    '    strategyBaseRule.QuantityFlag = Me.QuantityFlag
                                    '    strategyBaseRule.MaxStoplossAmount = Me.MaxStoplossAmount
                                    '    strategyBaseRule.FirstTradeTargetMultiplier = Me.FirstTradeTargetMultiplier
                                    '    strategyBaseRule.EarlyStoploss = Me.EarlyStoploss
                                    '    strategyBaseRule.ForwardTradeTargetMultiplier = Me.ForwardTradeTargetMultiplier
                                    '    strategyBaseRule.CapitalToBeUsed = Me.CapitalToBeUsed
                                    '    strategyBaseRule.CalculateRule(XDayRuleOutputPayload)
                                    'End Using
                                End If
                                If XDayRuleOutputPayload IsNot Nothing Then
                                    If XDayRuleOutputPayload.ContainsKey("Signal") Then XDayRuleSignalPayload = CType(XDayRuleOutputPayload("Signal"), Dictionary(Of Date, EntryDetails))
                                    'If XDayRuleOutputPayload.ContainsKey("Entry") Then XDayRuleEntryPayload = CType(XDayRuleOutputPayload("Entry"), Dictionary(Of Date, Decimal))
                                    'If XDayRuleOutputPayload.ContainsKey("Target") Then XDayRuleTargetPayload = CType(XDayRuleOutputPayload("Target"), Dictionary(Of Date, Decimal))
                                    'If XDayRuleOutputPayload.ContainsKey("Stoploss") Then XDayRuleStoplossPayload = CType(XDayRuleOutputPayload("Stoploss"), Dictionary(Of Date, Decimal))
                                    'If XDayRuleOutputPayload.ContainsKey("Quantity") Then XDayRuleQuantityPayload = CType(XDayRuleOutputPayload("Quantity"), Dictionary(Of Date, Integer))

                                    If XDayRuleOutputPayload.ContainsKey("ModifyStoploss") Then XDayRuleModifyStoplossPayload = CType(XDayRuleOutputPayload("ModifyStoploss"), Dictionary(Of Date, Decimal))
                                    If XDayRuleOutputPayload.ContainsKey("ModifyTarget") Then XDayRuleModifyTargetPayload = CType(XDayRuleOutputPayload("ModifyTarget"), Dictionary(Of Date, Decimal))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting1") Then XDayRuleSupporting1Payload = CType(XDayRuleOutputPayload("Supporting1"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting2") Then XDayRuleSupporting2Payload = CType(XDayRuleOutputPayload("Supporting2"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting3") Then XDayRuleSupporting3Payload = CType(XDayRuleOutputPayload("Supporting3"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting4") Then XDayRuleSupporting4Payload = CType(XDayRuleOutputPayload("Supporting4"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting5") Then XDayRuleSupporting5Payload = CType(XDayRuleOutputPayload("Supporting5"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting6") Then XDayRuleSupporting6Payload = CType(XDayRuleOutputPayload("Supporting6"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting7") Then XDayRuleSupporting7Payload = CType(XDayRuleOutputPayload("Supporting7"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting8") Then XDayRuleSupporting8Payload = CType(XDayRuleOutputPayload("Supporting8"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting9") Then XDayRuleSupporting9Payload = CType(XDayRuleOutputPayload("Supporting9"), Dictionary(Of Date, String))
                                    If XDayRuleOutputPayload.ContainsKey("Supporting10") Then XDayRuleSupporting10Payload = CType(XDayRuleOutputPayload("Supporting10"), Dictionary(Of Date, String))
                                End If
                                If currentDayOneMinuteStocksPayload Is Nothing Then currentDayOneMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                currentDayOneMinuteStocksPayload.Add(stock, currentDayOneMinutePayload)
                                If XDayXMinuteStocksPayload Is Nothing Then XDayXMinuteStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                                XDayXMinuteStocksPayload.Add(stock, XDayXMinutePayload)
                                If XDayRuleSignalStocksPayload Is Nothing Then XDayRuleSignalStocksPayload = New Dictionary(Of String, Dictionary(Of Date, EntryDetails))
                                XDayRuleSignalStocksPayload.Add(stock, XDayRuleSignalPayload)
                                'If XDayRuleEntryStocksPayload Is Nothing Then XDayRuleEntryStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                'XDayRuleEntryStocksPayload.Add(stock, XDayRuleEntryPayload)
                                'If XDayRuleTargetStocksPayload Is Nothing Then XDayRuleTargetStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                'XDayRuleTargetStocksPayload.Add(stock, XDayRuleTargetPayload)
                                'If XDayRuleStoplossStocksPayload Is Nothing Then XDayRuleStoplossStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                'XDayRuleStoplossStocksPayload.Add(stock, XDayRuleStoplossPayload)
                                'If XDayRuleQuantityStocksPayload Is Nothing Then XDayRuleQuantityStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Integer))
                                'XDayRuleQuantityStocksPayload.Add(stock, XDayRuleQuantityPayload)

                                If XDayRuleModifyStoplossStocksPayload Is Nothing Then XDayRuleModifyStoplossStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                XDayRuleModifyStoplossStocksPayload.Add(stock, XDayRuleModifyStoplossPayload)
                                If XDayRuleModifyTargetStocksPayload Is Nothing Then XDayRuleModifyTargetStocksPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                                XDayRuleModifyTargetStocksPayload.Add(stock, XDayRuleModifyTargetPayload)
                                If XDayRuleSupporting1StocksPayload Is Nothing Then XDayRuleSupporting1StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting1StocksPayload.Add(stock, XDayRuleSupporting1Payload)
                                If XDayRuleSupporting2StocksPayload Is Nothing Then XDayRuleSupporting2StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting2StocksPayload.Add(stock, XDayRuleSupporting2Payload)
                                If XDayRuleSupporting3StocksPayload Is Nothing Then XDayRuleSupporting3StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting3StocksPayload.Add(stock, XDayRuleSupporting3Payload)
                                If XDayRuleSupporting4StocksPayload Is Nothing Then XDayRuleSupporting4StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting4StocksPayload.Add(stock, XDayRuleSupporting4Payload)
                                If XDayRuleSupporting5StocksPayload Is Nothing Then XDayRuleSupporting5StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting5StocksPayload.Add(stock, XDayRuleSupporting5Payload)
                                If XDayRuleSupporting6StocksPayload Is Nothing Then XDayRuleSupporting6StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting6StocksPayload.Add(stock, XDayRuleSupporting6Payload)
                                If XDayRuleSupporting7StocksPayload Is Nothing Then XDayRuleSupporting7StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting7StocksPayload.Add(stock, XDayRuleSupporting7Payload)
                                If XDayRuleSupporting8StocksPayload Is Nothing Then XDayRuleSupporting8StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting8StocksPayload.Add(stock, XDayRuleSupporting8Payload)
                                If XDayRuleSupporting9StocksPayload Is Nothing Then XDayRuleSupporting9StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting9StocksPayload.Add(stock, XDayRuleSupporting9Payload)
                                If XDayRuleSupporting10StocksPayload Is Nothing Then XDayRuleSupporting10StocksPayload = New Dictionary(Of String, Dictionary(Of Date, String))
                                XDayRuleSupporting10StocksPayload.Add(stock, XDayRuleSupporting10Payload)
                            End If
                        End If

#Region "Cleanup"
                        ''If XDayOneMinutePayload IsNot Nothing Then XDayOneMinutePayload.Clear()
                        ''If XDayXMinutePayload IsNot Nothing Then XDayXMinutePayload.Clear()
                        ''If XDayXMinuteHAPayload IsNot Nothing Then XDayXMinuteHAPayload.Clear()
                        ''If currentDayOneMinutePayload IsNot Nothing Then currentDayOneMinutePayload.Clear()
                        ''If XDayRuleSignalPayload IsNot Nothing Then XDayRuleSignalPayload.Clear()
                        ''XDayOneMinutePayload = Nothing
                        ''XDayXMinutePayload = Nothing
                        ''XDayXMinuteHAPayload = Nothing
                        ''currentDayOneMinutePayload = Nothing
                        ''XDayRuleSignalPayload = Nothing

                        'If XDayRuleModifyStoplossPayload IsNot Nothing Then XDayRuleModifyStoplossPayload.Clear()
                        'If XDayRuleModifyTargetPayload IsNot Nothing Then XDayRuleModifyTargetPayload.Clear()
                        'If XDayRuleSupporting1Payload IsNot Nothing Then XDayRuleSupporting1Payload.Clear()
                        'If XDayRuleSupporting2Payload IsNot Nothing Then XDayRuleSupporting2Payload.Clear()
                        'If XDayRuleSupporting3Payload IsNot Nothing Then XDayRuleSupporting3Payload.Clear()
                        'If XDayRuleSupporting4Payload IsNot Nothing Then XDayRuleSupporting4Payload.Clear()
                        'If XDayRuleSupporting5Payload IsNot Nothing Then XDayRuleSupporting5Payload.Clear()
                        'If XDayRuleSupporting6Payload IsNot Nothing Then XDayRuleSupporting6Payload.Clear()
                        'If XDayRuleSupporting7Payload IsNot Nothing Then XDayRuleSupporting7Payload.Clear()
                        'If XDayRuleSupporting8Payload IsNot Nothing Then XDayRuleSupporting8Payload.Clear()
                        'If XDayRuleSupporting9Payload IsNot Nothing Then XDayRuleSupporting9Payload.Clear()
                        'If XDayRuleSupporting10Payload IsNot Nothing Then XDayRuleSupporting10Payload.Clear()
                        'If XDayRuleOutputPayload IsNot Nothing Then XDayRuleOutputPayload.Clear()

                        'XDayRuleModifyStoplossPayload = Nothing
                        'XDayRuleModifyTargetPayload = Nothing
                        'XDayRuleSupporting1Payload = Nothing
                        'XDayRuleSupporting2Payload = Nothing
                        'XDayRuleSupporting3Payload = Nothing
                        'XDayRuleSupporting4Payload = Nothing
                        'XDayRuleSupporting5Payload = Nothing
                        'XDayRuleSupporting6Payload = Nothing
                        'XDayRuleSupporting7Payload = Nothing
                        'XDayRuleSupporting8Payload = Nothing
                        'XDayRuleSupporting9Payload = Nothing
                        'XDayRuleSupporting10Payload = Nothing
                        'XDayRuleOutputPayload = Nothing
#End Region
                    Next
                    'If stockData IsNot Nothing AndAlso stockData.Count > 0 Then
                    '    If Data.PastIntradayData Is Nothing Then Data.PastIntradayData = New Dictionary(Of Date, Dictionary(Of String, Dictionary(Of Date, Payload)))
                    '    Data.PastIntradayData.Add(tradeCheckingDate, stockData)
                    'End If

                    '------------------------------------------------------------------------------------------------------------------------------------------------

                    If currentDayOneMinuteStocksPayload IsNot Nothing AndAlso currentDayOneMinuteStocksPayload.Count > 0 Then
                        OnHeartbeat(String.Format("Checking Trade on {0}", tradeCheckingDate.ToShortDateString))
                        Dim startMinute As TimeSpan = ExchangeStartTime
                        Dim endMinute As TimeSpan = ExchangeEndTime
                        Dim timeChart As Dictionary(Of String, Date) = Nothing
                        While startMinute < endMinute
                            Dim startSecond As TimeSpan = startMinute
                            Dim endSecond As TimeSpan = startMinute.Add(TimeSpan.FromMinutes(_SignalTimeFrame - 1))
                            endSecond = endSecond.Add(TimeSpan.FromSeconds(59))
                            Dim potentialCandleSignalTime As Date = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startMinute.Hours, startMinute.Minutes, startMinute.Seconds)
                            Dim potentialTickSignalTime As Date = Nothing
                            Dim currentMinuteCandlePayload As Payload = Nothing
                            Dim signalCandleTime As Date = Nothing

                            While startSecond <= endSecond
                                potentialTickSignalTime = New Date(tradeCheckingDate.Year, tradeCheckingDate.Month, tradeCheckingDate.Day, startSecond.Hours, startSecond.Minutes, startSecond.Seconds)
                                If potentialTickSignalTime.Second = 0 Then
                                    potentialCandleSignalTime = potentialTickSignalTime
                                End If
                                For Each stockName In stockList.Keys
                                    ''indibar
                                    ''for setting 1% mtm exit for each stock
                                    'Me.StockMaxProfitPerDay = CalculatePL(stockName, stockList(stockName)(2), stockList(stockName)(2) + (stockList(stockName)(2) * 1 / 100), stockList(stockName)(0), stockList(stockName)(1), tradeStockType)
                                    ''end indibar
                                    If IncludeSlippage AndAlso timeChart IsNot Nothing AndAlso timeChart.Count > 0 AndAlso timeChart.ContainsKey(stockName) Then
                                        If potentialTickSignalTime > timeChart(stockName) Then
                                            timeChart.Remove(stockName)
                                        Else
                                            Continue For
                                        End If
                                    End If
                                    Dim runningTrade As Trade = Nothing
                                    Dim runningTrade2 As Trade = Nothing
                                    'Get the current minute candle from the stock collection for this stock for that day
                                    Dim currentSecondTickPayload As List(Of Payload) = Nothing
                                    If currentDayOneMinuteStocksPayload.ContainsKey(stockName) AndAlso currentDayOneMinuteStocksPayload(stockName).ContainsKey(potentialCandleSignalTime) Then
                                        currentMinuteCandlePayload = currentDayOneMinuteStocksPayload(stockName)(potentialCandleSignalTime)
                                    End If
                                    signalCandleTime = potentialTickSignalTime.AddMinutes(-_SignalTimeFrame)
                                    'Now get the ticks for this minute and second
                                    If currentMinuteCandlePayload IsNot Nothing AndAlso currentMinuteCandlePayload.Ticks IsNot Nothing Then
                                        currentSecondTickPayload = currentMinuteCandlePayload.Ticks.FindAll(Function(x)
                                                                                                                Return x.PayloadDate = potentialTickSignalTime
                                                                                                            End Function)
                                    End If

                                    'Main Strategy Logic
                                    Dim finalEntryPrice As Double = Nothing
                                    Dim finalTargetPrice As Double = Nothing
                                    Dim finalTargetRemark As String = Nothing
                                    Dim finalStoplossPrice As Double = Nothing
                                    Dim finalStoplossRemark As String = Nothing
                                    Dim squareOffValue As Double = Nothing
                                    Dim quantity As Integer = Nothing
                                    Dim lotSize As Integer = stockList(stockName)(1)
                                    Dim entryBuffer As Decimal = Nothing
                                    Dim stoplossBuffer As Decimal = Nothing
                                    Dim supporting1 As String = If(XDayRuleSupporting1StocksPayload IsNot Nothing AndAlso XDayRuleSupporting1StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting1StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting1StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting1StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting2 As String = If(XDayRuleSupporting2StocksPayload IsNot Nothing AndAlso XDayRuleSupporting2StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting2StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting2StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting2StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting3 As String = If(XDayRuleSupporting3StocksPayload IsNot Nothing AndAlso XDayRuleSupporting3StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting3StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting3StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting3StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting4 As String = If(XDayRuleSupporting4StocksPayload IsNot Nothing AndAlso XDayRuleSupporting4StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting4StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting4StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting4StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting5 As String = If(XDayRuleSupporting5StocksPayload IsNot Nothing AndAlso XDayRuleSupporting5StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting5StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting5StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting5StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting6 As String = If(XDayRuleSupporting6StocksPayload IsNot Nothing AndAlso XDayRuleSupporting6StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting6StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting6StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting6StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting7 As String = If(XDayRuleSupporting7StocksPayload IsNot Nothing AndAlso XDayRuleSupporting7StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting7StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting7StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting7StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting8 As String = If(XDayRuleSupporting8StocksPayload IsNot Nothing AndAlso XDayRuleSupporting8StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting8StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting8StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting8StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting9 As String = If(XDayRuleSupporting9StocksPayload IsNot Nothing AndAlso XDayRuleSupporting9StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting9StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting9StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting9StocksPayload(stockName)(signalCandleTime), "")
                                    Dim supporting10 As String = If(XDayRuleSupporting10StocksPayload IsNot Nothing AndAlso XDayRuleSupporting10StocksPayload.ContainsKey(stockName) AndAlso XDayRuleSupporting10StocksPayload(stockName) IsNot Nothing AndAlso XDayRuleSupporting10StocksPayload(stockName).ContainsKey(signalCandleTime), XDayRuleSupporting10StocksPayload(stockName)(signalCandleTime), "")

                                    Dim tradeActive As Boolean = False
                                    Dim numberOfExecutedTradePerDay As Integer = 0
                                    Dim numberOfExecutedTradePerStockPerDay As Integer = 0
                                    If currentMinuteCandlePayload IsNot Nothing Then
                                        If Not CountTradesWithBreakevenMovement Then
                                            numberOfExecutedTradePerDay = NumberOfTradesPerDayWithoutBreakevenExit(currentMinuteCandlePayload.PayloadDate)
                                            numberOfExecutedTradePerStockPerDay = NumberOfTradesPerStockPerDayWithoutBreakevenExit(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                                        Else
                                            numberOfExecutedTradePerDay = NumberOfTradesPerDay(currentMinuteCandlePayload.PayloadDate)
                                            numberOfExecutedTradePerStockPerDay = NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                                        End If
                                    End If
                                    'If currentMinuteCandlePayload IsNot Nothing AndAlso numberOfExecutedTradePerDay < NumberOfTradePerDay AndAlso Not _OverallMTMReached AndAlso
                                    '    (_StockMTMReached Is Nothing OrElse (_StockMTMReached IsNot Nothing AndAlso _StockMTMReached.Count > 0 AndAlso
                                    '    _StockMTMReached.ContainsKey(stockName) AndAlso Not _StockMTMReached(stockName))) AndAlso
                                    '    (_StockTargetReached Is Nothing OrElse (_StockTargetReached IsNot Nothing AndAlso _StockTargetReached.Count > 0 AndAlso
                                    '    _StockTargetReached.ContainsKey(stockName) AndAlso Not _StockTargetReached(stockName))) Then
                                    If currentMinuteCandlePayload IsNot Nothing AndAlso numberOfExecutedTradePerDay < NumberOfTradePerDay Then
                                        If currentMinuteCandlePayload IsNot Nothing AndAlso numberOfExecutedTradePerStockPerDay < NumberOfTradePerStockPerDay AndAlso
                                           XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                           XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCandleTime) AndAlso
                                           XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuySignal = 1 Then

                                            'If ReverseSignalTrade Then
                                            '    tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Buy)
                                            'Else
                                            '    tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                            'End If

                                            If Not tradeActive Then
                                                finalEntryPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuyEntry
                                                entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)

                                                finalStoplossPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuyStoploss
                                                stoplossBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)
                                                finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalEntryPrice - finalStoplossPrice, 4))

                                                'If TrailingSL Then
                                                '    finalTargetPrice = finalEntryPrice + 100000
                                                'Else
                                                finalTargetPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuyTarget
                                                'End If
                                                finalTargetRemark = String.Format("Target: {0}", Math.Round(finalTargetPrice - finalEntryPrice, 4))

                                                squareOffValue = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuyTarget - finalEntryPrice
                                                quantity = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).BuyQuantity

                                                Dim tradeTag As String = System.Guid.NewGuid.ToString()

                                                If PartialExit Then
                                                    Dim modifiedQuantity As Integer = quantity / 2
                                                    runningTrade = New Trade(Me,
                                                             currentMinuteCandlePayload.TradingSymbol,
                                                             _StockType,
                                                             currentMinuteCandlePayload.PayloadDate,
                                                             Trade.TradeExecutionDirection.Buy,
                                                             finalEntryPrice,
                                                             entryBuffer,
                                                             Trade.TradeType.MIS,
                                                             Trade.TradeEntryCondition.Original,
                                                             "Signal Candle High Price",
                                                             modifiedQuantity,
                                                             lotSize,
                                                             finalTargetPrice,
                                                             finalTargetRemark,
                                                             finalStoplossPrice,
                                                             stoplossBuffer,
                                                             finalStoplossRemark,
                                                             currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)

                                                    finalTargetPrice = finalEntryPrice + 100
                                                    finalTargetRemark = String.Format("Target: 100")

                                                    runningTrade2 = New Trade(Me,
                                                             currentMinuteCandlePayload.TradingSymbol,
                                                             _StockType,
                                                             currentMinuteCandlePayload.PayloadDate,
                                                             Trade.TradeExecutionDirection.Buy,
                                                             finalEntryPrice,
                                                             entryBuffer,
                                                             Trade.TradeType.MIS,
                                                             Trade.TradeEntryCondition.Original,
                                                             "Signal Candle High Price",
                                                             quantity - modifiedQuantity,
                                                             lotSize,
                                                             finalTargetPrice,
                                                             finalTargetRemark,
                                                             finalStoplossPrice,
                                                             stoplossBuffer,
                                                             finalStoplossRemark,
                                                             currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
                                                Else
                                                    runningTrade = New Trade(Me,
                                                             currentMinuteCandlePayload.TradingSymbol,
                                                             _StockType,
                                                             currentMinuteCandlePayload.PayloadDate,
                                                             Trade.TradeExecutionDirection.Buy,
                                                             finalEntryPrice,
                                                             entryBuffer,
                                                             Trade.TradeType.MIS,
                                                             Trade.TradeEntryCondition.Original,
                                                             "Signal Candle High Price",
                                                             quantity,
                                                             lotSize,
                                                             finalTargetPrice,
                                                             finalTargetRemark,
                                                             finalStoplossPrice,
                                                             stoplossBuffer,
                                                             finalStoplossRemark,
                                                             currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
                                                End If
                                                'Dim potentialOpenTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                                'If potentialOpenTrades IsNot Nothing AndAlso potentialOpenTrades.Count > 0 Then
                                                '    For Each potentialOpenTrade In potentialOpenTrades
                                                '        If potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                '            If PartialExit Then
                                                '                If potentialOpenTrade.PotentialTarget - potentialOpenTrade.EntryPrice <= potentialOpenTrade.SquareOffValue - potentialOpenTrade.EntryPrice Then
                                                '                    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                '                ElseIf potentialOpenTrade.PotentialTarget - potentialOpenTrade.EntryPrice > potentialOpenTrade.SquareOffValue - potentialOpenTrade.EntryPrice Then
                                                '                    If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade2)
                                                '                End If
                                                '            Else
                                                '                'If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                '            End If
                                                '        ElseIf potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                '            'Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                '            'dummyPayload.PayloadDate = potentialTickSignalTime
                                                '            'CancelTrade(potentialOpenTrade, dummyPayload, "Opposite Direction Signal")
                                                '        End If
                                                '    Next
                                                'Else
                                                '    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                '    If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                                'End If
                                                If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                            End If
                                        End If

                                        If currentMinuteCandlePayload IsNot Nothing AndAlso numberOfExecutedTradePerStockPerDay < NumberOfTradePerStockPerDay AndAlso
                                            XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                            XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCandleTime) AndAlso
                                            XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellSignal = -1 Then

                                            'If ReverseSignalTrade Then
                                            '    tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionDirection.Sell)
                                            'Else
                                            '    tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                            'End If

                                            If Not tradeActive Then
                                                finalEntryPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellEntry
                                                entryBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)

                                                finalStoplossPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellStoploss
                                                stoplossBuffer = CalculateBuffer(finalEntryPrice, RoundOfType.Floor)
                                                finalStoplossRemark = String.Format("Stoploss: {0}", Math.Round(finalStoplossPrice - finalEntryPrice, 4))

                                                'If TrailingSL Then
                                                '    finalTargetPrice = finalEntryPrice - 100000
                                                'Else
                                                finalTargetPrice = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellTarget
                                                'End If
                                                finalTargetRemark = String.Format("Target: {0}", Math.Round(finalEntryPrice - finalTargetPrice, 4))

                                                squareOffValue = finalEntryPrice - XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellTarget
                                                quantity = XDayRuleSignalStocksPayload(stockName)(signalCandleTime).SellQuantity

                                                Dim tradeTag As String = System.Guid.NewGuid.ToString()

                                                If PartialExit Then
                                                    Dim modifiedQuantity As Integer = quantity / 2
                                                    runningTrade = New Trade(Me,
                                                                currentMinuteCandlePayload.TradingSymbol,
                                                                _StockType,
                                                                currentMinuteCandlePayload.PayloadDate,
                                                                Trade.TradeExecutionDirection.Sell,
                                                                finalEntryPrice,
                                                                entryBuffer,
                                                                Trade.TradeType.MIS,
                                                                Trade.TradeEntryCondition.Original,
                                                                "Signal Candle Low Price",
                                                                modifiedQuantity,
                                                                lotSize,
                                                                finalTargetPrice,
                                                                finalTargetRemark,
                                                                finalStoplossPrice,
                                                                stoplossBuffer,
                                                                finalStoplossRemark,
                                                                currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)

                                                    finalTargetPrice = finalEntryPrice - 100
                                                    finalTargetRemark = String.Format("Target: 100")

                                                    runningTrade2 = New Trade(Me,
                                                               currentMinuteCandlePayload.TradingSymbol,
                                                               _StockType,
                                                               currentMinuteCandlePayload.PayloadDate,
                                                               Trade.TradeExecutionDirection.Sell,
                                                               finalEntryPrice,
                                                               entryBuffer,
                                                               Trade.TradeType.MIS,
                                                               Trade.TradeEntryCondition.Original,
                                                               "Signal Candle Low Price",
                                                               quantity - modifiedQuantity,
                                                               lotSize,
                                                               finalTargetPrice,
                                                               finalTargetRemark,
                                                               finalStoplossPrice,
                                                               stoplossBuffer,
                                                               finalStoplossRemark,
                                                               currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
                                                Else
                                                    runningTrade = New Trade(Me,
                                                                currentMinuteCandlePayload.TradingSymbol,
                                                                _StockType,
                                                                currentMinuteCandlePayload.PayloadDate,
                                                                Trade.TradeExecutionDirection.Sell,
                                                                finalEntryPrice,
                                                                entryBuffer,
                                                                Trade.TradeType.MIS,
                                                                Trade.TradeEntryCondition.Original,
                                                                "Signal Candle Low Price",
                                                                quantity,
                                                                lotSize,
                                                                finalTargetPrice,
                                                                finalTargetRemark,
                                                                finalStoplossPrice,
                                                                stoplossBuffer,
                                                                finalStoplossRemark,
                                                                currentMinuteCandlePayload)

                                                    runningTrade.UpdateTrade(Tag:=tradeTag, SquareOffValue:=squareOffValue, Supporting1:=supporting1, Supporting2:=supporting2, Supporting3:=supporting3, Supporting4:=supporting4, Supporting5:=supporting5)
                                                End If
                                                'Dim potentialOpenTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                                'If potentialOpenTrades IsNot Nothing AndAlso potentialOpenTrades.Count > 0 Then
                                                '    For Each potentialOpenTrade In potentialOpenTrades
                                                '        If potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Sell Then
                                                '            If PartialExit Then
                                                '                If potentialOpenTrade.EntryPrice - potentialOpenTrade.PotentialTarget <= potentialOpenTrade.EntryPrice - potentialOpenTrade.SquareOffValue Then
                                                '                    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                '                ElseIf potentialOpenTrade.EntryPrice - potentialOpenTrade.PotentialTarget > potentialOpenTrade.EntryPrice - potentialOpenTrade.SquareOffValue Then
                                                '                    If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade2)
                                                '                End If
                                                '            Else
                                                '                'If runningTrade IsNot Nothing Then PlaceOrModifyOrder(potentialOpenTrade, runningTrade)
                                                '            End If
                                                '        ElseIf potentialOpenTrade.EntryDirection = Trade.TradeExecutionDirection.Buy Then
                                                '            'Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                '            'dummyPayload.PayloadDate = potentialTickSignalTime
                                                '            'CancelTrade(potentialOpenTrade, dummyPayload, "Opposite Direction Signal")
                                                '        End If
                                                '    Next
                                                'Else
                                                '    If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                '    If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                                'End If
                                                If runningTrade IsNot Nothing Then PlaceOrModifyOrder(runningTrade, Nothing)
                                                If runningTrade2 IsNot Nothing Then PlaceOrModifyOrder(runningTrade2, Nothing)
                                            End If
                                        End If
                                    Else
                                        Dim extraCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If extraCancelTrades IsNot Nothing AndAlso extraCancelTrades.Count > 0 Then
                                            For Each extraCancelTrade In extraCancelTrades
                                                Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                dummyPayload.PayloadDate = potentialTickSignalTime
                                                CancelTrade(extraCancelTrade, dummyPayload, String.Format("{0} Trade already executed", NumberOfTradePerDay))
                                            Next
                                        End If
                                        If extraCancelTrades IsNot Nothing Then extraCancelTrades.Clear()
                                        extraCancelTrades = Nothing
                                    End If
                                    If currentSecondTickPayload IsNot Nothing AndAlso currentSecondTickPayload.Count > 0 Then
                                        For Each tick In currentSecondTickPayload
                                            If IncludeSlippage AndAlso timeChart IsNot Nothing AndAlso timeChart.Count > 0 AndAlso timeChart.ContainsKey(stockName) Then
                                                If tick.PayloadDate > timeChart(stockName) Then
                                                    timeChart.Remove(stockName)
                                                Else
                                                    Continue For
                                                End If
                                            End If

                                            SetCurrentLTPForStock(currentMinuteCandlePayload, tick, Trade.TradeType.MIS)

                                            'Enter Trade
                                            Dim potentialEntryTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                            If potentialEntryTrades IsNot Nothing AndAlso potentialEntryTrades.Count > 0 Then
                                                Dim orderEnterd As Boolean = False
                                                For Each potentialEntryTrade In potentialEntryTrades
                                                    'If tick.Open <> potentialEntryTrade.EntryPrice Then
                                                    '    Continue For
                                                    'End If
                                                    If Not CountTradesWithBreakevenMovement Then
                                                        numberOfExecutedTradePerDay = NumberOfTradesPerDayWithoutBreakevenExit(currentMinuteCandlePayload.PayloadDate)
                                                        numberOfExecutedTradePerStockPerDay = NumberOfTradesPerStockPerDayWithoutBreakevenExit(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                                                    Else
                                                        numberOfExecutedTradePerDay = NumberOfTradesPerDay(currentMinuteCandlePayload.PayloadDate)
                                                        numberOfExecutedTradePerStockPerDay = NumberOfTradesPerStockPerDay(currentMinuteCandlePayload.PayloadDate, currentMinuteCandlePayload.TradingSymbol)
                                                    End If
                                                    If numberOfExecutedTradePerDay < NumberOfTradePerDay Then
                                                        If numberOfExecutedTradePerStockPerDay < NumberOfTradePerStockPerDay Then
                                                            If ReverseSignalTrade Then
                                                                tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS, potentialEntryTrade.EntryDirection)
                                                            Else
                                                                tradeActive = IsTradeActive(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                                            End If
                                                            If Not tradeActive Then
                                                                If IsAnyTradeOfTheStockTargetReached(currentMinuteCandlePayload, Trade.TradeType.MIS) Then
                                                                    CancelTrade(potentialEntryTrade, currentMinuteCandlePayload, "Previous Trade Target reached")
                                                                Else
                                                                    If SameDirectionTrade Then
                                                                        Dim lastTrade As Trade = GetLastExitTradeOfTheStock(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                                                        If lastTrade IsNot Nothing AndAlso lastTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso
                                                                            lastTrade.PLPoint > 0 AndAlso lastTrade.EntryDirection = potentialEntryTrade.EntryDirection Then
                                                                            Dim exitMinuteBlock As Date = New Date(lastTrade.ExitTime.Year,
                                                                                                                    lastTrade.ExitTime.Month,
                                                                                                                    lastTrade.ExitTime.Day,
                                                                                                                    lastTrade.ExitTime.Hour,
                                                                                                                    Math.Floor(lastTrade.ExitTime.Minute / _SignalTimeFrame) * _SignalTimeFrame, 0)
                                                                            Dim currentMinuteBlock As Date = New Date(currentMinuteCandlePayload.PayloadDate.Year,
                                                                                                                    currentMinuteCandlePayload.PayloadDate.Month,
                                                                                                                    currentMinuteCandlePayload.PayloadDate.Day,
                                                                                                                    currentMinuteCandlePayload.PayloadDate.Hour,
                                                                                                                    Math.Floor(currentMinuteCandlePayload.PayloadDate.Minute / _SignalTimeFrame) * _SignalTimeFrame, 0)
                                                                            If currentMinuteCandlePayload.PayloadDate >= exitMinuteBlock.AddMinutes(_SignalTimeFrame) Then
                                                                                If IsAnyCandleClosesAboveOrBelow(currentMinuteBlock, exitMinuteBlock, XDayXMinuteStocksPayload(stockName), potentialEntryTrade) Then
                                                                                    Dim placeOrderResponse As Tuple(Of Boolean, Date) = EnterTradeIfPossible(potentialEntryTrade, tick, GetForwardTicksWithLevel(currentDayOneMinuteStocksPayload(stockName), tick.PayloadDate))
                                                                                    If placeOrderResponse IsNot Nothing AndAlso placeOrderResponse.Item1 Then
                                                                                        If timeChart Is Nothing Then timeChart = New Dictionary(Of String, Date)
                                                                                        timeChart(stockName) = placeOrderResponse.Item2
                                                                                        orderEnterd = True
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        Else
                                                                            Dim placeOrderResponse As Tuple(Of Boolean, Date) = EnterTradeIfPossible(potentialEntryTrade, tick, GetForwardTicksWithLevel(currentDayOneMinuteStocksPayload(stockName), tick.PayloadDate))
                                                                            If placeOrderResponse IsNot Nothing AndAlso placeOrderResponse.Item1 Then
                                                                                If timeChart Is Nothing Then timeChart = New Dictionary(Of String, Date)
                                                                                timeChart(stockName) = placeOrderResponse.Item2
                                                                                orderEnterd = True
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        Dim lastTrade As Trade = GetLastExitTradeOfTheStock(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                                                        'If lastTrade Is Nothing OrElse
                                                                        '    (lastTrade IsNot Nothing AndAlso lastTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso
                                                                        '    lastTrade.PLPoint > 0 AndAlso lastTrade.EntryDirection <> potentialEntryTrade.EntryDirection) OrElse
                                                                        '    (lastTrade IsNot Nothing AndAlso lastTrade.ExitCondition = Trade.TradeExitCondition.StopLoss AndAlso
                                                                        '    lastTrade.PLPoint < 0) Then
                                                                        '    If EnterTradeIfPossible(potentialEntryTrade, tick) Then
                                                                        '        Console.WriteLine("")
                                                                        '    End If
                                                                        'End If
                                                                        If lastTrade Is Nothing OrElse
                                                                            (lastTrade IsNot Nothing AndAlso lastTrade.EntryDirection <> potentialEntryTrade.EntryDirection) Then
                                                                            Dim placeOrderResponse As Tuple(Of Boolean, Date) = EnterTradeIfPossible(potentialEntryTrade, tick, GetForwardTicksWithLevel(currentDayOneMinuteStocksPayload(stockName), tick.PayloadDate))
                                                                            If placeOrderResponse IsNot Nothing AndAlso placeOrderResponse.Item1 Then
                                                                                If timeChart Is Nothing Then timeChart = New Dictionary(Of String, Date)
                                                                                timeChart(stockName) = placeOrderResponse.Item2
                                                                                orderEnterd = True
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                                'If ReverseSignalTrade Then
                                                                '    If EnterTradeIfPossible(potentialEntryTrade, tick) Then
                                                                '        Console.WriteLine("")
                                                                '    End If
                                                                'Else
                                                                '    Dim lastExecutedTrade As Trade = GetLastExitTradeOfTheStock(currentMinuteCandlePayload, Trade.TradeType.MIS)
                                                                '    If lastExecutedTrade Is Nothing OrElse
                                                                '        (lastExecutedTrade IsNot Nothing AndAlso
                                                                '        GetDateTimeTillMinutes(potentialEntryTrade.EntryTime) > GetDateTimeTillMinutes(lastExecutedTrade.ExitTime)) Then
                                                                '        If EnterTradeIfPossible(potentialEntryTrade, tick) Then
                                                                '            Console.WriteLine("")
                                                                '        End If
                                                                '    End If
                                                                'End If
                                                            End If
                                                        Else
                                                            CancelTrade(potentialEntryTrade, currentMinuteCandlePayload, "Max Number Of Trade Per Stock reached")
                                                        End If
                                                    Else
                                                        CancelTrade(potentialEntryTrade, currentMinuteCandlePayload, "Max Number Of Trade Per Day reached")
                                                    End If
                                                Next
                                                If orderEnterd AndAlso IncludeSlippage Then Continue For
                                            End If

                                            'Modify Target Stoploss
                                            Dim potentialModifyTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialModifyTrades IsNot Nothing AndAlso potentialModifyTrades.Count > 0 Then
                                                For Each potentialModifyTrade In potentialModifyTrades
                                                    If potentialModifyTrade.CoreTradingSymbol = stockName Then
                                                        Dim signalCheckTime As Date = GetPreviousXMinuteCandleTime(potentialCandleSignalTime, XDayXMinuteStocksPayload(stockName), _SignalTimeFrame)
                                                        Dim modifiedTarget As Decimal = 0
                                                        Dim modifiedStoploss As Decimal = 0
                                                        Dim modifiedTargetRemark As String = Nothing
                                                        Dim modifiedStoplossRemark As String = Nothing
                                                        If ModifyTarget Then
                                                            If XDayRuleModifyTargetStocksPayload.ContainsKey(stockName) AndAlso
                                                                XDayRuleModifyTargetStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                                XDayRuleModifyTargetStocksPayload(stockName)(signalCheckTime) <> 0 Then
                                                                modifiedTarget = XDayRuleModifyTargetStocksPayload(stockName)(signalCheckTime)
                                                                modifiedTargetRemark = String.Format("Modified Target: {0}, Time:{1}", modifiedTarget, signalCheckTime.ToShortTimeString)
                                                                potentialModifyTrade.UpdateTrade(PotentialTarget:=modifiedTarget, TargetRemark:=modifiedTargetRemark)
                                                            End If
                                                        End If
                                                        If ModifyStoploss Then
                                                            If XDayRuleModifyStoplossStocksPayload.ContainsKey(stockName) AndAlso
                                                                XDayRuleModifyStoplossStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                                XDayRuleModifyStoplossStocksPayload(stockName)(signalCheckTime) <> 0 Then
                                                                modifiedStoploss = XDayRuleModifyStoplossStocksPayload(stockName)(signalCheckTime)
                                                                modifiedStoplossRemark = String.Format("Modified Stoploss: {0}, Time:{1}", modifiedStoploss, signalCheckTime.ToShortTimeString)
                                                                potentialModifyTrade.UpdateTrade(PotentialStopLoss:=modifiedStoploss, SLRemark:=modifiedStoplossRemark)
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If

                                            'Stock Point Check
                                            'If StockPLPoint(tradeCheckingDate, tick.TradingSymbol) >= 0.05 OrElse
                                            '        StockPLPoint(tradeCheckingDate, tick.TradingSymbol) <= -0.1 Then
                                            '    ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max Stock PL Point reached for the day")
                                            'End If

                                            'Stock MTM Check
                                            If ExitOnStockFixedTargetStoploss Then
                                                If StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) >= StockMaxProfitPerDay Then
                                                    ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max Stock Profit reached for the day")
                                                ElseIf StockPLAfterBrokerage(tradeCheckingDate, tick.TradingSymbol) <= StockMaxLossPerDay Then
                                                    ExitStockTradesByForce(tick, Trade.TradeType.MIS, "Max Stock Loss reached for the day")
                                                End If
                                            End If

                                            'OverAll MTM Check
                                            If ExitOnOverAllFixedTargetStoploss Then
                                                If AllPLAfterBrokerage(tradeCheckingDate) >= OverAllProfitPerDay Then
                                                    ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "Max Profit reached for the day")
                                                ElseIf AllPLAfterBrokerage(tradeCheckingDate) <= OverAllLossPerDay Then
                                                    ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "Max Loss reached for the day")
                                                End If
                                            End If

                                            'Exit Trade
                                            Dim potentialExitTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                            If potentialExitTrades IsNot Nothing AndAlso potentialExitTrades.Count > 0 Then
                                                Dim orderExited As Boolean = False
                                                For Each potentialExitTrade In potentialExitTrades
                                                    Dim exitOrderResponse As Tuple(Of Boolean, Date) = ExitTradeIfPossible(potentialExitTrade, tick, GetForwardTicksWithLevel(currentDayOneMinuteStocksPayload(stockName), tick.PayloadDate))
                                                    If exitOrderResponse IsNot Nothing AndAlso exitOrderResponse.Item1 Then
                                                        If timeChart Is Nothing Then timeChart = New Dictionary(Of String, Date)
                                                        timeChart(stockName) = exitOrderResponse.Item2
                                                        orderExited = True
                                                    End If
                                                Next
                                                If orderExited AndAlso IncludeSlippage Then Continue For
                                            End If

                                            'Trailing SL
                                            If TrailingSL Then
                                                Dim potentialSLMoveTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Inprogress)
                                                If potentialSLMoveTrades IsNot Nothing AndAlso potentialSLMoveTrades.Count > 0 Then
                                                    For Each potentialSLMoveTrade In potentialSLMoveTrades
                                                        Dim slMoveRemark As String = Nothing
                                                        Dim slPrice As Double = CalculateLogicalTrailingSL(potentialSLMoveTrade, tick.Open, tick.PayloadDate, slMoveRemark)
                                                        If Math.Round(potentialSLMoveTrade.EntryPrice, 4) <> Math.Round(slPrice, 4) Then
                                                            If potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Buy AndAlso
                                                                slPrice > potentialSLMoveTrade.PotentialStopLoss Then
                                                                MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, Math.Max(potentialSLMoveTrade.PotentialStopLoss, slPrice), slMoveRemark)
                                                            ElseIf potentialSLMoveTrade.EntryDirection = Trade.TradeExecutionDirection.Sell AndAlso
                                                                slPrice < potentialSLMoveTrade.PotentialStopLoss Then
                                                                MoveStopLoss(tick.PayloadDate, potentialSLMoveTrade, Math.Min(potentialSLMoveTrade.PotentialStopLoss, slPrice), slMoveRemark)
                                                            End If
                                                        End If
                                                        slMoveRemark = Nothing
                                                        slPrice = Nothing
                                                    Next
                                                End If
                                                If potentialSLMoveTrades IsNot Nothing Then potentialSLMoveTrades.Clear()
                                                potentialSLMoveTrades = Nothing
                                            End If


                                            If potentialEntryTrades IsNot Nothing Then potentialEntryTrades.Clear()
                                            potentialEntryTrades = Nothing
                                            If potentialModifyTrades IsNot Nothing Then potentialModifyTrades.Clear()
                                            potentialModifyTrades = Nothing
                                            If potentialExitTrades IsNot Nothing Then potentialExitTrades.Clear()
                                            potentialExitTrades = Nothing
                                        Next
                                    End If
                                    'Cancel trade
                                    If startSecond.Hours = endSecond.Hours AndAlso startSecond.Minutes = endSecond.Minutes AndAlso startSecond.Seconds = endSecond.Seconds Then
                                        Dim potentialCancelTrades As List(Of Trade) = GetSpecificTrades(currentMinuteCandlePayload, Trade.TradeType.MIS, Trade.TradeExecutionStatus.Open)
                                        If potentialCancelTrades IsNot Nothing AndAlso potentialCancelTrades.Count > 0 Then
                                            For Each potentialCancelTrade In potentialCancelTrades
                                                If potentialCancelTrade.CoreTradingSymbol = stockName Then
                                                    Dim dummyPayload As Payload = New Payload(Payload.CandleDataSource.Chart)
                                                    dummyPayload.PayloadDate = potentialTickSignalTime
                                                    Dim signalCheckTime As Date = GetCurrentXMinuteCandleTime(potentialCandleSignalTime, XDayXMinuteStocksPayload(stockName))
                                                    If XDayRuleSignalStocksPayload.ContainsKey(stockName) AndAlso
                                                        XDayRuleSignalStocksPayload(stockName).ContainsKey(signalCheckTime) AndAlso
                                                        (XDayRuleSignalStocksPayload(stockName)(signalCheckTime).BuySignal = 0 OrElse
                                                        XDayRuleSignalStocksPayload(stockName)(signalCheckTime).SellSignal = 0) Then
                                                        CancelTrade(potentialCancelTrade, dummyPayload, String.Format("Invalid Signal"))
                                                    End If
                                                End If
                                            Next
                                        End If
                                        If potentialCancelTrades IsNot Nothing Then potentialCancelTrades.Clear()
                                        potentialCancelTrades = Nothing
                                    End If

#Region "Cleanup"
                                    runningTrade = Nothing
                                    runningTrade2 = Nothing
                                    If currentSecondTickPayload IsNot Nothing Then currentSecondTickPayload.Clear()
                                    currentSecondTickPayload = Nothing

                                    finalEntryPrice = Nothing
                                    finalTargetPrice = Nothing
                                    finalTargetRemark = Nothing
                                    finalStoplossPrice = Nothing
                                    finalStoplossRemark = Nothing
                                    squareOffValue = Nothing
                                    quantity = Nothing
                                    lotSize = Nothing
                                    entryBuffer = Nothing
                                    stoplossBuffer = Nothing
                                    supporting1 = Nothing
                                    supporting2 = Nothing
                                    supporting3 = Nothing
                                    supporting4 = Nothing
                                    supporting5 = Nothing
                                    supporting6 = Nothing
                                    supporting7 = Nothing
                                    supporting8 = Nothing
                                    supporting9 = Nothing
                                    supporting10 = Nothing

                                    tradeActive = Nothing
                                    numberOfExecutedTradePerDay = Nothing
                                    numberOfExecutedTradePerStockPerDay = Nothing

#End Region
                                Next
                                startSecond = startSecond.Add(TimeSpan.FromSeconds(1))
                            End While   'Second Loop
                            'Force exit at day end
                            If startMinute = EODExitTime Then
                                ExitAllTradeByForce(potentialTickSignalTime, currentDayOneMinuteStocksPayload, Trade.TradeType.MIS, "EOD Force Exit")
                            End If
                            startMinute = startMinute.Add(TimeSpan.FromMinutes(_SignalTimeFrame))
                        End While   'Minute Loop
                    End If
#Region "Cleanup"
                    If currentDayOneMinuteStocksPayload IsNot Nothing Then currentDayOneMinuteStocksPayload.Clear()
                    If XDayXMinuteStocksPayload IsNot Nothing Then XDayXMinuteStocksPayload.Clear()
                    If XDayRuleSignalStocksPayload IsNot Nothing Then XDayRuleSignalStocksPayload.Clear()
                    If XDayRuleModifyStoplossStocksPayload IsNot Nothing Then XDayRuleModifyStoplossStocksPayload.Clear()
                    If XDayRuleModifyTargetStocksPayload IsNot Nothing Then XDayRuleModifyTargetStocksPayload.Clear()
                    If XDayRuleSupporting1StocksPayload IsNot Nothing Then XDayRuleSupporting1StocksPayload.Clear()
                    If XDayRuleSupporting2StocksPayload IsNot Nothing Then XDayRuleSupporting2StocksPayload.Clear()
                    If XDayRuleSupporting3StocksPayload IsNot Nothing Then XDayRuleSupporting3StocksPayload.Clear()
                    If XDayRuleSupporting4StocksPayload IsNot Nothing Then XDayRuleSupporting4StocksPayload.Clear()
                    If XDayRuleSupporting5StocksPayload IsNot Nothing Then XDayRuleSupporting5StocksPayload.Clear()
                    If XDayRuleSupporting6StocksPayload IsNot Nothing Then XDayRuleSupporting6StocksPayload.Clear()
                    If XDayRuleSupporting7StocksPayload IsNot Nothing Then XDayRuleSupporting7StocksPayload.Clear()
                    If XDayRuleSupporting8StocksPayload IsNot Nothing Then XDayRuleSupporting8StocksPayload.Clear()
                    If XDayRuleSupporting9StocksPayload IsNot Nothing Then XDayRuleSupporting9StocksPayload.Clear()
                    If XDayRuleSupporting10StocksPayload IsNot Nothing Then XDayRuleSupporting10StocksPayload.Clear()
                    currentDayOneMinuteStocksPayload = Nothing
                    XDayXMinuteStocksPayload = Nothing
                    XDayRuleSignalStocksPayload = Nothing
                    XDayRuleModifyStoplossStocksPayload = Nothing
                    XDayRuleModifyTargetStocksPayload = Nothing
                    XDayRuleSupporting1StocksPayload = Nothing
                    XDayRuleSupporting2StocksPayload = Nothing
                    XDayRuleSupporting3StocksPayload = Nothing
                    XDayRuleSupporting4StocksPayload = Nothing
                    XDayRuleSupporting5StocksPayload = Nothing
                    XDayRuleSupporting6StocksPayload = Nothing
                    XDayRuleSupporting7StocksPayload = Nothing
                    XDayRuleSupporting8StocksPayload = Nothing
                    XDayRuleSupporting9StocksPayload = Nothing
                    XDayRuleSupporting10StocksPayload = Nothing

#End Region


                End If
                SetOverallDrawUpDrawDownForTheDay(tradeCheckingDate)
                totalPL += AllPLAfterBrokerage(tradeCheckingDate)
                tradeCheckingDate = tradeCheckingDate.AddDays(1)

                'Serialization
                If TradesTaken IsNot Nothing AndAlso TradesTaken.Count > 0 Then
                    OnHeartbeat("Serializing Trades collection")
                    Utilities.Strings.SerializeFromCollectionUsingFileStream(Of Dictionary(Of Date, Dictionary(Of String, List(Of Trade))))(tradesFileName, TradesTaken)
                End If

#Region "Cleanup"
                'If tradeCheckingDate < endDate.Date Then
                If TradesTaken IsNot Nothing Then TradesTaken.Clear()
                'If CapitalMovement IsNot Nothing Then CapitalMovement.Clear()
                TradesTaken = Nothing
                'CapitalMovement = Nothing
                'End If

                If stockList IsNot Nothing Then stockList.Clear()
                stockList = Nothing

                GC.Collect()
#End Region

            End While   'Date Loop

            If CapitalMovement IsNot Nothing Then Utilities.Strings.SerializeFromCollection(Of Dictionary(Of Date, List(Of Capital)))(capitalFileName, CapitalMovement)

            PrintArrayToExcel(filename, tradesFileName, capitalFileName)
        End If
    End Function
#End Region

#Region "Stock Selection"
    Private Function GetStockData(tradingDate As Date) As Dictionary(Of String, Decimal())
        Dim ret As Dictionary(Of String, Decimal()) = Nothing
        If StockFileName IsNot Nothing Then
            Dim dt As DataTable = Nothing
            Using csvHelper As New Utilities.DAL.CSVHelper(StockFileName, ",", _canceller)
                dt = csvHelper.GetDataTableFromCSV(1)
            End Using
            If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                Dim counter As Integer = 0
                For i = 1 To dt.Rows.Count - 1
                    Dim rowDate As Date = dt.Rows(i)(0)
                    If rowDate.Date = tradingDate.Date Then
                        If ret Is Nothing Then ret = New Dictionary(Of String, Decimal())
                        Dim tradingSymbol As String = dt.Rows(i).Item(1)
                        Dim instrumentName As String = Nothing
                        If tradingSymbol.Contains("FUT") Then
                            instrumentName = tradingSymbol.Remove(tradingSymbol.Count - 8)
                        Else
                            instrumentName = tradingSymbol
                        End If
                        ret.Add(instrumentName, {dt.Rows(i).Item(5), dt.Rows(i).Item(5), dt.Rows(i).Item(3), dt.Rows(i).Item(4)})
                        counter += 1
                        If counter = Me.NumberOfTradeableStockPerDay Then Exit For
                    End If
                Next
            End If
        End If
        Return ret
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                PartialExit = Nothing
                NumberOfTradePerStockPerDay = Nothing
                NumberOfTradePerDay = Nothing
                CountTradesWithBreakevenMovement = Nothing
                TrailingSL = Nothing
                ReverseSignalTrade = Nothing
                ExitOnStockFixedTargetStoploss = Nothing
                StockMaxProfitPerDay = Nothing
                StockMaxLossPerDay = Nothing
                ExitOnOverAllFixedTargetStoploss = Nothing
                ModifyTarget = Nothing
                ModifyStoploss = Nothing
                SameDirectionTrade = Nothing
                NIFTY50Stocks = Nothing
                QuantityFlag = Nothing
                MaxStoplossAmount = Nothing
                StockFileName = Nothing
                FirstTradeTargetMultiplier = Nothing
                EarlyStoploss = Nothing
                ForwardTradeTargetMultiplier = Nothing
                CapitalToBeUsed = Nothing
                CandleBasedEntry = Nothing
                If _StockData IsNot Nothing Then _StockData.Clear()

                _canceller = Nothing
                _common = Nothing
                If TradesTaken IsNot Nothing Then TradesTaken.Clear()
                TradesTaken = Nothing
                TickSize = Nothing
                EODExitTime = Nothing
                LastTradeEntryTime = Nothing
                ExchangeStartTime = Nothing
                ExchangeEndTime = Nothing
                MarginMultiplier = Nothing

                InitialCapital = Nothing
                CapitalForPumpIn = Nothing
                MinimumEarnedCapitalToWithdraw = Nothing
                AmountToBeWithdrawn = Nothing
                NumberOfTradeableStockPerDay = Nothing
                PerTradeMaxProfitPercentage = Nothing
                PerTradeMaxLossPercentage = Nothing
                PerStockMaxProfitPercentage = Nothing
                PerStockMaxLossPercentage = Nothing
                OverAllProfitPerDay = Nothing
                OverAllLossPerDay = Nothing
                If CapitalMovement IsNot Nothing Then CapitalMovement.Clear()
                CapitalMovement = Nothing

                GC.Collect()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    Protected Overrides Sub Finalize()
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(False)
        MyBase.Finalize()
    End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
