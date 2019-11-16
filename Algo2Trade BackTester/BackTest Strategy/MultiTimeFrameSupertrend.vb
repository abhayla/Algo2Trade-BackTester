Imports System.Threading
Imports Algo2TradeBLL
Public Class MultiTimeFrameSupertrend
    Inherits Strategy
    Dim cts As CancellationTokenSource
    Dim cmn As Common = New Common(cts)

    Public Sub Run(start_date As DateTime, end_date As DateTime, validIndicatorTimeFrame As Dictionary(Of IndicatorTimeFrame, Boolean))
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat

        Dim mainSignalTimeFrame As Integer = 5
        Dim firstAlternateTimeFrame As Integer = 15
        Dim secondAlternateTimeFrame As Integer = 60
        Dim eodTimeFrame As Boolean = True
        Dim targetMultiplier As Double = 2
        Dim stopLossMultiplier As Double = 1
        'TO DO:
        'Change Time
        Dim exchangeStartTime As Date = "10:00:00"
        Dim exchangeEndTime As Date = "23:30:00"
        Dim endOfDay As DateTime = "23:00:00"

        Dim dataCheckDate As Date = "2018-10-05"

        Dim from_date As DateTime = start_date
        Dim to_Date As Date = end_date
        Dim chk_date As Date = from_date
        Dim dateCtr As Integer = 0
        Dim previousTradingDate As Date = Nothing
        Dim previousValidSignal As Boolean = False
        Dim previousMainSignal As Boolean = False
        Dim entryCondition As String = Nothing
        Dim cncNextDayOpenPriceExitFlag As Boolean = False

        Dim stockList As Dictionary(Of String, Double) = Nothing
        TradesTaken = New Dictionary(Of Date, Dictionary(Of String, List(Of Trade)))

        Dim tempStocklist As Dictionary(Of String, Double()) = Nothing
        'TO DO:
        'Change Stock Name
        tempStocklist = New Dictionary(Of String, Double()) From {{"ZINC", {0.5, 1}}, {"LEAD", {0.5, 1}}, {"ALUMINIUM", {0.5, 1}}, {"COPPER", {1.2, 1}}, {"NICKEL", {2.4, 1}}, {"SILVER", {170, 1}}, {"GOLD", {40, 1}}, {"CRUDEOIL", {8, 1}}, {"NATURALGAS", {0.6, 1}}}        '{"StockName",{targetBuffer,Dummy}}
        'tempStocklist = New Dictionary(Of String, Double()) From {{"ZINC", {0.5, 1}}}        '{"StockName",{targetBuffer,Dummy}}
        For Each tradingSymbol In tempStocklist.Keys
            If stockList Is Nothing Then stockList = New Dictionary(Of String, Double)
            stockList.Add(tradingSymbol, tempStocklist(tradingSymbol)(0))
        Next

        While chk_date <= to_Date
            dateCtr += 1
            OnHeartbeat(String.Format("Running for date:{0}/{1}", dateCtr, DateDiff(DateInterval.Day, from_date, to_Date) + 1))

            If stockList IsNot Nothing AndAlso stockList.Count > 0 Then
                Dim OneMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim MainXMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim FirstAlternateXMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim SecondAlternateXMinutePayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim EODPayload As Dictionary(Of String, Dictionary(Of Date, Payload)) = Nothing
                Dim ATRPaylaod As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim MACDPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim MACDSignalPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim HistogramPayload As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendMain1 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendMain2 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendMain3 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendMain4 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendFirstAlternate1 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendFirstAlternate2 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendFirstAlternate3 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendFirstAlternate4 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendSecondAlternate1 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendSecondAlternate2 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendSecondAlternate3 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendSecondAlternate4 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendEOD1 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendEOD2 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendEOD3 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendEOD4 As Dictionary(Of String, Dictionary(Of Date, Decimal)) = Nothing
                Dim SupertrendColorMain1 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorMain2 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorMain3 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorMain4 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorFirstAlternate1 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorFirstAlternate2 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorFirstAlternate3 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorFirstAlternate4 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorSecondAlternate1 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorSecondAlternate2 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorSecondAlternate3 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorSecondAlternate4 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorEOD1 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorEOD2 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorEOD3 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing
                Dim SupertrendColorEOD4 As Dictionary(Of String, Dictionary(Of Date, Color)) = Nothing

                For Each item In stockList.Keys
                    Dim current_OneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim current_mainXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_OneMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_mainXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_firstAlternateXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_secondAlternateXMinutePayload As Dictionary(Of Date, Payload) = Nothing
                    Dim temp_eodPayload As Dictionary(Of Date, Payload) = Nothing

                    'TO DO:
                    'Change fetching data from database table name & start date
                    Dim currentTradingSymbol As String = cmn.GetCurrentTradingSymbol(Common.DataBaseTable.EOD_Commodity, item, chk_date)
                    If PastIntradayData IsNot Nothing AndAlso PastIntradayData.ContainsKey(currentTradingSymbol) Then
                        temp_OneMinutePayload = PastIntradayData(currentTradingSymbol)
                    Else
                        temp_OneMinutePayload = cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.Intraday_Commodity, currentTradingSymbol, chk_date.AddMonths(-3), dataCheckDate)
                        If PastIntradayData Is Nothing Then PastIntradayData = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        PastIntradayData.Add(currentTradingSymbol, temp_OneMinutePayload)
                    End If
                    If PastEODData IsNot Nothing AndAlso PastEODData.ContainsKey(currentTradingSymbol) Then
                        temp_eodPayload = PastEODData(currentTradingSymbol)
                    Else
                        temp_eodPayload = cmn.GetRawPayloadForSpecificTradingSymbol(Common.DataBaseTable.EOD_Commodity, currentTradingSymbol, dataCheckDate.AddYears(-2), dataCheckDate)
                        If PastEODData Is Nothing Then PastEODData = New Dictionary(Of String, Dictionary(Of Date, Payload))
                        PastEODData.Add(currentTradingSymbol, temp_eodPayload)
                    End If
                    OnHeartbeat(String.Format("Processing Data for {0}", chk_date.ToShortDateString))

                    If temp_OneMinutePayload IsNot Nothing AndAlso temp_OneMinutePayload.Count > 0 AndAlso temp_eodPayload IsNot Nothing AndAlso temp_eodPayload.Count > 0 Then

                        For Each tempKeys In temp_OneMinutePayload.Keys
                            If tempKeys.Date = chk_date.Date Then
                                If current_OneMinutePayload Is Nothing Then current_OneMinutePayload = New Dictionary(Of Date, Payload)
                                current_OneMinutePayload.Add(tempKeys, temp_OneMinutePayload(tempKeys))
                            End If
                        Next

                        If current_OneMinutePayload IsNot Nothing AndAlso current_OneMinutePayload.Count > 0 Then
                            Dim outputATR As Dictionary(Of Date, Decimal) = Nothing
                            Dim ouputMACDPayload As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputsignalPayload As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputhistogramPayload As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendMain1 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendMain2 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendMain3 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendMain4 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendFirstAlternate1 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendFirstAlternate2 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendFirstAlternate3 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendFirstAlternate4 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendSecondAlternate1 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendSecondAlternate2 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendSecondAlternate3 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendSecondAlternate4 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendEOD1 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendEOD2 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendEOD3 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendEOD4 As Dictionary(Of Date, Decimal) = Nothing
                            Dim outputSupertrendColorMain1 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorMain2 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorMain3 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorMain4 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorFirstAlternate1 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorFirstAlternate2 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorFirstAlternate3 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorFirstAlternate4 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorSecondAlternate1 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorSecondAlternate2 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorSecondAlternate3 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorSecondAlternate4 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorEOD1 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorEOD2 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorEOD3 As Dictionary(Of Date, Color) = Nothing
                            Dim outputSupertrendColorEOD4 As Dictionary(Of Date, Color) = Nothing

                            temp_mainXMinutePayload = cmn.ConvertPayloadsToXMinutes(temp_OneMinutePayload, mainSignalTimeFrame)
                            temp_firstAlternateXMinutePayload = cmn.ConvertPayloadsToXMinutes(temp_OneMinutePayload, firstAlternateTimeFrame)
                            temp_secondAlternateXMinutePayload = cmn.ConvertPayloadsToXMinutes(temp_OneMinutePayload, secondAlternateTimeFrame)

                            Indicator.ATR.CalculateATR(10, temp_mainXMinutePayload, outputATR)

                            Indicator.MACD.CalculateMACD(12, 26, 9, temp_mainXMinutePayload, ouputMACDPayload, outputsignalPayload, outputhistogramPayload)

                            Indicator.Supertrend.CalculateSupertrend(7, 3, temp_mainXMinutePayload, outputSupertrendMain1, outputSupertrendColorMain1)
                            Indicator.Supertrend.CalculateSupertrend(10, 3, temp_mainXMinutePayload, outputSupertrendMain2, outputSupertrendColorMain2)
                            Indicator.Supertrend.CalculateSupertrend(11, 2, temp_mainXMinutePayload, outputSupertrendMain3, outputSupertrendColorMain3)
                            Indicator.Supertrend.CalculateSupertrend(144, 3, temp_mainXMinutePayload, outputSupertrendMain4, outputSupertrendColorMain4)

                            Indicator.Supertrend.CalculateSupertrend(7, 3, temp_firstAlternateXMinutePayload, outputSupertrendFirstAlternate1, outputSupertrendColorFirstAlternate1)
                            Indicator.Supertrend.CalculateSupertrend(10, 3, temp_firstAlternateXMinutePayload, outputSupertrendFirstAlternate2, outputSupertrendColorFirstAlternate2)
                            Indicator.Supertrend.CalculateSupertrend(11, 2, temp_firstAlternateXMinutePayload, outputSupertrendFirstAlternate3, outputSupertrendColorFirstAlternate3)
                            Indicator.Supertrend.CalculateSupertrend(144, 3, temp_firstAlternateXMinutePayload, outputSupertrendFirstAlternate4, outputSupertrendColorFirstAlternate4)

                            Indicator.Supertrend.CalculateSupertrend(7, 3, temp_secondAlternateXMinutePayload, outputSupertrendSecondAlternate1, outputSupertrendColorSecondAlternate1)
                            Indicator.Supertrend.CalculateSupertrend(10, 3, temp_secondAlternateXMinutePayload, outputSupertrendSecondAlternate2, outputSupertrendColorSecondAlternate2)
                            Indicator.Supertrend.CalculateSupertrend(11, 2, temp_secondAlternateXMinutePayload, outputSupertrendSecondAlternate3, outputSupertrendColorSecondAlternate3)
                            Indicator.Supertrend.CalculateSupertrend(144, 3, temp_secondAlternateXMinutePayload, outputSupertrendSecondAlternate4, outputSupertrendColorSecondAlternate4)

                            Indicator.Supertrend.CalculateSupertrend(7, 3, temp_eodPayload, outputSupertrendEOD1, outputSupertrendColorEOD1, True)
                            Indicator.Supertrend.CalculateSupertrend(10, 3, temp_eodPayload, outputSupertrendEOD2, outputSupertrendColorEOD2, True)
                            Indicator.Supertrend.CalculateSupertrend(11, 2, temp_eodPayload, outputSupertrendEOD3, outputSupertrendColorEOD3, True)

                            For Each tempKeys In temp_mainXMinutePayload.Keys
                                If tempKeys.Date = chk_date.Date Then
                                    If current_mainXMinutePayload Is Nothing Then current_mainXMinutePayload = New Dictionary(Of Date, Payload)
                                    current_mainXMinutePayload.Add(tempKeys, temp_mainXMinutePayload(tempKeys))
                                End If
                            Next

                            If OneMinutePayload Is Nothing Then OneMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            OneMinutePayload.Add(item, current_OneMinutePayload)
                            If MainXMinutePayload Is Nothing Then MainXMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            MainXMinutePayload.Add(item, current_mainXMinutePayload)
                            If FirstAlternateXMinutePayload Is Nothing Then FirstAlternateXMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            FirstAlternateXMinutePayload.Add(item, temp_firstAlternateXMinutePayload)
                            If SecondAlternateXMinutePayload Is Nothing Then SecondAlternateXMinutePayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            SecondAlternateXMinutePayload.Add(item, temp_secondAlternateXMinutePayload)
                            If EODPayload Is Nothing Then EODPayload = New Dictionary(Of String, Dictionary(Of Date, Payload))
                            EODPayload.Add(item, temp_eodPayload)
                            If ATRPaylaod Is Nothing Then ATRPaylaod = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            ATRPaylaod.Add(item, outputATR)
                            If MACDPayload Is Nothing Then MACDPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            MACDPayload.Add(item, ouputMACDPayload)
                            If MACDSignalPayload Is Nothing Then MACDSignalPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            MACDSignalPayload.Add(item, outputsignalPayload)
                            If HistogramPayload Is Nothing Then HistogramPayload = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            HistogramPayload.Add(item, outputhistogramPayload)
                            If SupertrendMain1 Is Nothing Then SupertrendMain1 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendMain1.Add(item, outputSupertrendMain1)
                            If SupertrendMain2 Is Nothing Then SupertrendMain2 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendMain2.Add(item, outputSupertrendMain2)
                            If SupertrendMain3 Is Nothing Then SupertrendMain3 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendMain3.Add(item, outputSupertrendMain3)
                            If SupertrendMain4 Is Nothing Then SupertrendMain4 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendMain4.Add(item, outputSupertrendMain4)
                            If SupertrendFirstAlternate1 Is Nothing Then SupertrendFirstAlternate1 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendFirstAlternate1.Add(item, outputSupertrendFirstAlternate1)
                            If SupertrendFirstAlternate2 Is Nothing Then SupertrendFirstAlternate2 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendFirstAlternate2.Add(item, outputSupertrendFirstAlternate2)
                            If SupertrendFirstAlternate3 Is Nothing Then SupertrendFirstAlternate3 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendFirstAlternate3.Add(item, outputSupertrendFirstAlternate3)
                            If SupertrendFirstAlternate4 Is Nothing Then SupertrendFirstAlternate4 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendFirstAlternate4.Add(item, outputSupertrendFirstAlternate4)
                            If SupertrendSecondAlternate1 Is Nothing Then SupertrendSecondAlternate1 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendSecondAlternate1.Add(item, outputSupertrendSecondAlternate1)
                            If SupertrendSecondAlternate2 Is Nothing Then SupertrendSecondAlternate2 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendSecondAlternate2.Add(item, outputSupertrendSecondAlternate2)
                            If SupertrendSecondAlternate3 Is Nothing Then SupertrendSecondAlternate3 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendSecondAlternate3.Add(item, outputSupertrendSecondAlternate3)
                            If SupertrendSecondAlternate4 Is Nothing Then SupertrendSecondAlternate4 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendSecondAlternate4.Add(item, outputSupertrendSecondAlternate4)
                            If SupertrendEOD1 Is Nothing Then SupertrendEOD1 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendEOD1.Add(item, outputSupertrendEOD1)
                            If SupertrendEOD2 Is Nothing Then SupertrendEOD2 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendEOD2.Add(item, outputSupertrendEOD2)
                            If SupertrendEOD3 Is Nothing Then SupertrendEOD3 = New Dictionary(Of String, Dictionary(Of Date, Decimal))
                            SupertrendEOD3.Add(item, outputSupertrendEOD3)
                            If SupertrendColorMain1 Is Nothing Then SupertrendColorMain1 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorMain1.Add(item, outputSupertrendColorMain1)
                            If SupertrendColorMain2 Is Nothing Then SupertrendColorMain2 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorMain2.Add(item, outputSupertrendColorMain2)
                            If SupertrendColorMain3 Is Nothing Then SupertrendColorMain3 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorMain3.Add(item, outputSupertrendColorMain3)
                            If SupertrendColorMain4 Is Nothing Then SupertrendColorMain4 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorMain4.Add(item, outputSupertrendColorMain4)
                            If SupertrendColorFirstAlternate1 Is Nothing Then SupertrendColorFirstAlternate1 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorFirstAlternate1.Add(item, outputSupertrendColorFirstAlternate1)
                            If SupertrendColorFirstAlternate2 Is Nothing Then SupertrendColorFirstAlternate2 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorFirstAlternate2.Add(item, outputSupertrendColorFirstAlternate2)
                            If SupertrendColorFirstAlternate3 Is Nothing Then SupertrendColorFirstAlternate3 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorFirstAlternate3.Add(item, outputSupertrendColorFirstAlternate3)
                            If SupertrendColorFirstAlternate4 Is Nothing Then SupertrendColorFirstAlternate4 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorFirstAlternate4.Add(item, outputSupertrendColorFirstAlternate4)
                            If SupertrendColorSecondAlternate1 Is Nothing Then SupertrendColorSecondAlternate1 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorSecondAlternate1.Add(item, outputSupertrendColorSecondAlternate1)
                            If SupertrendColorSecondAlternate2 Is Nothing Then SupertrendColorSecondAlternate2 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorSecondAlternate2.Add(item, outputSupertrendColorSecondAlternate2)
                            If SupertrendColorSecondAlternate3 Is Nothing Then SupertrendColorSecondAlternate3 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorSecondAlternate3.Add(item, outputSupertrendColorSecondAlternate3)
                            If SupertrendColorSecondAlternate4 Is Nothing Then SupertrendColorSecondAlternate4 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorSecondAlternate4.Add(item, outputSupertrendColorSecondAlternate4)
                            If SupertrendColorEOD1 Is Nothing Then SupertrendColorEOD1 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorEOD1.Add(item, outputSupertrendColorEOD1)
                            If SupertrendColorEOD2 Is Nothing Then SupertrendColorEOD2 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorEOD2.Add(item, outputSupertrendColorEOD2)
                            If SupertrendColorEOD3 Is Nothing Then SupertrendColorEOD3 = New Dictionary(Of String, Dictionary(Of Date, Color))
                            SupertrendColorEOD3.Add(item, outputSupertrendColorEOD3)
                        End If
                    End If
                Next

                If OneMinutePayload IsNot Nothing AndAlso OneMinutePayload.Count > 0 AndAlso EODPayload IsNot Nothing AndAlso EODPayload.Count > 0 Then
                    Dim startTime As Date = exchangeStartTime
                    Dim endTime As Date = exchangeEndTime
                    While startTime < endTime
                        Dim signalFlag As Boolean = False
                        For Each stockName In stockList.Keys
                            OnHeartbeat(String.Format("Checking Trade for {0} on {1}", stockName, chk_date.ToShortDateString))
                            Dim tempStockPayload As Dictionary(Of Date, Payload) = MainXMinutePayload(stockName)
                            If startTime = exchangeStartTime Then
                                startTime = New DateTime(startTime.Year, startTime.Month, startTime.Day, tempStockPayload.Keys.FirstOrDefault.Hour, tempStockPayload.Keys.FirstOrDefault.Minute, tempStockPayload.Keys.FirstOrDefault.Second)
                                endTime = New DateTime(endTime.Year, endTime.Month, endTime.Day, tempStockPayload.Keys.LastOrDefault.Hour, tempStockPayload.Keys.LastOrDefault.Minute, tempStockPayload.Keys.LastOrDefault.Second)
                                endTime = endTime.AddMinutes(mainSignalTimeFrame)
                                endOfDay = endTime.AddMinutes(-30)
                            End If
                            If tempStockPayload IsNot Nothing AndAlso tempStockPayload.Count > 0 Then
                                Dim tempStockMACDPayload As Dictionary(Of Date, Decimal) = MACDPayload(stockName)
                                Dim tempStockATRPayload As Dictionary(Of Date, Decimal) = ATRPaylaod(stockName)
                                Dim potentialSignalTime As Date = New DateTime(chk_date.Year, chk_date.Month, chk_date.Day, startTime.Hour, startTime.Minute, startTime.Second)
                                If Not tempStockPayload.ContainsKey(potentialSignalTime) Then
                                    Continue For
                                End If
                                Dim signalTime As Date = Nothing
                                Dim validSignal As Boolean = True
                                Dim previousTempMainSignal As Boolean = False
                                Dim signalColor As Color = Color.Red
                                Dim runningTrade1 As Trade = Nothing
                                Dim runningTrade2 As Trade = Nothing
                                Dim totalTrade As Integer = Nothing

                                'Entry Trade Block Start
                                Dim CNCSpecifictrade As List(Of Trade) = GetSpecificTrades(previousTradingDate, stockName, TradeExecutionStatus.Inprogress)
                                Dim itemSpecificTrade As List(Of Trade) = Nothing
                                If startTime <= endOfDay Then
                                    itemSpecificTrade = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                End If
                                If CNCSpecifictrade IsNot Nothing AndAlso CNCSpecifictrade.Count > 0 Then
                                    totalTrade = totalTrade + CNCSpecifictrade.Count
                                End If
                                If itemSpecificTrade IsNot Nothing AndAlso itemSpecificTrade.Count > 0 Then
                                    totalTrade = totalTrade + itemSpecificTrade.Count
                                End If

                                If Not IsTradeActive(chk_date, stockName) Or totalTrade < 2 Then
                                    validSignal = validSignal And GetSupertrendSignal(potentialSignalTime, signalColor,
                                            SupertrendColorMain1(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_7_3),
                                            SupertrendColorMain2(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_10_3),
                                            SupertrendColorMain3(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_11_2),
                                            SupertrendColorMain4(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_144_3))

                                    If validSignal = False Then
                                        validSignal = True
                                        signalColor = Color.Green
                                        validSignal = validSignal And GetSupertrendSignal(potentialSignalTime, signalColor,
                                            SupertrendColorMain1(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_7_3),
                                            SupertrendColorMain2(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_10_3),
                                            SupertrendColorMain3(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_11_2),
                                            SupertrendColorMain4(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_Supertrend_144_3))
                                    End If
                                    previousTempMainSignal = validSignal
                                    Dim checkTime As Date = GetSpecificTime(potentialSignalTime, FirstAlternateXMinutePayload(stockName), 2)
                                    validSignal = validSignal And GetSupertrendSignal(checkTime, signalColor,
                                            SupertrendColorFirstAlternate1(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF15_Supertrend_7_3),
                                            SupertrendColorFirstAlternate2(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF15_Supertrend_10_3),
                                            SupertrendColorFirstAlternate3(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF15_Supertrend_11_2),
                                            SupertrendColorFirstAlternate4(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF15_Supertrend_144_3))

                                    checkTime = GetSpecificTime(potentialSignalTime, SecondAlternateXMinutePayload(stockName), 2)
                                    validSignal = validSignal And GetSupertrendSignal(checkTime, signalColor,
                                            SupertrendColorSecondAlternate1(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF60_Supertrend_7_3),
                                            SupertrendColorSecondAlternate2(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF60_Supertrend_10_3),
                                            SupertrendColorSecondAlternate3(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF60_Supertrend_11_2),
                                            SupertrendColorSecondAlternate4(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF60_Supertrend_144_3))

                                    checkTime = GetSpecificTime(potentialSignalTime, EODPayload(stockName), 2)
                                    validSignal = validSignal And GetSupertrendSignal(checkTime, signalColor,
                                            SupertrendColorEOD1(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF1D_Supertrend_7_3),
                                            SupertrendColorEOD2(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF1D_Supertrend_10_3),
                                            SupertrendColorEOD3(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF1D_Supertrend_11_2))

                                    validSignal = validSignal And GetMACDSignal(potentialSignalTime, signalColor,
                                                                      MACDPayload(stockName), validIndicatorTimeFrame(IndicatorTimeFrame.TF5_MACD_12_26_9))

                                    If startTime < endOfDay Then
                                        If validSignal = True AndAlso previousMainSignal = True AndAlso validSignal <> previousValidSignal Then
                                            signalTime = potentialSignalTime
                                            signalFlag = True
                                            If entryCondition = "Supertrend Cut" Then
                                                entryCondition = "Supertrend Cut But Color Not Change"
                                            Else
                                                entryCondition = "Fresh Signal"
                                            End If
                                        ElseIf validSignal = True AndAlso validSignal <> previousValidSignal Then
                                            signalTime = potentialSignalTime.AddMinutes(mainSignalTimeFrame)
                                            signalFlag = True
                                            If entryCondition = "Supertrend Cut" Then
                                                entryCondition = "Supertrend Cut But Color Not Change"
                                            Else
                                                entryCondition = "Fresh Signal"
                                            End If
                                        End If
                                    End If
                                    previousMainSignal = previousTempMainSignal
                                    previousValidSignal = validSignal

                                    If signalFlag Then
                                        signalFlag = False
                                        Dim currentStockPayload As Payload = tempStockPayload(signalTime)
                                        Dim currentStockATR As Decimal = tempStockATRPayload(currentStockPayload.PreviousCandlePayload.PayloadDate)
                                        Dim currentStockSuperTrend1 As Decimal = SupertrendMain1(stockName)(currentStockPayload.PreviousCandlePayload.PayloadDate)
                                        Dim currentStockSuperTrend2 As Decimal = SupertrendMain2(stockName)(currentStockPayload.PreviousCandlePayload.PayloadDate)
                                        Dim currentStockSuperTrend3 As Decimal = SupertrendMain3(stockName)(currentStockPayload.PreviousCandlePayload.PayloadDate)
                                        Dim currentStockSuperTrend4 As Decimal = SupertrendMain4(stockName)(currentStockPayload.PreviousCandlePayload.PayloadDate)

                                        If signalColor = Color.Green Then
                                            runningTrade1 = New Trade
                                            With runningTrade1
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentStockPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryTime = currentStockPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .EntryCondition = entryCondition
                                                .SignalCandle = currentStockPayload
                                                .TradingSymbol = currentStockPayload.TradingSymbol
                                                .TradingDate = currentStockPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice + targetMultiplier * currentStockATR
                                                '.PotentialSL = Math.Min(Math.Max(currentStockSuperTrend1, Math.Max(currentStockSuperTrend2, Math.Max(currentStockSuperTrend3, currentStockSuperTrend4))), currentStockPayload.PreviousCandlePayload.Close - stopLossMultiplier * currentStockATR)
                                                .PotentialSL = Math.Max(currentStockSuperTrend1, Math.Max(currentStockSuperTrend2, Math.Max(currentStockSuperTrend3, currentStockSuperTrend4)))
                                            End With
                                            runningTrade2 = New Trade
                                            With runningTrade2
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentStockPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryTime = currentStockPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .EntryCondition = entryCondition
                                                .SignalCandle = currentStockPayload
                                                .TradingSymbol = currentStockPayload.TradingSymbol
                                                .TradingDate = currentStockPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = Decimal.MaxValue
                                                '.PotentialSL = Math.Min(Math.Max(currentStockSuperTrend1, Math.Max(currentStockSuperTrend2, Math.Max(currentStockSuperTrend3, currentStockSuperTrend4))), currentStockPayload.PreviousCandlePayload.Close - stopLossMultiplier * currentStockATR)
                                                .PotentialSL = Math.Max(currentStockSuperTrend1, Math.Max(currentStockSuperTrend2, Math.Max(currentStockSuperTrend3, currentStockSuperTrend4)))
                                            End With
                                        ElseIf signalColor = Color.Red Then
                                            runningTrade1 = New Trade
                                            With runningTrade1
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentStockPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryTime = currentStockPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .EntryCondition = entryCondition
                                                .SignalCandle = currentStockPayload
                                                .TradingSymbol = currentStockPayload.TradingSymbol
                                                .TradingDate = currentStockPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice - targetMultiplier * currentStockATR
                                                '.PotentialSL = Math.Max(Math.Min(currentStockSuperTrend1, Math.Min(currentStockSuperTrend2, Math.Min(currentStockSuperTrend3, currentStockSuperTrend4))), currentStockPayload.PreviousCandlePayload.Close + stopLossMultiplier * currentStockATR)
                                                .PotentialSL = Math.Min(currentStockSuperTrend1, Math.Min(currentStockSuperTrend2, Math.Min(currentStockSuperTrend3, currentStockSuperTrend4)))
                                            End With
                                            runningTrade2 = New Trade
                                            With runningTrade2
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentStockPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryTime = currentStockPayload.PayloadDate
                                                .EntryType = "MIS"
                                                .EntryCondition = entryCondition
                                                .SignalCandle = currentStockPayload
                                                .TradingSymbol = currentStockPayload.TradingSymbol
                                                .TradingDate = currentStockPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = Decimal.MinValue
                                                '.PotentialSL = Math.Max(Math.Min(currentStockSuperTrend1, Math.Min(currentStockSuperTrend2, Math.Min(currentStockSuperTrend3, currentStockSuperTrend4))), currentStockPayload.PreviousCandlePayload.Close + stopLossMultiplier * currentStockATR)
                                                .PotentialSL = Math.Min(currentStockSuperTrend1, Math.Min(currentStockSuperTrend2, Math.Min(currentStockSuperTrend3, currentStockSuperTrend4)))
                                            End With
                                        End If
                                    End If
                                    If totalTrade = 1 Then
                                        If runningTrade1 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade1)
                                    Else
                                        If runningTrade1 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade1)
                                        If runningTrade2 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade2)
                                    End If

                                End If
                                'Entry Trade Block End

                                entryCondition = Nothing

                                'Exit Trades Block Start
                                Dim currentPayload As Payload = tempStockPayload(potentialSignalTime)
                                Dim currentSuperTrendColor1 As Color = SupertrendColorMain1(stockName)(potentialSignalTime)
                                Dim currentSuperTrendColor2 As Color = SupertrendColorMain2(stockName)(potentialSignalTime)
                                Dim currentSuperTrendColor3 As Color = SupertrendColorMain3(stockName)(potentialSignalTime)
                                Dim currentSuperTrendColor4 As Color = SupertrendColorMain4(stockName)(potentialSignalTime)
                                Dim currentATR As Decimal = tempStockATRPayload(potentialSignalTime)
                                Dim currentSuperTrend1 As Decimal = SupertrendMain1(stockName)(potentialSignalTime)
                                Dim currentSuperTrend2 As Decimal = SupertrendMain2(stockName)(potentialSignalTime)
                                Dim currentSuperTrend3 As Decimal = SupertrendMain3(stockName)(potentialSignalTime)
                                Dim currentSuperTrend4 As Decimal = SupertrendMain4(stockName)(potentialSignalTime)

                                Dim CNCSignalColor As Color = Nothing

                                Dim tradeExit As Boolean = True
                                Dim cncExit As Boolean = True
                                Dim CNCExitSpecificTrade As List(Of Trade) = GetSpecificTrades(previousTradingDate, stockName, TradeExecutionStatus.Inprogress)
                                If CNCExitSpecificTrade IsNot Nothing AndAlso CNCExitSpecificTrade.Count > 0 Then
                                    If cncNextDayOpenPriceExitFlag Then
                                        cncNextDayOpenPriceExitFlag = False
                                        For Each cnc In CNCExitSpecificTrade
                                            MoveStopLoss(cnc, tempStockPayload(potentialSignalTime).Open)
                                            cncExit = cncExit And ExitTradeIfPossible(cnc, currentPayload, endOfDay, TypeOfStock.Commodity, False)
                                        Next
                                    Else
                                        For Each cnc In CNCExitSpecificTrade
                                            'TO DO:
                                            'Change Type of Stock
                                            cncExit = cncExit And ExitTradeIfPossible(cnc, currentPayload, endOfDay, TypeOfStock.Commodity, False)
                                        Next
                                    End If
                                End If
                                If cncExit = False Then
                                    CNCExitSpecificTrade = GetSpecificTrades(previousTradingDate, stockName, TradeExecutionStatus.Inprogress)
                                    If CNCExitSpecificTrade IsNot Nothing AndAlso CNCExitSpecificTrade.Count > 0 Then
                                        Dim CNCtradeColor As Color = If(CNCExitSpecificTrade(0).EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                        If currentSuperTrendColor1 <> CNCtradeColor OrElse currentSuperTrendColor2 <> CNCtradeColor OrElse currentSuperTrendColor3 <> CNCtradeColor OrElse currentSuperTrendColor4 <> CNCtradeColor Then
                                            For Each cnc In CNCExitSpecificTrade
                                                'Try
                                                If tempStockPayload.ContainsKey(potentialSignalTime.AddMinutes(mainSignalTimeFrame)) Then
                                                        MoveStopLoss(cnc, tempStockPayload(potentialSignalTime.AddMinutes(mainSignalTimeFrame)).Open)
                                                    Else
                                                        'MoveStopLoss(cnc, tempStockPayload(potentialSignalTime).Close)
                                                        cncNextDayOpenPriceExitFlag = True
                                                    End If
                                                'Catch ex As Exception
                                                '    Console.WriteLine(ex.ToString)
                                                'End Try
                                            Next
                                        End If
                                    End If
                                End If

                                Dim itemExitSpecificTrade As List(Of Trade) = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                    For Each item In itemExitSpecificTrade
                                        'TO DO:
                                        'Change Type of Stock
                                        tradeExit = tradeExit And ExitTradeIfPossible(item, currentPayload, endOfDay, TypeOfStock.Commodity, False)
                                    Next
                                End If
                                If tradeExit = False Then
                                    itemExitSpecificTrade = GetSpecificTrades(chk_date, stockName, TradeExecutionStatus.Inprogress)
                                    If itemExitSpecificTrade IsNot Nothing AndAlso itemExitSpecificTrade.Count > 0 Then
                                        Dim tradeColor As Color = If(itemExitSpecificTrade(0).EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                        If currentSuperTrendColor1 <> tradeColor OrElse currentSuperTrendColor2 <> tradeColor OrElse currentSuperTrendColor3 <> tradeColor OrElse currentSuperTrendColor4 <> tradeColor Then
                                            For Each item In itemExitSpecificTrade
                                                'Try
                                                If tempStockPayload.ContainsKey(potentialSignalTime.AddMinutes(mainSignalTimeFrame)) Then
                                                        MoveStopLoss(item, tempStockPayload(potentialSignalTime.AddMinutes(mainSignalTimeFrame)).Open)
                                                    End If
                                                'Catch ex As Exception
                                                '    Console.WriteLine(ex.ToString)
                                                'End Try
                                            Next
                                        End If
                                    End If
                                End If
                                'Exit Trade Block End

                                'CNC Entry Start
                                If startTime = endOfDay OrElse (startTime > endOfDay And Not IsTradeActive(chk_date, stockName) And validSignal <> previousValidSignal) Then
                                    Dim eodOfDayTrade As Trade = GetEndOfDayTrade(chk_date, stockName, TradeExecutionStatus.Close)
                                    If eodOfDayTrade IsNot Nothing AndAlso eodOfDayTrade.ExitCondition = TradeExitCondition.EndOfDay Then
                                        CNCSignalColor = If(eodOfDayTrade.EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                        If CNCSignalColor = Color.Green Then
                                            runningTrade1 = New Trade
                                            With runningTrade1
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Buy
                                                .EntryTime = currentPayload.PayloadDate
                                                .EntryType = "CNC"
                                                .EntryCondition = "CNC Conversion Entry"
                                                .SignalCandle = currentPayload
                                                .TradingSymbol = currentPayload.TradingSymbol
                                                .TradingDate = currentPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice + targetMultiplier * tempStockATRPayload(currentPayload.PreviousCandlePayload.PayloadDate)
                                                '.PotentialSL = Math.Min(Math.Max(SupertrendMain1(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Max(SupertrendMain2(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Max(SupertrendMain3(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), SupertrendMain4(stockName)(currentPayload.PreviousCandlePayload.PayloadDate)))), currentPayload.PreviousCandlePayload.Close - stopLossMultiplier * tempStockATRPayload(currentPayload.PreviousCandlePayload.PayloadDate))
                                                .PotentialSL = Math.Max(SupertrendMain1(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Max(SupertrendMain2(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Max(SupertrendMain3(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), SupertrendMain4(stockName)(currentPayload.PreviousCandlePayload.PayloadDate))))
                                            End With
                                        ElseIf CNCSignalColor = Color.Red Then
                                            runningTrade1 = New Trade
                                            With runningTrade1
                                                .TradingStatus = TradeExecutionStatus.Inprogress
                                                .EntryPrice = currentPayload.Open
                                                .EntryDirection = TradeExecutionDirection.Sell
                                                .EntryTime = currentPayload.PayloadDate
                                                .EntryType = "CNC"
                                                .EntryCondition = "CNC Conversion Entry"
                                                .SignalCandle = currentPayload
                                                .TradingSymbol = currentPayload.TradingSymbol
                                                .TradingDate = currentPayload.PayloadDate.Date
                                                .Quantity = 1
                                                .PotentialTP = .EntryPrice - targetMultiplier * tempStockATRPayload(currentPayload.PreviousCandlePayload.PayloadDate)
                                                '.PotentialSL = Math.Max(Math.Min(SupertrendMain1(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Min(SupertrendMain2(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Min(SupertrendMain3(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), SupertrendMain4(stockName)(currentPayload.PreviousCandlePayload.PayloadDate)))), currentPayload.PreviousCandlePayload.Close + stopLossMultiplier * tempStockATRPayload(currentPayload.PreviousCandlePayload.PayloadDate))
                                                .PotentialSL = Math.Min(SupertrendMain1(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Min(SupertrendMain2(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), Math.Min(SupertrendMain3(stockName)(currentPayload.PreviousCandlePayload.PayloadDate), SupertrendMain4(stockName)(currentPayload.PreviousCandlePayload.PayloadDate))))
                                            End With
                                        End If
                                        If runningTrade1 IsNot Nothing Then EnterOrder(chk_date, stockName, runningTrade1)
                                        previousTradingDate = chk_date
                                    End If
                                End If
                                'CNC Entry End

                                'Signal Change Block Start
                                Dim lastTradeColor As Color = Nothing
                                Dim lastTrade As Trade = GetLastTrade(chk_date, stockName, TradeExecutionStatus.Close)
                                If lastTrade IsNot Nothing Then
                                    lastTradeColor = If(lastTrade.EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                Else
                                    lastTrade = GetLastTrade(previousTradingDate, stockName, TradeExecutionStatus.Close)
                                    If lastTrade IsNot Nothing Then
                                        lastTradeColor = If(lastTrade.EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                    Else
                                        lastTrade = GetLastTrade(previousTradingDate, stockName, TradeExecutionStatus.Inprogress)
                                        If lastTrade IsNot Nothing Then
                                            lastTradeColor = If(lastTrade.EntryDirection = TradeExecutionDirection.Buy, Color.Green, Color.Red)
                                        End If
                                    End If
                                End If
                                If lastTradeColor = Color.Green Then
                                    If currentSuperTrendColor1 = lastTradeColor AndAlso currentSuperTrendColor2 = lastTradeColor AndAlso currentSuperTrendColor3 = lastTradeColor AndAlso currentSuperTrendColor4 = lastTradeColor Then
                                        If currentPayload.Low < Math.Max(currentSuperTrend1, Math.Max(currentSuperTrend2, Math.Max(currentSuperTrend3, currentSuperTrend4))) Then
                                            previousValidSignal = False
                                            entryCondition = "Supertrend Cut"
                                        End If
                                    End If
                                ElseIf lastTradeColor = Color.Red Then
                                    If currentSuperTrendColor1 = lastTradeColor AndAlso currentSuperTrendColor2 = lastTradeColor AndAlso currentSuperTrendColor3 = lastTradeColor AndAlso currentSuperTrendColor4 = lastTradeColor Then
                                        If currentPayload.High > Math.Min(currentSuperTrend1, Math.Min(currentSuperTrend2, Math.Min(currentSuperTrend3, currentSuperTrend4))) Then
                                            previousValidSignal = False
                                            entryCondition = "Supertrend Cut"
                                        End If
                                    End If
                                End If
                                'Signal Change Block End
                            End If
                        Next
                        startTime = startTime.AddMinutes(5)
                    End While
                End If
            End If
            chk_date = chk_date.AddDays(1)
        End While
        For Each stock In stockList.Keys
            Dim curDate As Date = System.DateTime.Now
            Dim filename As String = String.Format("Multi Timeframe Supertrend {0} {1}-{2}-{3}_{4}-{5}-{6}", stock,
                                                   curDate.Year, curDate.Month, curDate.Day, curDate.Hour, curDate.Minute, curDate.Second)
            PrintArrayToExcel(filename, stock)
        Next
    End Sub

    Private Function GetSupertrendSignal(ByVal signalTime As Date, ByVal signalColor As Color, ParamArray supertrends() As Object) As Boolean
        Dim ret As Boolean = True
        For i As Integer = 0 To supertrends.Count - 1 Step 2
            'Try
            If supertrends(i + 1) Then
                    ret = ret And supertrends(i)(signalTime) = signalColor
                End If
            'Catch ex As Exception
            '    Console.WriteLine(ex.ToString)
            'End Try
        Next
        Return ret
    End Function
    Private Function GetMACDSignal(ByVal signalTime As Date, ByVal signalColor As Color, ParamArray macd() As Object) As Boolean
        Dim ret As Boolean = True
        For i As Integer = 0 To macd.Count - 1 Step 2
            'Try
            If macd(i + 1) Then
                    If signalColor = Color.Green Then
                        ret = ret And macd(i)(signalTime) > 0
                    ElseIf signalColor = Color.Red Then
                        ret = ret And macd(i)(signalTime) < 0
                    End If
                End If
            'Catch ex As Exception
            '    Console.WriteLine(ex.ToString)
            'End Try
        Next
        Return ret
    End Function
    Public Enum IndicatorTimeFrame
        TF5 = 1
        TF15
        TF60
        TF1D
        TF5_ATR_10
        TF5_MACD_12_26_9
        TF5_Supertrend_7_3
        TF5_Supertrend_10_3
        TF5_Supertrend_11_2
        TF5_Supertrend_144_3
        TF15_Supertrend_7_3
        TF15_Supertrend_10_3
        TF15_Supertrend_11_2
        TF15_Supertrend_144_3
        TF60_Supertrend_7_3
        TF60_Supertrend_10_3
        TF60_Supertrend_11_2
        TF60_Supertrend_144_3
        TF1D_Supertrend_7_3
        TF1D_Supertrend_10_3
        TF1D_Supertrend_11_2
        TF1D_Supertrend_144_3
    End Enum
End Class
