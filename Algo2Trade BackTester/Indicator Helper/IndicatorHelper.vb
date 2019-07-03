Imports MySql.Data.MySqlClient
Imports Algo2TradeBLL
Imports System.Threading

Public Class IndicatorHelper
    Implements IDisposable
#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

#Region "Enum"
    Public Enum InstrumentCategory
        Cash = 0
        Currency
        Commodity
        Futures
    End Enum
#End Region

#Region "Variables"
    Private _name As String
    Private _category As InstrumentCategory
    Private _open As Boolean
    Private _close As Boolean
    Private _high As Boolean
    Private _low As Boolean
    Private _volume As Boolean
    Private _canceller As CancellationTokenSource
    Dim cmn As Common = New Common(_canceller)
#End Region

#Region "Constructor"
    Public Sub New(ByVal name As String, ByVal category As InstrumentCategory, ByVal open As Boolean, ByVal close As Boolean, ByVal high As Boolean, ByVal low As Boolean, ByVal volume As Boolean, ByVal canceller As CancellationTokenSource)
        _name = name
        _category = category
        _open = open
        _low = low
        _high = high
        _close = close
        _volume = volume
        _canceller = canceller
    End Sub
#End Region

#Region "Private Functions"
    Private Function FetchData(name As String, tr_date As DateTime) As Dictionary(Of Date, Payload)
        AddHandler cmn.Heartbeat, AddressOf OnHeartbeat
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Select Case _category
            Case InstrumentCategory.Cash
                inputpayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Cash, name, tr_date.AddDays(-7), tr_date)
            Case InstrumentCategory.Currency
                inputpayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Currency, name, tr_date.AddDays(-7), tr_date)
            Case InstrumentCategory.Commodity
                inputpayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Commodity, name, tr_date.AddDays(-7), tr_date)
            Case InstrumentCategory.Futures
                inputpayload = cmn.GetRawPayload(Common.DataBaseTable.Intraday_Futures, name, tr_date.AddDays(-7), tr_date)
        End Select
        Return inputpayload
    End Function
#End Region

#Region "Public Functions"
    Public Function VWAP_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo vwap_label
        End If
        Dim vwappayload As New Dictionary(Of Date, Decimal)
        Indicator.VWAP.CalculateVWAP(inputpayload, vwappayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("VWAP")
        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("VWAP") = vwappayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
vwap_label: Return outputdt
    End Function
    Public Function Fractals_Calculate(fractalsPeriod As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo fractals_label
        End If
        Dim fractalshighpayload As New Dictionary(Of Date, Decimal)
        Dim fractalslowpayload As New Dictionary(Of Date, Decimal)
        Indicator.Fractals.CalculateFractal(fractalsPeriod, inputpayload, fractalshighpayload, fractalslowpayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Fractals High")
        outputdt.Columns.Add("Fractals Low")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Fractals High") = fractalshighpayload(tempKeys)
                row("Fractals Low") = fractalslowpayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
fractals_label: Return outputdt
    End Function
    Public Function SMA_Calculate(smaperiod As Integer, smafield As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo SMA_label
        End If

        Dim smapayload As New Dictionary(Of Date, Decimal)
        Dim f As Integer = smafield
        Select Case f
            Case 0
                Indicator.SMA.CalculateSMA(smaperiod, Payload.PayloadFields.Open, inputpayload, smapayload)
            Case 1
                Indicator.SMA.CalculateSMA(smaperiod, Payload.PayloadFields.Low, inputpayload, smapayload)
            Case 2
                Indicator.SMA.CalculateSMA(smaperiod, Payload.PayloadFields.High, inputpayload, smapayload)
            Case 3
                Indicator.SMA.CalculateSMA(smaperiod, Payload.PayloadFields.Close, inputpayload, smapayload)
            Case 4
                Indicator.SMA.CalculateSMA(smaperiod, Payload.PayloadFields.Volume, inputpayload, smapayload)
        End Select

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Simple Moving Avarage")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Simple Moving Avarage") = smapayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
SMA_label: Return outputdt
    End Function
    Public Function EMA_Calculate(emaperiod As Integer, emafield As Integer, tr_date As DateTime)
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo EMA_label
        End If
        Dim emapayload As New Dictionary(Of Date, Decimal)
        Dim f As Integer = emafield
        Select Case f
            Case 0
                Indicator.EMA.CalculateEMA(emaperiod, Payload.PayloadFields.Open, inputpayload, emapayload)
            Case 1
                Indicator.EMA.CalculateEMA(emaperiod, Payload.PayloadFields.Low, inputpayload, emapayload)
            Case 2
                Indicator.EMA.CalculateEMA(emaperiod, Payload.PayloadFields.High, inputpayload, emapayload)
            Case 3
                Indicator.EMA.CalculateEMA(emaperiod, Payload.PayloadFields.Close, inputpayload, emapayload)
            Case 4
                Indicator.EMA.CalculateEMA(emaperiod, Payload.PayloadFields.Volume, inputpayload, emapayload)
        End Select

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Exponential Moving Average")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Exponential Moving Average") = emapayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
EMA_label: Return outputdt
    End Function
    Public Function SMI_Calculate(K As Integer, Ks As Integer, Kds As Integer, D As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo SMI_label
        End If

        Dim xminutepayload As New Dictionary(Of Date, Payload)
        Dim SMIsignalpayload As New Dictionary(Of Date, Decimal)
        Dim SMIEMAsignalpayload As New Dictionary(Of Date, Decimal)

        xminutepayload = cmn.ConvertPayloadsToXMinutes(inputpayload, 15)
        Indicator.SMI.CalculateSMI(K, Ks, Kds, D, xminutepayload, SMIsignalpayload, SMIEMAsignalpayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("SMI Signal")
        outputdt.Columns.Add("SMI_EMA Signal")

        For Each tempKeys In xminutepayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = xminutepayload(tempKeys).PayloadDate
                row("Trading Symbol") = xminutepayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = xminutepayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = xminutepayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = xminutepayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = xminutepayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = xminutepayload(tempKeys).Volume
                End If
                row("SMI Signal") = SMIsignalpayload(tempKeys)
                row("SMI_EMA Signal") = SMIEMAsignalpayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
SMI_label: Return outputdt
    End Function
    Public Function ATR_Calculate(ATRPeriod As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo ATR_label
        End If

        Dim ATRpayload As New Dictionary(Of Date, Decimal)
        Indicator.ATR.CalculateATR(ATRPeriod, inputpayload, ATRpayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("ATR")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("ATR") = ATRpayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
ATR_label: Return outputdt
    End Function
    Public Function Core_Indicator(name As String, tr_date As DateTime) As Dictionary(Of Date, Payload)
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim OneMinutePayload As Dictionary(Of Date, Payload) = Nothing
        inputPayload = FetchData(name, tr_date)
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            OneMinutePayload = New Dictionary(Of Date, Payload)

            For Each tempKeys In inputPayload.Keys
                'If tempKeys.Date = tr_date.Date Then
                OneMinutePayload.Add(tempKeys, inputPayload(tempKeys))
                'End If
            Next
        End If
        Return OneMinutePayload
    End Function
    Public Function CCI_Calculate(cciperiod As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo CCI_label
        End If

        Dim CCIoutput As New Dictionary(Of Date, Decimal)
        Indicator.CCI.CalculateCCI(cciperiod, inputpayload, CCIoutput)
        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Commodity Channel Index")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Commodity Channel Index") = CCIoutput(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
CCI_label: Return outputdt
    End Function
    Public Function RSI_Calculate(rsiperiod As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo RSI_label
        End If

        Dim RSIoutput As New Dictionary(Of Date, Decimal)
        Indicator.RSI.CalculateRSI(rsiperiod, inputpayload, RSIoutput)
        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Relative Strength Index")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Relative Strength Index") = RSIoutput(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
RSI_label: Return outputdt
    End Function
    Public Function MACD_Calculate(fastMA As Integer, slowMA As Integer, signal As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo MACD_label
        End If

        Dim MACDPayload As Dictionary(Of Date, Decimal) = Nothing
        Dim signalPayload As Dictionary(Of Date, Decimal) = Nothing
        Dim histogramPayload As Dictionary(Of Date, Decimal) = Nothing

        Indicator.MACD.CalculateMACD(fastMA, slowMA, signal, inputpayload, MACDPayload, signalPayload, histogramPayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("MACD")
        outputdt.Columns.Add("Signal")
        outputdt.Columns.Add("Histogram")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("MACD") = MACDPayload(tempKeys)
                row("Signal") = signalPayload(tempKeys)
                row("Histogram") = histogramPayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
MACD_label: Return outputdt
    End Function
    Public Function HikenAshiConversion(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo HA_label
        End If

        Dim outputHAPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputpayload, outputHAPayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("HA Open")
        outputdt.Columns.Add("HA Low")
        outputdt.Columns.Add("HA High")
        outputdt.Columns.Add("HA Close")
        outputdt.Columns.Add("HA Volume")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("HA Open") = outputHAPayload(tempKeys).Open
                row("HA Low") = outputHAPayload(tempKeys).Low
                row("HA High") = outputHAPayload(tempKeys).High
                row("HA Close") = outputHAPayload(tempKeys).Close
                row("HA Volume") = outputHAPayload(tempKeys).Volume
                outputdt.Rows.Add(row)
            End If
        Next
HA_label: Return outputdt
    End Function
    Public Function Supertrend_Calculate(period As Integer, multiplier As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo Supertrend_label
        End If

        Dim supertrendPayload As Dictionary(Of Date, Decimal) = Nothing
        Dim supertrendColor As Dictionary(Of Date, Color) = Nothing
        Indicator.Supertrend.CalculateSupertrend(period, multiplier, inputpayload, supertrendPayload, supertrendColor)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("SUPERTREND")
        outputdt.Columns.Add("Signal")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("SUPERTREND") = supertrendPayload(tempKeys)
                row("Signal") = If(supertrendColor(tempKeys) = Color.Red, "SELL", "BUY")
                outputdt.Rows.Add(row)
            End If
        Next
Supertrend_label: Return outputdt
    End Function
    Public Function ATRTrailingStop_Calculate(ATRPeriod As Integer, Multiplier As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo ATRTrailingStop_label
        End If

        Dim ATRTrailingStopPayload As Dictionary(Of Date, Decimal) = Nothing
        Dim ATRTrailingStopColorPayload As Dictionary(Of Date, Color) = Nothing
        Indicator.ATRTrailingStop.CalculateATRTrailingStop(ATRPeriod, Multiplier, inputpayload, ATRTrailingStopPayload, ATRTrailingStopColorPayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("ATR Trailing Stop")
        outputdt.Columns.Add("ATR Trailing Stop Color")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("ATR Trailing Stop") = ATRTrailingStopPayload(tempKeys)
                row("ATR Trailing Stop Color") = ATRTrailingStopColorPayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
ATRTrailingStop_label: Return outputdt
    End Function
    Public Function Fractal1_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim outputHAPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputpayload, outputHAPayload)

        inputpayload = Nothing
        inputpayload = outputHAPayload

        Dim outputPayload As Dictionary(Of Date, Integer) = Nothing
        Fractal1.CalculateFractal1(50, outputHAPayload, outputPayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Signal")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Signal") = outputPayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function FractalWithSMA_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim outputHAPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputpayload, outputHAPayload)

        inputpayload = Nothing
        inputpayload = outputHAPayload

        Dim outputSignalPayload As Dictionary(Of Date, Color) = Nothing
        Dim outputEntryPricePayload As Dictionary(Of Date, Double) = Nothing
        Dim outputStoplossPricePayload As Dictionary(Of Date, Double) = Nothing
        FractalWithSMA.CalculateFractalWithSMA(50, outputHAPayload, outputSignalPayload, outputEntryPricePayload, outputStoplossPricePayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Signal Direction")
        outputdt.Columns.Add("Entry Price")
        outputdt.Columns.Add("Stoploss Price")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Signal Direction") = If(outputSignalPayload(tempKeys) = Color.Green, "Buy", If(outputSignalPayload(tempKeys) = Color.Red, "Sell", "None"))
                row("Entry Price") = outputEntryPricePayload(tempKeys)
                row("Stoploss Price") = outputStoplossPricePayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function FractalRetracementEntry_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim outputHAPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputpayload, outputHAPayload)

        inputpayload = Nothing
        inputpayload = outputHAPayload

        Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
        Dim outputEntryPricePayload As Dictionary(Of Date, Double) = Nothing
        StrategyRules.FractalRetracementEntryRule.CalculateFractalRetracementEntry(50, outputHAPayload, outputSignalPayload, outputEntryPricePayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Signal Direction")
        outputdt.Columns.Add("Entry Price")

        For Each tempKeys In inputpayload.Keys
            Dim temp_date As String = tempKeys.ToString().Substring(0, tempKeys.ToString().IndexOf(" "))
            Dim t_date As String = tr_date.ToString().Substring(0, tr_date.ToString().IndexOf(" "))
            If temp_date = t_date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Signal Direction") = outputSignalPayload(tempKeys)
                row("Entry Price") = outputEntryPricePayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function ATRBands_Calculate(ATRShift As Integer, ATRPeriod As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim ATRBandHighpayload As New Dictionary(Of Date, Decimal)
        Dim ATRBandLowpayload As New Dictionary(Of Date, Decimal)
        Indicator.ATRBands.CalculateATRBands(ATRShift, ATRPeriod, Payload.PayloadFields.Close, inputpayload, ATRBandHighpayload, ATRBandLowpayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("ATR High Band")
        outputdt.Columns.Add("ATR Low Band")

        For Each tempKeys In inputpayload.Keys
            If tempKeys.Date = tr_date.Date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("ATR High Band") = ATRBandHighpayload(tempKeys)
                row("ATR Low Band") = ATRBandLowpayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function BollingerBands_Calculate(period As Integer, sdMultiplier As Integer, tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim bandHighpayload As New Dictionary(Of Date, Decimal)
        Dim bandLowpayload As New Dictionary(Of Date, Decimal)
        Dim SMApayload As New Dictionary(Of Date, Decimal)
        Indicator.BollingerBands.CalculateBollingerBands(period, Payload.PayloadFields.Close, sdMultiplier, inputpayload, bandHighpayload, bandLowpayload, SMApayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("High Band")
        outputdt.Columns.Add("Low Band")

        For Each tempKeys In inputpayload.Keys
            If tempKeys.Date = tr_date.Date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("High Band") = bandHighpayload(tempKeys)
                row("Low Band") = bandLowpayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function Renko_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim outputPayload As New Dictionary(Of String, Payload)
        Indicator.Algo2TradeRenko.ConvertToRenko(0.7, inputpayload, outputPayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        outputdt.Columns.Add("Renko Open")
        outputdt.Columns.Add("Renko Low")
        outputdt.Columns.Add("Renko High")
        outputdt.Columns.Add("Renko Close")
        outputdt.Columns.Add("Renko Color")

        For Each tempKeys In outputPayload.Keys
            If outputPayload(tempKeys).PayloadDate.Date = tr_date.Date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = outputPayload(tempKeys).PayloadDate
                row("Trading Symbol") = outputPayload(tempKeys).TradingSymbol
                row("Renko Open") = outputPayload(tempKeys).Open
                row("Renko Low") = outputPayload(tempKeys).Low
                row("Renko High") = outputPayload(tempKeys).High
                row("Renko Close") = outputPayload(tempKeys).Close
                row("Renko Color") = outputPayload(tempKeys).CandleColor

                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function NaughtyBoyVWAP_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload Is Nothing Then
            GoTo label
        End If

        Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
        Dim outputEntryPricePayload As Dictionary(Of Date, Double) = Nothing
        Dim outputStoplossPricePayload As Dictionary(Of Date, Double) = Nothing
        StrategyRules.NaughtyBoyVWAPRule.CalculateNaughtyBoyVWAPRule(inputpayload, outputSignalPayload, outputEntryPricePayload, outputStoplossPricePayload)

        outputdt = New DataTable
        outputdt.Columns.Add("PayloadDate")
        outputdt.Columns.Add("Trading Symbol")
        If _open = True Then
            outputdt.Columns.Add("Open")
        End If
        If _low = True Then
            outputdt.Columns.Add("Low")
        End If
        If _high = True Then
            outputdt.Columns.Add("High")
        End If
        If _close = True Then
            outputdt.Columns.Add("Close")
        End If
        If _volume = True Then
            outputdt.Columns.Add("Volume")
        End If
        outputdt.Columns.Add("Signal Direction")
        outputdt.Columns.Add("Entry Price")
        outputdt.Columns.Add("Stoploss Price")

        For Each tempKeys In inputpayload.Keys
            If tempKeys.Date = tr_date.Date Then
                Dim row As DataRow = outputdt.NewRow
                row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                If _open = True Then
                    row("Open") = inputpayload(tempKeys).Open
                End If
                If _low = True Then
                    row("Low") = inputpayload(tempKeys).Low
                End If
                If _high = True Then
                    row("High") = inputpayload(tempKeys).High
                End If
                If _close = True Then
                    row("Close") = inputpayload(tempKeys).Close
                End If
                If _volume = True Then
                    row("Volume") = inputpayload(tempKeys).Volume
                End If
                row("Signal Direction") = outputSignalPayload(tempKeys)
                row("Entry Price") = outputEntryPricePayload(tempKeys)
                row("Stoploss Price") = outputStoplossPricePayload(tempKeys)
                outputdt.Rows.Add(row)
            End If
        Next
label:  Return outputdt
    End Function
    Public Function VWAPDoubleConfirmation_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload IsNot Nothing AndAlso inputpayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputStoplossPricePayload As Dictionary(Of Date, Double) = Nothing
            StrategyRules.VWAPDoubleConfirmationRule.CalculateNaughtyBoyVWAPRule(1, 14, inputpayload, outputSignalPayload, outputStoplossPricePayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Signal Direction")
            outputdt.Columns.Add("Stoploss Price")

            For Each tempKeys In inputpayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputpayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputpayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputpayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputpayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputpayload(tempKeys).Volume
                    End If
                    row("Signal Direction") = outputSignalPayload(tempKeys)
                    row("Stoploss Price") = outputStoplossPricePayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function InTheTrend_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload IsNot Nothing AndAlso inputpayload.Count > 0 Then
            Dim outputHighEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputLowEntryPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputVWAPPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            Dim outputPearsingCandleSignalPayload As Dictionary(Of Date, Integer) = Nothing
            StrategyRules.InTheTrendRule.CalculateInTheTrendRule(1, 5, inputpayload, outputVWAPPayload, outputHighEntryPayload, outputLowEntryPayload, outputSignalPayload, outputPearsingCandleSignalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("VWAP")
            outputdt.Columns.Add("Long Entry")
            outputdt.Columns.Add("Short Entry")
            outputdt.Columns.Add("Signal")
            outputdt.Columns.Add("Pearsing Candle Signal")

            For Each tempKeys In inputpayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputpayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputpayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputpayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputpayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputpayload(tempKeys).Volume
                    End If
                    row("VWAP") = outputVWAPPayload(tempKeys)
                    row("Long Entry") = outputHighEntryPayload(tempKeys)
                    row("Short Entry") = outputLowEntryPayload(tempKeys)
                    row("Signal") = outputSignalPayload(tempKeys)
                    row("Pearsing Candle Signal") = outputPearsingCandleSignalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function InTheTrendPearsing_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload IsNot Nothing AndAlso inputpayload.Count > 0 Then
            Dim outputUpperBandPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputLowerBandPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            StrategyRules.InTheTrendPearsingRule.CalculateInTheTrendPearsingRule(1, 5, inputpayload, outputUpperBandPayload, outputLowerBandPayload, outputSignalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Upper Band")
            outputdt.Columns.Add("Lower Band")
            outputdt.Columns.Add("Signal")

            For Each tempKeys In inputpayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputpayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputpayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputpayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputpayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputpayload(tempKeys).Volume
                    End If
                    row("Upper Band") = outputUpperBandPayload(tempKeys)
                    row("Lower Band") = outputLowerBandPayload(tempKeys)
                    row("Signal") = outputSignalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function SwingHighLow_Calculate(tr_date As DateTime) As DataTable
        Dim inputpayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputpayload = FetchData(_name, tr_date)
        If inputpayload IsNot Nothing AndAlso inputpayload.Count > 0 Then
            Dim outputHighPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim outputLowPayload As Dictionary(Of Date, Decimal) = Nothing
            Indicator.SwingHighLow.CalculateSwingHighLow(inputpayload, False, outputHighPayload, outputLowPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Swing High")
            outputdt.Columns.Add("Swing Low")

            For Each tempKeys In inputpayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputpayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputpayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputpayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputpayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputpayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputpayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputpayload(tempKeys).Volume
                    End If
                    row("Swing High") = outputHighPayload(tempKeys)
                    row("Swing Low") = outputLowPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function DoubleU_Calculate(tr_date As DateTime) As DataTable
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputPayload = FetchData(_name, tr_date)
        Dim inputHKPayload As Dictionary(Of Date, Payload) = Nothing
        Indicator.HeikenAshi.ConvertToHeikenAshi(inputPayload, inputHKPayload)
        inputPayload = Nothing
        inputPayload = inputHKPayload
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            StrategyRules.DoubleURule.CalculateDoubleURule(inputPayload, outputSignalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Signal")

            For Each tempKeys In inputPayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputPayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputPayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputPayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputPayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputPayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputPayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputPayload(tempKeys).Volume
                    End If
                    row("Signal") = outputSignalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function MRWithOneFractalAndMA_Calculate(tr_date As DateTime) As DataTable
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputPayload = FetchData(_name, tr_date)
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim outputSignalPayload As Dictionary(Of Date, Integer) = Nothing
            StrategyRules.MRWithOneFractalAndMA.CalculateMRWithOneFractalAndMA(inputPayload, outputSignalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Signal")

            For Each tempKeys In inputPayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = inputPayload(tempKeys).PayloadDate
                    row("Trading Symbol") = inputPayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = inputPayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = inputPayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = inputPayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = inputPayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = inputPayload(tempKeys).Volume
                    End If
                    row("Signal") = outputSignalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function JOYMA4_NowIShouldGo_Calculate(tr_date As DateTime) As DataTable
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputPayload = FetchData(_name, tr_date)
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim XMinutePayload As Dictionary(Of Date, Payload) = cmn.ConvertPayloadsToXMinutes(inputPayload, 15)

            Dim buyPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim sellPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim typePayload As Dictionary(Of Date, Integer) = Nothing
            Dim buyTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim sellTargetPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim remarksPayload As Dictionary(Of Date, String) = Nothing
            OnHeartbeat("Calculating")
            StrategyRules.JOYMA4_NowIShouldGo.CalculateJOYMA4_NowIShouldGo(XMinutePayload, buyPayload, sellPayload, typePayload, buyTargetPayload, sellTargetPayload, remarksPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Buy")
            outputdt.Columns.Add("Sell")
            outputdt.Columns.Add("Type")
            outputdt.Columns.Add("Buy Target")
            outputdt.Columns.Add("Sell Target")
            outputdt.Columns.Add("Remarks")

            For Each tempKeys In XMinutePayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = XMinutePayload(tempKeys).PayloadDate
                    row("Trading Symbol") = XMinutePayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = XMinutePayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = XMinutePayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = XMinutePayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = XMinutePayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = XMinutePayload(tempKeys).Volume
                    End If
                    row("Buy") = buyPayload(tempKeys)
                    row("Sell") = sellPayload(tempKeys)
                    row("Type") = typePayload(tempKeys)
                    row("Buy Target") = buyTargetPayload(tempKeys)
                    row("Sell Target") = sellTargetPayload(tempKeys)
                    row("Remarks") = remarksPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function HigherHighHigherLowPattern_Calculate(tr_date As DateTime) As DataTable
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputPayload = FetchData(_name, tr_date)
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim XMinutePayload As Dictionary(Of Date, Payload) = cmn.ConvertPayloadsToXMinutes(inputPayload, 15)

            Dim signalPayload As Dictionary(Of Date, Integer) = Nothing
            OnHeartbeat("Calculating")
            StrategyRules.HigherHighHigherLowPattern.CalculateHigherHighHigherLowPattern(XMinutePayload, signalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("Signal")

            For Each tempKeys In XMinutePayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = XMinutePayload(tempKeys).PayloadDate
                    row("Trading Symbol") = XMinutePayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = XMinutePayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = XMinutePayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = XMinutePayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = XMinutePayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = XMinutePayload(tempKeys).Volume
                    End If
                    row("Signal") = signalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
    Public Function TII_Calculate(tr_date As DateTime) As DataTable
        Dim inputPayload As Dictionary(Of Date, Payload) = Nothing
        Dim outputdt As DataTable = Nothing
        inputPayload = FetchData(_name, tr_date)
        If inputPayload IsNot Nothing AndAlso inputPayload.Count > 0 Then
            Dim XMinutePayload As Dictionary(Of Date, Payload) = inputPayload

            Dim signalPayload As Dictionary(Of Date, Decimal) = Nothing
            Dim TIIPayload As Dictionary(Of Date, Decimal) = Nothing
            OnHeartbeat("Calculating")
            Indicator.TrendIntensityIndex.CalculateTII(Payload.PayloadFields.Close, 14, 9, XMinutePayload, TIIPayload, signalPayload)

            outputdt = New DataTable
            outputdt.Columns.Add("PayloadDate")
            outputdt.Columns.Add("Trading Symbol")
            If _open = True Then
                outputdt.Columns.Add("Open")
            End If
            If _low = True Then
                outputdt.Columns.Add("Low")
            End If
            If _high = True Then
                outputdt.Columns.Add("High")
            End If
            If _close = True Then
                outputdt.Columns.Add("Close")
            End If
            If _volume = True Then
                outputdt.Columns.Add("Volume")
            End If
            outputdt.Columns.Add("TII")
            outputdt.Columns.Add("Signal")

            For Each tempKeys In XMinutePayload.Keys
                If tempKeys.Date = tr_date.Date Then
                    Dim row As DataRow = outputdt.NewRow
                    row("PayloadDate") = XMinutePayload(tempKeys).PayloadDate
                    row("Trading Symbol") = XMinutePayload(tempKeys).TradingSymbol
                    If _open = True Then
                        row("Open") = XMinutePayload(tempKeys).Open
                    End If
                    If _low = True Then
                        row("Low") = XMinutePayload(tempKeys).Low
                    End If
                    If _high = True Then
                        row("High") = XMinutePayload(tempKeys).High
                    End If
                    If _close = True Then
                        row("Close") = XMinutePayload(tempKeys).Close
                    End If
                    If _volume = True Then
                        row("Volume") = XMinutePayload(tempKeys).Volume
                    End If
                    row("TII") = TIIPayload(tempKeys)
                    row("Signal") = signalPayload(tempKeys)
                    outputdt.Rows.Add(row)
                End If
            Next
        End If
        Return outputdt
    End Function
#End Region

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class
