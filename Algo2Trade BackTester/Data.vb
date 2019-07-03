Imports Algo2TradeBLL
Module Data
    Public PastIntradayData As Dictionary(Of String, Dictionary(Of Date, Payload))
    Public PastEODData As Dictionary(Of String, Dictionary(Of Date, Payload))
    Public PastData As Dictionary(Of String, Dictionary(Of Date, Payload))
    Public IndicatorData1 As Dictionary(Of String, Dictionary(Of Date, Decimal))
    Public IndicatorData2 As Dictionary(Of String, Dictionary(Of Date, Decimal))
End Module
