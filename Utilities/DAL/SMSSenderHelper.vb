Imports System.Net.Http
Imports System.Threading
Imports Utilities.Network

Namespace DAL
    Public Class SMSSenderHelper

        Private _apikey As String
        Protected _canceller As CancellationTokenSource

        Public Sub New(ByVal apiKey As String, ByVal canceller As CancellationTokenSource)
            _apikey = apiKey
            _canceller = canceller
        End Sub

        Public Async Function sendSMSGet(ByVal url As String, ByVal message As String, ByVal numbers As String, ByVal sender As String) As Task
            Dim proxyToBeUsed As HttpProxy = Nothing
            Dim ret As List(Of String) = Nothing

            Using browser As New HttpBrowser(proxyToBeUsed, Net.DecompressionMethods.GZip, New TimeSpan(0, 1, 0), _canceller)
                'AddHandler browser.DocumentDownloadComplete, AddressOf OnDocumentDownloadCompete
                'AddHandler browser.Heartbeat, AddressOf OnHeartbeat
                'AddHandler browser.WaitingFor, AddressOf OnWaitingFor
                'AddHandler browser.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                'Get to the landing page first
                Dim l As Tuple(Of Uri, Object) = Await browser.NonPOSTRequestAsync("https://api.textlocal.in/bulk_json",
                                                                                 HttpMethod.Get,
                                                                                 Nothing,
                                                                                 True,
                                                                                 Nothing,
                                                                                 False,
                                                                                 Nothing).ConfigureAwait(False)
                If l Is Nothing OrElse l.Item2 Is Nothing Then
                    Throw New ApplicationException(String.Format("No response in the additional site's historical race results landing page: {0}", "https://zerodha.com/static/app.js"))
                End If
                'RaiseEvent Heartbeat("Parsing additional site's...")
                If l IsNot Nothing AndAlso l.Item2 IsNot Nothing Then
                    Dim counter As Integer = 0
                    For Each item In l.Item2
                        If counter = 1 Then
                            Dim jString As KeyValuePair(Of String, String) = item
                        End If
                        counter += 1
                    Next
                    'Dim jString As Dictionary(Of String, String) = l.Item2
                    'If jString IsNot Nothing Then
                    'Dim map As String = Utilities.Strings.GetTextBetween("COMMODITY_MULTIPLIER_MAP=", "},", jString)
                    'If map IsNot Nothing Then
                    '    map = map & "}"
                    '    Dim retDictionary = Utilities.Strings.JsonDeserialize(map)
                    'End If

                    'RaiseEvent Heartbeat("Parsing additional site's landing page post contents...")
                    'ret = GetAdditionalSitesAllMeetingsURL(htmlDoc)
                    'End If
                End If
            End Using
            'Return ret
            '    Dim apikey = "4Zmp70fpTeg-odYWEiw1okZ7OXEYJaY2sYWBY481qH"
            '    Dim message = "Buying alert on CRUDEOILM18SEPFUT at 10-Sep-2018 20:01"
            '    Dim numbers = "919874795959"
            '    Dim strGet As String
            '    Dim sender = "TXTLCL"
            '    Dim url As String = "https://api.textlocal.in/send/?"

            '    strGet = url + "apikey=" + apikey _
            '    + "&numbers=" + numbers _
            '    + "&message=" + WebUtility.UrlEncode(message) _
            '    + "&sender=" + sender

            '    Dim webClient As New System.Net.WebClient
            '    Dim result As String = webClient.DownloadString(strGet)
            '    Console.WriteLine(result)
            '    Return result
        End Function
    End Class
End Namespace
