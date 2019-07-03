Imports System.Text.RegularExpressions
Imports NLog
Imports System.Net.WebUtility
Imports System.Text
Imports System.Runtime.CompilerServices
Imports System.ComponentModel
Imports System.Globalization
Imports System.Security.Cryptography
Imports System.Web.Script.Serialization
Imports System.IO

Namespace Strings
    Public Module StringManipulation
#Region "Logging and Status Progress"
        Public logger As Logger = LogManager.GetCurrentClassLogger
#End Region

#Region "Enums"
        Public Enum StringMatchoptions
            MatchFullWord
            MatchPartialWord
        End Enum
        Public Enum HyperLinkType
            Internal = 1
            External
            Both
        End Enum
#End Region

#Region "Private Attributes"
#End Region

#Region "Public Attribites"
#End Region

#Region "Private Methods"
#End Region

#Region "Public Methods"
        Public Function JsonSerialize(x As Object) As String
            Dim jss = New JavaScriptSerializer()
            Return jss.Serialize(x)
        End Function
        Public Function JsonDeserialize(Json As String) As Dictionary(Of String, Object)
            Dim jss = New JavaScriptSerializer()
            Dim dict As Dictionary(Of String, Object) = jss.Deserialize(Of Dictionary(Of String, Object))(Json)
            Return dict
        End Function
        Public Function SHA256(Data As String) As String
            Dim sha256__1 As New SHA256Managed()
            Dim hexhash As New StringBuilder()
            Dim hash As Byte() = sha256__1.ComputeHash(Encoding.UTF8.GetBytes(Data), 0, Encoding.UTF8.GetByteCount(Data))
            For Each b As Byte In hash
                hexhash.Append(b.ToString("x2"))
            Next
            Return hexhash.ToString()
        End Function
        Public Function GetTextBetween(ByVal firstText As String, ByVal secondText As String, ByVal searchFromText As String) As String
            Dim regex As New Regex(firstText & "(.*?)" & secondText, RegexOptions.Singleline) 'SingleLine options makes it return multilined text
            Dim match As Match = regex.Match(searchFromText)
            Return match.Groups(1).Value
        End Function
        Public Function StripTags(ByVal html As String) As String
            ' Remove HTML tags.
            Return Regex.Replace(html, "<.*?>", "")
        End Function
        Public Function RemoveBeginningAndEndingBlanks(ByVal inputStr As String) As String
            Return Regex.Replace(inputStr, "^\s+$[\r\n]*", "", RegexOptions.Multiline)
        End Function
        Public Function ConvertHTMLToReadableText(ByVal inputHTML As String) As String
            Dim result As String = inputHTML

            'First replace the anchor tags with the just the URL
            Dim doc As New HtmlAgilityPack.HtmlDocument
            doc.LoadHtml(result)
            Dim links = doc.DocumentNode.Descendants("a")
            If links IsNot Nothing AndAlso links.Count > 0 Then

                For linkCtr As Integer = links.Count - 1 To 0 Step -1
                    Dim anchorLink = links(linkCtr)
                    'Dim url As String = String.Empty
                    'If anchorLink.Attributes("href") IsNot Nothing AndAlso anchorLink.Attributes("href").Value IsNot Nothing Then
                    '    url = anchorLink.Attributes("href").Value
                    'End If
                    Dim anchorText As String = String.Empty
                    If anchorLink.InnerHtml IsNot Nothing Then
                        anchorText = anchorLink.InnerHtml
                    End If
                    Dim newNode As HtmlAgilityPack.HtmlNode = doc.CreateElement("span")
                    'newNode.InnerHtml = url
                    newNode.InnerHtml = anchorText

                    anchorLink.ParentNode.InsertBefore(newNode, anchorLink)
                    anchorLink.Remove()
                Next
            End If

            result = doc.DocumentNode.InnerHtml


            ' Remove HTML Development formatting
            result = result.Replace(vbCr, "#NEW LINE#")
            ' Replace line breaks with space because browsers inserts space
            result = result.Replace(vbLf, "#NEW LINE#")
            ' Replace line breaks with space because browsers inserts space
            result = result.Replace(vbTab, String.Empty)
            ' Remove step-formatting
            result = System.Text.RegularExpressions.Regex.Replace(result, "( )+", " ")
            ' Remove repeating speces becuase browsers ignore them
            ' Remove the header (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*head([^>])*>", "<head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*head( )*>)", "</head>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<head>).*(</head>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all scripts (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*script([^>])*>", "<script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*script( )*>)", "</script>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            'result = System.Text.RegularExpressions.Regex.Replace(result, @"(<script>)([^(<script>\.</script>)])*(</script>)",string.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<script>).*(</script>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' remove all styles (prepare first by clearing attributes)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*style([^>])*>", "<style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<( )*(/)( )*style( )*>)", "</style>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(<style>).*(</style>)", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)


            ' insert tabs in spaces of <td> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*td([^>])*>", vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line breaks in places of <BR> and <LI> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*br( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*li( )*>", vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' insert line paragraphs (double line breaks) in place if <P>, <DIV> and <TR> tags
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*div([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*tr([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "<( )*p([^>])*>", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' Remove remaining tags like <a>, links, images, comments etc - anything thats enclosed inside < >
            'result = System.Text.RegularExpressions.Regex.Replace(result, "<[^>]*>", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove except for <a> all other remaining tags like <a>, links, images, comments etc - anything thats enclosed inside < >
            result = System.Text.RegularExpressions.Regex.Replace(result, "<(?!\/?a(?=>|\s.*>))\/?.*?>", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' replace special characters:
            result = System.Text.RegularExpressions.Regex.Replace(result, "&nbsp;", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&quot;", """", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&bull;", " * ", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&lsaquo;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&rsaquo;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&trade;", "(tm)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&frasl;", "/", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&lt;", "<", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&gt;", ">", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&copy;", "(c)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "&reg;", "(r)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove all others. More can be added, see http://hotwired.lycos.com/webmonkey/reference/special_characters/
            result = System.Text.RegularExpressions.Regex.Replace(result, "&(.{2,6});", String.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase)

            ' for testng
            'System.Text.RegularExpressions.Regex.Replace(result, this.txtRegex.Text,string.Empty, System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            ' make line breaking consistent
            result = result.Replace(vbLf, vbCr)

            ' Remove extra line breaks and tabs: replace over 2 breaks with 2 and over 4 tabs with 4. 
            ' Prepare first to remove any whitespaces inbetween the escaped characters and remove redundant tabs inbetween linebreaks
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbCr & ")", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbTab & ")", vbTab & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbTab & ")( )+(" & vbCr & ")", vbTab & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")( )+(" & vbTab & ")", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+(" & vbCr & ")", vbCr & vbCr, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove redundant tabs
            result = System.Text.RegularExpressions.Regex.Replace(result, "(" & vbCr & ")(" & vbTab & ")+", vbCr & vbTab, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            ' Remove multible tabs followind a linebreak with just one tab
            Dim breaks As String = vbCr & vbCr & vbCr
            ' Initial replacement target string for linebreaks
            Dim tabs As String = vbTab & vbTab & vbTab & vbTab & vbTab
            ' Initial replacement target string for tabs
            For index As Integer = 0 To result.Length - 1
                result = result.Replace(breaks, vbCr & vbCr)
                result = result.Replace(tabs, vbTab & vbTab & vbTab & vbTab)
                breaks = breaks + vbCr
                tabs = tabs + vbTab
            Next

            ' Thats it.
            result = Replace(result, "#NEW LINE#", vbCr)
            result = Replace(result, vbCr, vbCrLf)
            Return RemoveBeginningAndEndingBlanks(result)
        End Function
        Public Function EncodeString(ByVal inputString As String)
            Return HtmlEncode(inputString)
        End Function
        Public Function DecodeString(ByVal inputString As String)
            Return HtmlDecode(inputString)
        End Function
        Public Function ContainsText(ByVal inputString As String, ByVal searchString As String, ByVal options As StringMatchoptions) As Boolean
            Dim ret = False
            Select Case options
                Case StringMatchoptions.MatchFullWord
                    If Regex.Match(inputString, String.Format("(?<![-/])\b{0}\b(?!-)", searchString), RegexOptions.IgnoreCase).Success Then
                        ret = True
                    Else
                        ret = False
                    End If
                Case StringMatchoptions.MatchPartialWord
                    If Regex.Match(inputString, searchString, RegexOptions.IgnoreCase).Success Then
                        ret = True
                    Else
                        ret = False
                    End If
            End Select
            Return ret
        End Function
        Public Function GetWordCount(ByVal inputString As String) As Long
            Dim ret As Integer = 0
            If inputString IsNot Nothing AndAlso inputString.Trim.Length > 0 Then
                inputString = inputString.Trim
                ret = Regex.Matches(inputString, "\w+").Count
            End If
            Return ret
        End Function
        Public Function GetWordCount(ByVal inputString As String, ByVal searchString As String, ByVal options As StringMatchoptions) As Long
            Dim ret As Integer = 0
            Select Case options
                Case StringMatchoptions.MatchFullWord
                    ret = Regex.Matches(inputString, String.Format("(?<![-/])\b{0}\b(?!-)", searchString), RegexOptions.IgnoreCase).Count
                Case StringMatchoptions.MatchPartialWord
                    ret = Regex.Matches(inputString, searchString, RegexOptions.IgnoreCase).Count
            End Select
            Return ret
        End Function
        Public Function GetWords(ByVal inputString As String) As String()
            Dim ret() As String = Nothing
            If inputString IsNot Nothing AndAlso inputString.Trim.Length > 0 Then
                inputString = inputString.Trim
                Dim matches As MatchCollection = Regex.Matches(inputString, "\w+")
                ret = matches.Cast(Of Match)().[Select](Function(m) m.Value).ToArray()
            End If
            Return ret
        End Function
        Public Function GetWordByWordNumber(ByVal wordNumber As Integer, ByVal numBerOfWordsToExtract As Integer, ByVal inputString As String) As String
            Dim ret As String = String.Empty
            If inputString IsNot Nothing AndAlso inputString.Trim.Length > 0 Then
                inputString = inputString.Trim
                Dim matches As MatchCollection = Regex.Matches(inputString, "\w+")
                Dim runningWordCounter As Integer = 1
                Dim startIndex As Integer = 0
                For Each matchedWord As Match In matches
                    If runningWordCounter = wordNumber Then
                        'Get the start index of this word
                        Dim indexDetails As Capture = matchedWord.Captures(0)
                        startIndex = indexDetails.Index + 1 'Since comparer return zero index for a string matching in the first word
                        Exit For
                    End If
                    runningWordCounter += 1
                Next
                If startIndex > 0 Then
                    Dim modifiedSourceString As String = Mid(inputString, startIndex).Trim
                    'Now from this string we have to match numBerOfWordsToExtract
                    matches = Nothing
                    matches = Regex.Matches(modifiedSourceString, "\w+")
                    runningWordCounter = 1
                    Dim endIndex As Integer = 0
                    For Each matchedWord As Match In matches
                        If runningWordCounter = numBerOfWordsToExtract Then
                            'Get the start index of this word
                            Dim indexDetails As Capture = matchedWord.Captures(0)
                            If indexDetails IsNot Nothing Then
                                endIndex = indexDetails.Index + indexDetails.Length 'I have not added a 1 like before because here we are talking of endindex where I hadded the length
                                Exit For
                            Else
                                'Now match somehow so better to exit and return null
                                Exit For
                            End If
                        End If
                        runningWordCounter += 1
                    Next
                    If endIndex > 0 Then
                        ret = Left(modifiedSourceString, endIndex)
                    End If
                End If
                matches = Nothing
            End If
            Return ret
        End Function
        Public Function GetCleanedHTML(ByVal inputStr As String) As String
            Dim ret As String = inputStr
            'Remove the blank lines
            ret = Regex.Replace(ret, "^\s*$[\r\n]*", String.Empty, RegexOptions.Multiline)
            'value = Regex.Replace(x, "^\r?\n?$", String.Empty, RegexOptions.Multiline)
            'value = Regex.Replace(x, "^\s*$\n", String.Empty, RegexOptions.Multiline)
            'value = Regex.Replace(x, "^\s*$\r", String.Empty, RegexOptions.Multiline)
            'value = Regex.Replace(x, "(\r)?(^\s*$)+", String.Empty, RegexOptions.Multiline)
            'value = Regex.Replace(x, "(\n)?(^\s*$)+", String.Empty, RegexOptions.Multiline)
            'Remove more than one white spaces
            ret = Regex.Replace(ret, "[ ]{2,}", " ", RegexOptions.None)


            ret = ret.Trim
            'resultString = Regex.Replace(subjectString, @"^\s+$[\r\n]*", "", RegexOptions.Multiline);
            'string fix = Regex.Replace(original, @"^\s*$\n", string.Empty, RegexOptions.Multiline);
            Return ret
        End Function
        Public Function GetStringSimilarityPercentage(ByVal firstString As String, ByVal secondString As String) As Double
            If firstString Is Nothing Then firstString = ""
            If secondString Is Nothing Then secondString = ""
            Dim numMatch As Integer = 0
            Dim numNotMatch As Integer = 0
            Dim numCharLargestString As Integer = 0
            Dim strFirstLength As Integer
            Dim strSecondLength As Integer
            Dim counter As Integer
            Dim percentage As Double
            Dim LoopControl As Integer
            strFirstLength = firstString.Length()
            strSecondLength = secondString.Length()
            If strFirstLength > strSecondLength Then
                LoopControl = strSecondLength - 1
                numCharLargestString = strFirstLength
            Else
                LoopControl = strFirstLength - 1
                numCharLargestString = strSecondLength
            End If
            For counter = 0 To LoopControl
                If firstString(counter).CompareTo(secondString(counter)) = 0 Then
                    numMatch += 1
                Else
                    numNotMatch += 1
                End If
            Next
            percentage = numMatch * 100 / numCharLargestString
            Return percentage
        End Function
        Public Function GetHyperLinks(ByVal inputText As String, ByVal linkType As HyperLinkType, Optional ByVal baseDomain As String = Nothing) As List(Of String)
            If baseDomain Is Nothing Then linkType = HyperLinkType.Both
            Dim ret As New List(Of String)
            'Get the links
            Dim regger As New Regex("((([A-Za-z]{3,9}:(?:\/\/)?)(?:[-;:&=\+\$,\w]+@)?[A-Za-z0-9.-]+|(?:www.|[-;:&=\+\$,\w]+@)[A-Za-z0-9.-]+)((?:\/[\+~%\/.\w-_]*)?\??(?:[-\+=&;%@.\w_]*)#?(?:[\w]*))?)", RegexOptions.Multiline)
            For Each mt As Match In regger.Matches(inputText)
                Select Case linkType
                    Case HyperLinkType.External
                        If Not mt.Value.ToLower.Contains(baseDomain.ToLower) Then
                            ret.Add(mt.Value)
                        End If
                    Case HyperLinkType.Internal
                        If mt.Value.ToLower.Contains(baseDomain.ToLower) Then
                            ret.Add(mt.Value)
                        End If
                    Case HyperLinkType.Both
                        ret.Add(mt.Value)
                End Select
            Next
            Return ret
        End Function
        Public Function GetEnumValue(ByVal enumType As Type, ByVal enumText As String) As Integer
            Return [Enum].Parse(enumType, enumText)
        End Function
        Public Function GetParsedDateValueFromString(ByVal parsedString As String, ByVal format As String) As Date
            Dim ret As Date = Date.MinValue
            If parsedString IsNot Nothing AndAlso IsDate(parsedString) Then
                ret = Date.ParseExact(parsedString, format, CultureInfo.InvariantCulture)
            End If
            Return ret
        End Function
        Public Function GetParsedDoubleValueFromString(ByVal parsedString As String) As Double
            Dim ret As Double = Double.MinValue
            parsedString = Replace(parsedString, ",", "")
            If parsedString IsNot Nothing AndAlso IsNumeric(parsedString) Then
                ret = CDbl(parsedString)
            End If
            Return ret
        End Function
        Public Function GetParsedLongValueFromString(ByVal parsedString As String) As Long
            Dim ret As Long = Long.MinValue
            parsedString = Replace(parsedString, ",", "")
            If parsedString IsNot Nothing AndAlso IsNumeric(parsedString) Then
                ret = CLng(parsedString)
            End If
            Return ret
        End Function
        Public Function StringToDate(ByVal DateString As String) As DateTime?
            Try
                If DateString.Length = 10 Then
                    Return DateTime.ParseExact(DateString, "yyyy-MM-dd", Nothing)
                Else
                    Return DateTime.ParseExact(DateString, "yyyy-MM-dd HH:mm:ss", Nothing)
                End If
            Catch e As Exception
                Return Nothing
            End Try
        End Function
#End Region

        Public Function GetEnumDescription(ByVal EnumConstant As [Enum]) As String
            Dim attr() As DescriptionAttribute = DirectCast(EnumConstant.GetType().GetField(EnumConstant.ToString()).GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
            Return If(attr.Length > 0, attr(0).Description, EnumConstant.ToString)
        End Function

        Public Sub SerializeFromCollection(Of T)(ByVal outputFilePath As String, ByVal collectionToBeSerialized As T)
            'serialize
            Using stream As Stream = File.Open(outputFilePath, FileMode.Create)
                Dim bformatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                bformatter.Serialize(stream, collectionToBeSerialized)
            End Using
        End Sub
        Public Sub SerializeFromCollectionUsingFileStream(Of T)(ByVal outputFilePath As String, ByVal collectionToBeSerialized As T)
            'serialize
            Using stream As New FileStream(outputFilePath, FileMode.Append)
                Dim bformatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                bformatter.Serialize(stream, collectionToBeSerialized)
            End Using
        End Sub
        Public Function DeserializeToCollection(Of T)(ByVal inputFilePath As String) As T
            Using stream As Stream = File.Open(inputFilePath, FileMode.Open)
                Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                Return DirectCast(binaryFormatter.Deserialize(stream), T)
            End Using
        End Function
        Public Function DeserializeToCollectionUsingFileStream(Of T)(ByVal inputFilePath As String) As List(Of T)
            Dim ret As List(Of T) = Nothing
            Using stream As New FileStream(inputFilePath, FileMode.Open)
                Dim binaryFormatter = New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter()
                While stream.Position <> stream.Length
                    If ret Is Nothing Then ret = New List(Of T)
                    ret.Add(DirectCast(binaryFormatter.Deserialize(stream), T))
                End While
                Return ret
            End Using
        End Function
    End Module
End Namespace