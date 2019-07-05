Imports Utilities.DAL
Imports System.Threading
Module ExcelModifier

    Private _cts As CancellationTokenSource
    Sub Main()
        _cts = New CancellationTokenSource
        'Dim command As String = Environment.CommandLine
        'If command IsNot Nothing Then
        'Dim args As String() = Command.Trim.Split(" ")
        Dim args As String() = Environment.GetCommandLineArgs
        If args.Count > 0 Then
            For i = 1 To args.Count - 1
                Console.WriteLine(args(i))
            Next
        End If
        ModifyExcel(args(1), args(2), args(3), args(4), args(5))
        'End If
    End Sub

    Public Sub ModifyExcel(ByVal filePath As String, ByVal initialCapital As Decimal, ByVal capitalForPumpIn As Decimal, ByVal minimumEarnedCapitalToWithdraw As Decimal, ByVal amountToBeWithdrawn As Decimal)
        Try
            Console.WriteLine("Opening Excel")
            Using excelWriter As New ExcelHelper(filePath, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, New CancellationTokenSource)
                excelWriter.SetActiveSheet("Data")
                Dim rowCout As Long = excelWriter.GetLastRow
                Dim columnCout As Long = excelWriter.GetLastCol
                Dim range As String = excelWriter.GetNamedRange(1, rowCout - 1, 1, columnCout - 1)
                Console.WriteLine("Creating Pivot Table")
                excelWriter.CreatPivotTable("Data", range, "Summary", "Month", "PL After Brokerage")
                excelWriter.CreateNewSheet("Day Pivot")
                excelWriter.CreatPivotTable("Data", range, "Day Pivot", "Trading Date", "PL After Brokerage")

                excelWriter.SetActiveSheet("Day Pivot")
                Dim dayPivotData As Object(,) = excelWriter.GetExcelInMemory()

                excelWriter.SetActiveSheet("Data")
                excelWriter.SetData(1, columnCout + 1, "Available Capital")
                excelWriter.SetData(1, columnCout + 2, "Withdrawn Capital")
                excelWriter.SetData(1, columnCout + 3, "Pump In Capital")
                excelWriter.SetData(1, columnCout + 4, "Effective Capital For Next Day")
                excelWriter.SetData(1, columnCout + 5, "Equity Curve DrawUp")
                excelWriter.SetData(1, columnCout + 6, "Equity Curve DrawDown")

                Dim availableCapital As Decimal = initialCapital
                Dim startingRow As Integer = 2
                Dim totalWithdraw As Decimal = 0
                Dim totalPump As Decimal = 0
                Dim drawUp As Decimal = 0
                Dim drawDown As Decimal = 0
                Dim maxDrawUp As Decimal = Decimal.MinValue
                Dim minDrawDown As Decimal = Decimal.MaxValue
                Dim totalDays As Integer = 0
                Dim totalWinningDays As Integer = 0
                Dim totalLossDays As Integer = 0
                For i As Integer = 2 To dayPivotData.GetLength(0) - 2 Step 1
                    Console.WriteLine(String.Format("Writing for day {0} of {1}", i - 1, dayPivotData.GetLength(0) - 3))
                    Dim runningDate As Date = Date.FromOADate(dayPivotData(i, 1))
                    Dim totalPLOfTheDay As Decimal = dayPivotData(i, 2)
                    totalDays += 1
                    If totalPLOfTheDay > 0 Then
                        totalWinningDays += 1
                    Else
                        totalLossDays += 1
                    End If
                    If totalPLOfTheDay > 0 Then
                        drawUp += totalPLOfTheDay
                        If drawDown < 0 Then
                            drawDown += totalPLOfTheDay
                            If drawDown > 0 Then drawDown = 0
                        End If
                    ElseIf totalPLOfTheDay <= 0 Then
                        drawDown += totalPLOfTheDay
                        If drawUp > 0 Then
                            drawUp += totalPLOfTheDay
                            If drawUp < 0 Then drawUp = 0
                        End If
                    Else
                        Throw New ApplicationException("Check pl data. Excel Modifier")
                    End If


                    availableCapital = Math.Round(availableCapital + totalPLOfTheDay, 2)
                    Dim withdraw As Decimal = 0
                    If availableCapital >= minimumEarnedCapitalToWithdraw Then
                        withdraw = amountToBeWithdrawn
                        totalWithdraw += withdraw
                    End If
                    Dim pumpInCapital As Decimal = 0
                    If availableCapital < capitalForPumpIn Then
                        pumpInCapital = initialCapital - availableCapital
                        totalPump += pumpInCapital
                    End If
                    Dim effectiveCapital As Decimal = Math.Round(availableCapital + pumpInCapital - withdraw, 2)
                    For row As Integer = startingRow To rowCout Step 1
                        'Console.WriteLine(String.Format("Checking row {0} of {1}", row, rowCout))
                        Dim currentDate As Date = excelWriter.GetData(row, 1)
                        If currentDate.Date = runningDate.Date Then
                            'Console.WriteLine(String.Format("Writting row {0} of {1}", row, rowCout))
                            'excelWriter.SetData(row, columnCout + 1, availableCapital, "##,##,###.00", ExcelHelper.XLAlign.Right)
                            'excelWriter.SetData(row, columnCout + 2, withdraw, "##,##,###.00", ExcelHelper.XLAlign.Right)
                            'excelWriter.SetData(row, columnCout + 3, pumpInCapital, "##,##,###.00", ExcelHelper.XLAlign.Right)
                            'excelWriter.SetData(row, columnCout + 4, effectiveCapital, "##,##,###.00", ExcelHelper.XLAlign.Right)
                            If row = rowCout Then
                                excelWriter.SetData(row, columnCout + 1, availableCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                excelWriter.SetData(row, columnCout + 2, withdraw, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                excelWriter.SetData(row, columnCout + 3, pumpInCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                excelWriter.SetData(row, columnCout + 4, effectiveCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                excelWriter.SetData(row, columnCout + 5, drawUp, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                excelWriter.SetData(row, columnCout + 6, drawDown, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                startingRow = row
                                Exit For
                            End If
                        Else
                            excelWriter.SetData(row - 1, columnCout + 1, availableCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            excelWriter.SetData(row - 1, columnCout + 2, withdraw, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            excelWriter.SetData(row - 1, columnCout + 3, pumpInCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            excelWriter.SetData(row - 1, columnCout + 4, effectiveCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            excelWriter.SetData(row - 1, columnCout + 5, drawUp, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            excelWriter.SetData(row - 1, columnCout + 6, drawDown, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                            startingRow = row
                            Exit For
                        End If
                    Next
                    availableCapital = effectiveCapital
                    maxDrawUp = Math.Max(maxDrawUp, drawUp)
                    minDrawDown = Math.Min(minDrawDown, drawDown)
                Next

                excelWriter.SetActiveSheet("Summary")
                Dim monthPivotData As Object(,) = excelWriter.GetExcelInMemory()
                Dim totalMonth As Integer = monthPivotData.GetLength(0) - 3

                excelWriter.SetData(1, 6, "Initial Capital")
                excelWriter.SetData(1, 7, initialCapital, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(2, 6, "Minimum Earned Capital To Withdraw")
                excelWriter.SetData(2, 7, minimumEarnedCapitalToWithdraw, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(3, 6, "Amount To Be Withdrawn")
                excelWriter.SetData(3, 7, amountToBeWithdrawn, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(4, 6, "Capital For Pump In")
                excelWriter.SetData(4, 7, capitalForPumpIn, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(7, 6, "Total Withdrawn")
                excelWriter.SetData(7, 7, totalWithdraw, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(8, 6, "Total Pump In")
                excelWriter.SetData(8, 7, totalPump, "##,##,##0.00", ExcelHelper.XLAlign.Right)

                Dim totalRI As Decimal = Math.Round(((totalWithdraw / (totalPump + initialCapital)) / totalMonth) * 100, 2)
                excelWriter.SetData(9, 6, "Return Of Investment(Per Month)")
                excelWriter.SetData(9, 7, String.Format("{0}%", totalRI), ExcelHelper.XLAlign.Right)

                Dim n As Integer = 15
                excelWriter.SetData(n + 1, 6, "Max DrawDown")
                excelWriter.SetData(n + 1, 7, minDrawDown, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(n + 2, 6, "Max DrawUp")
                excelWriter.SetData(n + 2, 7, maxDrawUp, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(n + 3, 6, "Total Days")
                excelWriter.SetData(n + 3, 7, totalDays, "##,##,##0", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(n + 4, 6, "Total Winning Days")
                excelWriter.SetData(n + 4, 7, totalWinningDays, "##,##,##0", ExcelHelper.XLAlign.Right)
                excelWriter.SetData(n + 5, 6, "Total Losing Days")
                excelWriter.SetData(n + 5, 7, totalLossDays, "##,##,##0", ExcelHelper.XLAlign.Right)
                n = 6
                excelWriter.SetData(n + 6, 6, "Day Win Ratio")
                excelWriter.SetData(n + 6, 7, Math.Round((totalWinningDays / totalDays) * 100, 2), "##,##,##0.00", ExcelHelper.XLAlign.Right)

                excelWriter.SetActiveSheet("Summary")
                Console.WriteLine("Saving excel...")
                excelWriter.SaveExcel()
            End Using
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        Finally
            Console.WriteLine("Process complete")
        End Try
    End Sub
End Module
