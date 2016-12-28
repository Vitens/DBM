Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
'
' This file is part of DBM.
'
' DBM is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' DBM is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with DBM.  If not, see <http://www.gnu.org/licenses/>.

<assembly:System.Reflection.AssemblyTitle("DBMTester")>

Module DBMTester

    Private InternationalFormat As Boolean=False

    Private Function FormatNumber(Value As Double) As String
        If InternationalFormat Then
            Return Value.ToString("0.####",System.Globalization.CultureInfo.InvariantCulture)
        Else
            Return Value.ToString("0.####")
        End If
    End Function

    Private Function FormatDateTime(Timestamp As DateTime) As String
        If InternationalFormat Then
            Return Timestamp.ToUniversalTime.ToString("s") & "Z" ' ISO 8601 UTC
        Else
            Return Timestamp.ToString("s") ' ISO 8601
        End If
    End Function

    Public Sub Calculate(StartTimestamp As DateTime,EndTimestamp As DateTime,InputPointDriver As DBM.DBMPointDriver,CorrelationPoints As Collections.Generic.List(Of DBM.DBMCorrelationPoint))
        Dim Output As String
        Dim Result As DBM.DBMResult
        Dim _DBM As New DBM.DBM
        Do While StartTimestamp<=EndTimestamp
            Output=FormatDateTime(StartTimestamp) & vbTab
            Result=_DBM.Result(InputPointDriver,CorrelationPoints,StartTimestamp)
            Output &= FormatNumber(Result.Factor) & vbTab & FormatNumber(Result.Prediction.MeasuredValue) & vbTab & FormatNumber(Result.Prediction.PredictedValue) & vbTab & FormatNumber(Result.Prediction.LowerControlLimit) & vbTab & FormatNumber(Result.Prediction.UpperControlLimit)
            If Result.Factor<>0 And CorrelationPoints.Count>0 Then ' If an event is found and a correlation point is available
                Output &= vbTab & FormatNumber(Result.AbsoluteErrorStats.Count) & vbTab & FormatNumber(Result.AbsoluteErrorStats.Slope) & vbTab & FormatNumber(Result.AbsoluteErrorStats.Angle) & vbTab & FormatNumber(Result.AbsoluteErrorStats.Intercept) & vbTab & FormatNumber(Result.AbsoluteErrorStats.StandardError) & vbTab & FormatNumber(Result.AbsoluteErrorStats.Correlation) & vbTab & FormatNumber(Result.AbsoluteErrorStats.ModifiedCorrelation) & vbTab & FormatNumber(Result.AbsoluteErrorStats.Determination)
                Output &= vbTab & FormatNumber(Result.RelativeErrorStats.Count) & vbTab & FormatNumber(Result.RelativeErrorStats.Slope) & vbTab & FormatNumber(Result.RelativeErrorStats.Angle) & vbTab & FormatNumber(Result.RelativeErrorStats.Intercept) & vbTab & FormatNumber(Result.RelativeErrorStats.StandardError) & vbTab & FormatNumber(Result.RelativeErrorStats.Correlation) & vbTab & FormatNumber(Result.RelativeErrorStats.ModifiedCorrelation) & vbTab & FormatNumber(Result.RelativeErrorStats.Determination)
                For Each CorrelationPoint In CorrelationPoints
                    Result=_DBM.Result(CorrelationPoint.PointDriver,Nothing,StartTimestamp)
                    Output &= vbTab & FormatNumber(Result.Prediction.MeasuredValue) & vbTab & FormatNumber(Result.Prediction.PredictedValue) & vbTab & FormatNumber(Result.Prediction.LowerControlLimit) & vbTab & FormatNumber(Result.Prediction.UpperControlLimit)
                Next
            End If
            Console.WriteLine(Output)
            StartTimestamp=DateAdd("s",DBM.DBMParameters.CalculationInterval,StartTimestamp) ' Next interval
        Loop
    End Sub

    Public Sub Main
        Dim Substrings() As String
        Dim InputPointDriver As DBM.DBMPointDriver=Nothing
        Dim CorrelationPoints As New Collections.Generic.List(Of DBM.DBMCorrelationPoint)
        Dim StartTimestamp,EndTimestamp As DateTime
        Dim NumThreads,IntervalsPerThread As Integer
        Dim ThreadEndTimestamp As DateTime
        Dim Thread As System.Threading.Thread
        For Each CommandLineArg In Environment.GetCommandLineArgs ' Parse command line arguments
            If Text.RegularExpressions.Regex.IsMatch(CommandLineArg,"^[-/](.+)=(.+)$") Then ' Parameter=Value
                Substrings=CommandLineArg.Split(New Char(){"="c},2)
                Try
                    Select Case Mid(Substrings(0),2).ToLower
                        Case "i"
                            InputPointDriver=New DBM.DBMPointDriver(Substrings(1))
                        Case "c"
                            CorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Substrings(1)),False))
                        Case "cs"
                            CorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Substrings(1)),True))
                        Case "iv"
                            DBM.DBMParameters.CalculationInterval=Convert.ToInt32(Substrings(1))
                        Case "p"
                            DBM.DBMParameters.ComparePatterns=Convert.ToInt32(Substrings(1))
                        Case "ep"
                            DBM.DBMParameters.EMAPreviousPeriods=Convert.ToInt32(Substrings(1))
                        Case "ci"
                            DBM.DBMParameters.ConfidenceInterval=Convert.ToDouble(Substrings(1))
                        Case "cp"
                            DBM.DBMParameters.CorrelationPreviousPeriods=Convert.ToInt32(Substrings(1))
                        Case "ct"
                            DBM.DBMParameters.CorrelationThreshold=Convert.ToDouble(Substrings(1))
                        Case "st"
                            StartTimestamp=Convert.ToDateTime(Substrings(1))
                        Case "et"
                            EndTimestamp=Convert.ToDateTime(Substrings(1))
                        Case "f"
                            If Substrings(1).ToLower="local" Then
                                InternationalFormat=False
                            ElseIf Substrings(1).ToLower="intl" Then
                                InternationalFormat=True
                            End If
                    End Select
                Catch
                End Try
            End If
        Next
        If InputPointDriver Is Nothing Or StartTimestamp=DateTime.MinValue Then ' Perform unit tests
            Console.Write(DBM.DBM.Version)
        Else
            If EndTimestamp=DateTime.MinValue Then
                EndTimestamp=StartTimestamp ' No end timestamp, set to start timestamp
            Else
                EndTimestamp=DateAdd("s",-DBM.DBMParameters.CalculationInterval,EndTimestamp) ' Remove one interval from end timestamp
            End If
            NumThreads=CInt(Math.Min(2^6-1,Math.Max(1,System.Environment.ProcessorCount*2-1))) ' Maximum number of threads (1 to 63) based on number of processors
            IntervalsPerThread=CInt(Math.Ceiling(Math.Max(24*3600/DBM.DBMParameters.CalculationInterval,(DateDiff("s",StartTimestamp,EndTimestamp)/DBM.DBMParameters.CalculationInterval+1)/NumThreads))) ' Minimum of 1 day per thread
            Do While StartTimestamp<=EndTimestamp
                ThreadEndTimestamp=DateAdd("s",DBM.DBMParameters.CalculationInterval*(IntervalsPerThread-1),StartTimestamp) ' Minimum of 1 day per thread
                If ThreadEndTimestamp>EndTimestamp Then
                    ThreadEndTimestamp=EndTimestamp
                End If
                Thread=New System.Threading.Thread(Sub()Calculate(StartTimestamp,ThreadEndTimestamp,InputPointDriver,CorrelationPoints))
                Thread.Start ' Start thread
                System.Threading.Thread.Sleep(100) ' Wait a bit before starting the next thread
                StartTimestamp=DateAdd("s",DBM.DBMParameters.CalculationInterval,ThreadEndTimestamp)
            Loop
        End If
    End Sub

End Module
