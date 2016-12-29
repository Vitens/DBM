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

    Private Function FormatDateTime(Timestamp As DateTime) As String
        If InternationalFormat Then
            Return Timestamp.ToUniversalTime.ToString("s") & "Z" ' ISO 8601 UTC
        Else
            Return Timestamp.ToString("s") ' ISO 8601
        End If
    End Function

    Private Function Separator As String
        If InternationalFormat Then
            Return ","
        Else
            Return vbTab
        End If
    End Function

    Private Function FormatNumber(Value As Double) As String
        If InternationalFormat Then
            Return Value.ToString("0.####",System.Globalization.CultureInfo.InvariantCulture)
        Else
            Return Value.ToString("0.####")
        End If
    End Function

    Public Sub Main
        Dim Substrings() As String
        Dim InputPointDriver As DBM.DBMPointDriver=Nothing
        Dim CorrelationPoints As New Collections.Generic.List(Of DBM.DBMCorrelationPoint)
        Dim StartTimestamp,EndTimestamp As DateTime
        Dim Output As String
        Dim Result As DBM.DBMResult
        Dim _DBM As New DBM.DBM
        For Each CommandLineArg In Environment.GetCommandLineArgs ' Parse command line arguments
            If Text.RegularExpressions.Regex.IsMatch(CommandLineArg,"^[-/](.+)=(.+)$") Then ' Parameter=Value
                Substrings=CommandLineArg.Split(New Char(){"="c},2)
                Try
                    Select Case Substrings(0).Substring(1).ToLower
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
                EndTimestamp=EndTimestamp.AddSeconds(-DBM.DBMParameters.CalculationInterval) ' Remove one interval from end timestamp
            End If
            Do While StartTimestamp<=EndTimestamp
                Output=FormatDateTime(StartTimestamp) & Separator
                Result=_DBM.Result(InputPointDriver,CorrelationPoints,StartTimestamp)
                Output &= FormatNumber(Result.Factor) & Separator & FormatNumber(Result.Prediction.MeasuredValue) & Separator & FormatNumber(Result.Prediction.PredictedValue) & Separator & FormatNumber(Result.Prediction.LowerControlLimit) & Separator & FormatNumber(Result.Prediction.UpperControlLimit)
                If Result.Factor<>0 And CorrelationPoints.Count>0 Then ' If an event is found and a correlation point is available
                    Output &= Separator & FormatNumber(Result.AbsoluteErrorStats.Count) & Separator & FormatNumber(Result.AbsoluteErrorStats.Slope) & Separator & FormatNumber(Result.AbsoluteErrorStats.Angle) & Separator & FormatNumber(Result.AbsoluteErrorStats.Intercept) & Separator & FormatNumber(Result.AbsoluteErrorStats.StandardError) & Separator & FormatNumber(Result.AbsoluteErrorStats.Correlation) & Separator & FormatNumber(Result.AbsoluteErrorStats.ModifiedCorrelation) & Separator & FormatNumber(Result.AbsoluteErrorStats.Determination)
                    Output &= Separator & FormatNumber(Result.RelativeErrorStats.Count) & Separator & FormatNumber(Result.RelativeErrorStats.Slope) & Separator & FormatNumber(Result.RelativeErrorStats.Angle) & Separator & FormatNumber(Result.RelativeErrorStats.Intercept) & Separator & FormatNumber(Result.RelativeErrorStats.StandardError) & Separator & FormatNumber(Result.RelativeErrorStats.Correlation) & Separator & FormatNumber(Result.RelativeErrorStats.ModifiedCorrelation) & Separator & FormatNumber(Result.RelativeErrorStats.Determination)
                    For Each CorrelationPoint In CorrelationPoints
                        Result=_DBM.Result(CorrelationPoint.PointDriver,Nothing,StartTimestamp)
                        Output &= Separator & FormatNumber(Result.Prediction.MeasuredValue) & Separator & FormatNumber(Result.Prediction.PredictedValue) & Separator & FormatNumber(Result.Prediction.LowerControlLimit) & Separator & FormatNumber(Result.Prediction.UpperControlLimit)
                    Next
                End If
                Console.WriteLine(Output)
                StartTimestamp=StartTimestamp.AddSeconds(DBM.DBMParameters.CalculationInterval) ' Next interval
            Loop
        End If
    End Sub

End Module
