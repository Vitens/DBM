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

Namespace Vitens.DynamicBandwidthMonitor

    Public Class DBMTester

        Private Shared InternationalFormat As Boolean = False

        Private Shared Function FormatDateTime(Timestamp As DateTime) As String
            If InternationalFormat Then
                Return Timestamp.ToUniversalTime.ToString("s") & "Z" ' ISO 8601 UTC
            Else
                Return Timestamp.ToString("s") ' ISO 8601
            End If
        End Function

        Private Shared Function Separator As String
            If InternationalFormat Then
                Return ","
            Else
                Return "	" ' Tab character
            End If
        End Function

        Private Shared Function FormatNumber(Value As Double) As String
            If InternationalFormat Then
                Return Value.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture)
            Else
                Return Value.ToString("0.####")
            End If
        End Function

        Public Shared Sub Main
            Dim Substrings() As String
            Dim InputPointDriver As DBMPointDriver = Nothing
            Dim CorrelationPoints As New Collections.Generic.List(Of DBMCorrelationPoint)
            Dim StartTimestamp, EndTimestamp As DateTime
            Dim Result As DBMResult
            Dim _DBM As New DBM
            For Each CommandLineArg In Environment.GetCommandLineArgs ' Parse command line arguments
                If Text.RegularExpressions.Regex.IsMatch(CommandLineArg, "^[-/](.+)=(.+)$") Then ' Parameter = Value
                    Substrings = CommandLineArg.Split(New Char(){"="c}, 2)
                    Try
                        If Substrings(0).Substring(1).ToLower.Equals("i") Then
                            InputPointDriver = New DBMPointDriver(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("c") Then
                            CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver(Substrings(1)), False))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("cs") Then
                            CorrelationPoints.Add(New DBMCorrelationPoint(New DBMPointDriver(Substrings(1)), True))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("iv") Then
                            DBMParameters.CalculationInterval = Convert.ToInt32(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("p") Then
                            DBMParameters.ComparePatterns = Convert.ToInt32(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("ep") Then
                            DBMParameters.EMAPreviousPeriods = Convert.ToInt32(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("ci") Then
                            DBMParameters.ConfidenceInterval = Convert.ToDouble(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("cp") Then
                            DBMParameters.CorrelationPreviousPeriods = Convert.ToInt32(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("ct") Then
                            DBMParameters.CorrelationThreshold = Convert.ToDouble(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("st") Then
                            StartTimestamp = Convert.ToDateTime(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("et") Then
                            EndTimestamp = Convert.ToDateTime(Substrings(1))
                        ElseIf Substrings(0).Substring(1).ToLower.Equals("f") Then
                            If Substrings(1).ToLower.Equals("local") Then
                                InternationalFormat = False
                            ElseIf Substrings(1).ToLower.Equals("intl") Then
                                InternationalFormat = True
                            End If
                        End If
                    Catch
                    End Try
                End If
            Next
            If InputPointDriver Is Nothing Or StartTimestamp = DateTime.MinValue Then ' Perform unit tests
                Console.Write(DBM.Version)
            Else
                If EndTimestamp = DateTime.MinValue Then
                    EndTimestamp = StartTimestamp ' No end timestamp, set to start timestamp
                Else
                    EndTimestamp = EndTimestamp.AddSeconds(-DBMParameters.CalculationInterval) ' Remove one interval from end timestamp
                End If
                Do While StartTimestamp <= EndTimestamp
                    Console.Write(FormatDateTime(StartTimestamp) & Separator)
                    Result = _DBM.Result(InputPointDriver, CorrelationPoints, StartTimestamp)
                    Console.Write(FormatNumber(Result.Factor) & Separator & FormatNumber(Result.Prediction.MeasuredValue) & Separator & FormatNumber(Result.Prediction.PredictedValue) & Separator & FormatNumber(Result.Prediction.LowerControlLimit) & Separator & FormatNumber(Result.Prediction.UpperControlLimit))
                    If Result.Factor <> 0 And CorrelationPoints.Count > 0 Then ' If an event is found and a correlation point is available
                        Console.Write(Separator & FormatNumber(Result.AbsoluteErrorStats.Count) & Separator & FormatNumber(Result.AbsoluteErrorStats.Slope) & Separator & FormatNumber(Result.AbsoluteErrorStats.Angle) & Separator & FormatNumber(Result.AbsoluteErrorStats.Intercept) & Separator & FormatNumber(Result.AbsoluteErrorStats.StandardError) & Separator & FormatNumber(Result.AbsoluteErrorStats.Correlation) & Separator & FormatNumber(Result.AbsoluteErrorStats.ModifiedCorrelation) & Separator & FormatNumber(Result.AbsoluteErrorStats.Determination))
                        Console.Write(Separator & FormatNumber(Result.RelativeErrorStats.Count) & Separator & FormatNumber(Result.RelativeErrorStats.Slope) & Separator & FormatNumber(Result.RelativeErrorStats.Angle) & Separator & FormatNumber(Result.RelativeErrorStats.Intercept) & Separator & FormatNumber(Result.RelativeErrorStats.StandardError) & Separator & FormatNumber(Result.RelativeErrorStats.Correlation) & Separator & FormatNumber(Result.RelativeErrorStats.ModifiedCorrelation) & Separator & FormatNumber(Result.RelativeErrorStats.Determination))
                        For Each CorrelationPoint In CorrelationPoints
                            Result = _DBM.Result(CorrelationPoint.PointDriver, Nothing, StartTimestamp)
                            Console.Write(Separator & FormatNumber(Result.Prediction.MeasuredValue) & Separator & FormatNumber(Result.Prediction.PredictedValue) & Separator & FormatNumber(Result.Prediction.LowerControlLimit) & Separator & FormatNumber(Result.Prediction.UpperControlLimit))
                        Next
                    End If
                    Console.WriteLine
                    StartTimestamp = StartTimestamp.AddSeconds(DBMParameters.CalculationInterval) ' Next interval
                Loop
            End If
        End Sub

    End Class

End Namespace
