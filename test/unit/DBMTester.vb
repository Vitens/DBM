Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

    Public Sub Main
        Dim _DBM As New DBM.DBM
        Dim Fields() As String
        Dim InputDBMPointDriver As DBM.DBMPointDriver=Nothing
        Dim DBMCorrelationPoints As New Collections.Generic.List(Of DBM.DBMCorrelationPoint)
        Dim StartTimestamp,EndTimestamp As DateTime
        Dim DBMResult As DBM.DBMResult
        Dim Ticks As Int64
        Console.WriteLine(DBM.DBMFunctions.DBMVersion & vbCrLf)
        For Each CommandLineArg As String In Environment.GetCommandLineArgs
            If Text.RegularExpressions.Regex.IsMatch(CommandLineArg,"^[-/](.+)=(.+)$") Then
                Fields=Split(CommandLineArg,"=",2)
                Try
                    Select Case Mid(Fields(0),2).ToLower
                        Case "i"
                            Console.WriteLine("Input point: " & Fields(1))
                            InputDBMPointDriver=New DBM.DBMPointDriver(Fields(1))
                        Case "c"
                            Console.WriteLine("Correlation point: " & Fields(1))
                            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Fields(1)),False))
                        Case "cs"
                            Console.WriteLine("Correlation point (substract input): " & Fields(1))
                            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Fields(1)),True))
                        Case "iv"
                            Console.Write("Calculation interval: ")
                            DBM.DBMParameters.CalculationInterval=Convert.ToInt32(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.CalculationInterval)
                        Case "p"
                            Console.Write("Compare patterns: ")
                            DBM.DBMParameters.ComparePatterns=Convert.ToInt32(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.ComparePatterns)
                        Case "ep"
                            Console.Write("EMA previous periods: ")
                            DBM.DBMParameters.EMAPreviousPeriods=Convert.ToInt32(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.EMAPreviousPeriods)
                        Case "ci"
                            Console.Write("Confidence interval: ")
                            DBM.DBMParameters.ConfidenceInterval=Convert.ToDouble(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.ConfidenceInterval)
                        Case "cp"
                            Console.Write("Correlation previous periods: ")
                            DBM.DBMParameters.CorrelationPreviousPeriods=Convert.ToInt32(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.CorrelationPreviousPeriods)
                        Case "ct"
                            Console.Write("Correlation threshold: ")
                            DBM.DBMParameters.CorrelationThreshold=Convert.ToDouble(Fields(1))
                            Console.WriteLine(DBM.DBMParameters.CorrelationThreshold)
                        Case "st"
                            Console.Write("Start timestamp: ")
                            StartTimestamp=Convert.ToDateTime(Fields(1))
                            Console.WriteLine(StartTimestamp)
                        Case "et"
                            Console.Write("End timestamp: ")
                            EndTimestamp=DateAdd("s",-DBM.DBMParameters.CalculationInterval,Convert.ToDateTime(Fields(1)))
                            Console.WriteLine(EndTimestamp)
                    End Select
                Catch ex As Exception
                    Console.WriteLine(ex.Message)
                    End
                End Try
            End If
        Next
        If InputDBMPointDriver Is Nothing Or StartTimestamp=DateTime.MinValue Then
            Ticks=DateTime.Now.Ticks
            Console.Write(" * Unit tests --> ")
            If DBM.DBMUnitTests.TestResults Then
                Console.Write("PASSED")
            Else
                Console.Write("FAILED")
            End If
            Console.WriteLine(" (" & Math.Round((DateTime.Now.Ticks-Ticks)/10000) & "ms)")
        Else
            If EndTimestamp=DateTime.MinValue Then EndTimestamp=StartTimestamp
            Do While StartTimestamp<=EndTimestamp
                Console.Write(StartTimestamp & vbTab)
                DBMResult=_DBM.Calculate(InputDBMPointDriver,DBMCorrelationPoints,StartTimestamp)
                Console.Write(DBMResult.Factor & vbTab & DBMResult.CurrValue & vbTab & DBMResult.PredValue & vbTab & DBMResult.LowContrLimit & vbTab & DBMResult.UppContrLimit & vbTab)
                Console.Write(DBMResult.AbsErrorStats.Count & vbTab & DBMResult.AbsErrorStats.Slope & vbTab & DBMResult.AbsErrorStats.Angle & vbTab & DBMResult.AbsErrorStats.Intercept & vbTab & DBMResult.AbsErrorStats.StDevSLinReg & vbTab & DBMResult.AbsErrorStats.Correlation & vbTab & DBMResult.AbsErrorStats.ModifiedCorrelation & vbTab & DBMResult.AbsErrorStats.Determination & vbTab)
                Console.Write(DBMResult.RelErrorStats.Count & vbTab & DBMResult.RelErrorStats.Slope & vbTab & DBMResult.RelErrorStats.Angle & vbTab & DBMResult.RelErrorStats.Intercept & vbTab & DBMResult.RelErrorStats.StDevSLinReg & vbTab & DBMResult.RelErrorStats.Correlation & vbTab & DBMResult.RelErrorStats.ModifiedCorrelation & vbTab & DBMResult.RelErrorStats.Determination)
                For Each thisDBMCorrelationPoint As DBM.DBMCorrelationPoint In DBMCorrelationPoints
                    DBMResult=_DBM.Calculate(thisDBMCorrelationPoint.DBMPointDriver,Nothing,StartTimestamp)
                    Console.Write(vbTab & DBMResult.CurrValue & vbTab & DBMResult.PredValue & vbTab & DBMResult.LowContrLimit & vbTab & DBMResult.UppContrLimit)
                Next
                Console.Write(vbCrLf)
                StartTimestamp=DateAdd("s",DBM.DBMParameters.CalculationInterval,StartTimestamp)
            Loop
        End If
    End Sub

End Module
