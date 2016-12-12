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
        Dim Ticks As Int64
        Dim DBMResult As DBM.DBMResult
        For Each CommandLineArg As String In Environment.GetCommandLineArgs ' Parse command line arguments
            If Text.RegularExpressions.Regex.IsMatch(CommandLineArg,"^[-/](.+)=(.+)$") Then ' Parameter=Value
                Fields=CommandLineArg.Split(New Char(){"="c},2)
                Try
                    Select Case Mid(Fields(0),2).ToLower
                        Case "i"
                            InputDBMPointDriver=New DBM.DBMPointDriver(Fields(1))
                        Case "c"
                            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Fields(1)),False))
                        Case "cs"
                            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(New DBM.DBMPointDriver(Fields(1)),True))
                        Case "iv"
                            DBM.DBMParameters.CalculationInterval=Convert.ToInt32(Fields(1))
                        Case "p"
                            DBM.DBMParameters.ComparePatterns=Convert.ToInt32(Fields(1))
                        Case "ep"
                            DBM.DBMParameters.EMAPreviousPeriods=Convert.ToInt32(Fields(1))
                        Case "ci"
                            DBM.DBMParameters.ConfidenceInterval=Convert.ToDouble(Fields(1))
                        Case "cp"
                            DBM.DBMParameters.CorrelationPreviousPeriods=Convert.ToInt32(Fields(1))
                        Case "ct"
                            DBM.DBMParameters.CorrelationThreshold=Convert.ToDouble(Fields(1))
                        Case "st"
                            StartTimestamp=Convert.ToDateTime(Fields(1))
                        Case "et"
                            EndTimestamp=Convert.ToDateTime(Fields(1))
                    End Select
                Catch
                End Try
            End If
        Next
        If InputDBMPointDriver Is Nothing Or StartTimestamp=DateTime.MinValue Then ' Perform unit tests
            Console.WriteLine(DBM.DBMFunctions.DBMVersion & vbCrLf)
            Ticks=DateTime.Now.Ticks
            Console.Write(" * Unit tests --> ")
            If DBM.DBMUnitTests.TestResults Then
                Console.Write("PASSED")
            Else
                Console.Write("FAILED")
            End If
            Console.WriteLine(" (" & Math.Round((DateTime.Now.Ticks-Ticks)/10000) & "ms)")
        Else
            If EndTimestamp=DateTime.MinValue Then
                EndTimestamp=StartTimestamp ' No end timestamp, set to start timestamp
            Else
                EndTimestamp=DateAdd("s",-DBM.DBMParameters.CalculationInterval,EndTimestamp) ' Remove one interval from end timestamp
            End If
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
                StartTimestamp=DateAdd("s",DBM.DBMParameters.CalculationInterval,StartTimestamp) ' Next interval
            Loop
        End If
    End Sub

End Module
