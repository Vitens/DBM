Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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


Imports System
Imports System.Collections.Generic
Imports System.Environment
Imports System.Globalization.CultureInfo
Imports System.Text.RegularExpressions
Imports System.Threading.Thread
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMTester")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTester


    ' DBMTester is a command line utility that can be used to quickly
    ' calculate DBM results using the CSV driver.


    Private Shared Function Separator As String

      If CurrentThread.CurrentCulture Is InvariantCulture Then
        Return ","
      Else
        Return "	" ' Tab character
      End If

    End Function


    Private Shared Function FormatNumber(Value As Double) As String

      Return Value.ToString("0.####")

    End Function


    Public Shared Sub Main

      Dim CommandLineArg, Substrings(), Parameter, Value As String
      Dim InputPointDriver As DBMPointDriver = Nothing
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim StartTimestamp, EndTimestamp As DateTime
      Dim Result As DBMResult
      Dim DBM As New DBM

      ' Parse command line arguments
      For Each CommandLineArg In GetCommandLineArgs
        ' Parameter=Value
        If Regex.IsMatch(CommandLineArg, "^[-/].+=.+$") Then
          Substrings = CommandLineArg.Split(New Char(){"="c}, 2)
          Parameter = Substrings(0).Substring(1).ToLower
          Value = Substrings(1)
          Try
            If Parameter.Equals("i") Then
              InputPointDriver = New DBMPointDriver(Value)
            ElseIf Parameter.Equals("c") Then
              CorrelationPoints.Add(
                New DBMCorrelationPoint(New DBMPointDriver(Value), False))
            ElseIf Parameter.Equals("cs") Then
              CorrelationPoints.Add(
                New DBMCorrelationPoint(New DBMPointDriver(Value), True))
            ElseIf Parameter.Equals("iv") Then
              CalculationInterval = Convert.ToInt32(Value)
            ElseIf Parameter.Equals("us") Then
              UseSundayForHolidays = Convert.ToBoolean(Value)
            ElseIf Parameter.Equals("p") Then
              ComparePatterns = Convert.ToInt32(Value)
            ElseIf Parameter.Equals("ep") Then
              EMAPreviousPeriods = Convert.ToInt32(Value)
            ElseIf Parameter.Equals("oi") Then
              OutlierCI = Convert.ToDouble(Value)
            ElseIf Parameter.Equals("bi") Then
              BandwidthCI = Convert.ToDouble(Value)
            ElseIf Parameter.Equals("cp") Then
              CorrelationPreviousPeriods = Convert.ToInt32(Value)
            ElseIf Parameter.Equals("ct") Then
              CorrelationThreshold = Convert.ToDouble(Value)
            ElseIf Parameter.Equals("ra") Then
              RegressionAngleRange = Convert.ToDouble(Value)
            ElseIf Parameter.Equals("st") Then
              StartTimestamp = Convert.ToDateTime(Value)
            ElseIf Parameter.Equals("et") Then
              EndTimestamp = Convert.ToDateTime(Value)
            ElseIf Parameter.Equals("f") Then
              If Value.ToLower.Equals("intl") Then
                CurrentThread.CurrentCulture = InvariantCulture
              End If
            End If
          Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Exit Sub
          End Try
        End If
      Next CommandLineArg

      If InputPointDriver IsNot Nothing And
        StartTimestamp > DateTime.MinValue Then

        If EndTimestamp = DateTime.MinValue Then
          ' No end timestamp, set one interval after start timestamp.
          EndTimestamp = NextInterval(StartTimestamp)
        End If

        For Each Result In DBM.GetResults(InputPointDriver, CorrelationPoints,
          StartTimestamp, EndTimestamp) ' Get results for time range.

          With Result
            Console.WriteLine(.Timestamp.ToString("s") & Separator &
            FormatNumber(.Factor) & Separator &
            FormatNumber(.ForecastItem.Measurement) & Separator &
            FormatNumber(.ForecastItem.Forecast) & Separator &
            FormatNumber(.ForecastItem.LowerControlLimit) & Separator &
            FormatNumber(.ForecastItem.UpperControlLimit))
          End With

        Next Result

      End If

    End Sub


  End Class


End Namespace
