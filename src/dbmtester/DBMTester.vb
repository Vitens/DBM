Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014, 2015, 2016, 2017, 2018  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


' Assembly title
<assembly:System.Reflection.AssemblyTitle("DBMTester")>


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTester


    ' DBMTester is a command line utility that can be used to quickly
    ' calculate DBM results using the CSV driver.


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
        Return Value.ToString("0.####", InvariantCulture)
      Else
        Return Value.ToString("0.####")
      End If

    End Function


    Public Shared Sub Main

      Dim CommandLineArg, Substrings(), Parameter, Value As String
      Dim InputPointDriver As DBMPointDriver = Nothing
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim StartTimestamp, EndTimestamp As DateTime
      Dim Result As DBMResult
      Dim _DBM As New DBM

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
              CorrelationPoints.Add _
                (New DBMCorrelationPoint(New DBMPointDriver(Value), False))
            ElseIf Parameter.Equals("cs") Then
              CorrelationPoints.Add _
                (New DBMCorrelationPoint(New DBMPointDriver(Value), True))
            ElseIf Parameter.Equals("iv") Then
              CalculationInterval = Convert.ToInt32(Value)
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
              If Value.ToLower.Equals("local") Then
                InternationalFormat = False
              ElseIf Value.ToLower.Equals("intl") Then
                InternationalFormat = True
              End If
            End If
          Catch ex As Exception
            Console.WriteLine(ex.ToString)
            Exit Sub
          End Try
        End If
      Next

      If InputPointDriver IsNot Nothing And _
        StartTimestamp > DateTime.MinValue Then
        If EndTimestamp = DateTime.MinValue Then
          ' No end timestamp, set to start timestamp
          EndTimestamp = StartTimestamp
        Else
          ' Remove one interval from end timestamp
          EndTimestamp = EndTimestamp.AddSeconds(-CalculationInterval)
        End If
        _DBM.PrepareData(InputPointDriver, CorrelationPoints, _
          StartTimestamp, EndTimestamp.AddSeconds(CalculationInterval))
        Do While StartTimestamp <= EndTimestamp
          Result = _DBM.Result _
            (InputPointDriver, CorrelationPoints, StartTimestamp)
          With Result
            Console.WriteLine(FormatDateTime(.Timestamp) & Separator & _
            FormatNumber(.Factor) & Separator & _
            FormatNumber(.PredictionData.MeasuredValue) & Separator & _
            FormatNumber(.PredictionData.PredictedValue) & Separator & _
            FormatNumber(.PredictionData.LowerControlLimit) & Separator & _
            FormatNumber(.PredictionData.UpperControlLimit))
          End With
          ' Next interval
          StartTimestamp = StartTimestamp.AddSeconds(CalculationInterval)
        Loop
      End If

    End Sub


  End Class


End Namespace
