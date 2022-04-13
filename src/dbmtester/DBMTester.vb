' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics
Imports Vitens.DynamicBandwidthMonitor.DBMStrings


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


    Private Shared Function FormatNumber(value As Double) As String

      Return value.ToString(DecimalFormat)

    End Function


    Public Shared Sub Main

      Dim commandLineArg, substrings(), parameter, value As String
      Dim inputPointDriver As DBMPointDriver = Nothing
      Dim correlationPoints As New List(Of DBMCorrelationPoint)
      Dim startTimestamp, endTimestamp As DateTime
      Dim results As List(Of DBMResult)
      Dim dbm As New DBM
      Dim result As DBMResult

      ' Parse command line arguments
      For Each commandLineArg In GetCommandLineArgs
        ' Parameter=Value
        If Regex.IsMatch(commandLineArg, "^[-/].+=.+$") Then
          substrings = commandLineArg.Split(New Char(){"="c}, 2)
          parameter = substrings(0).Substring(1).ToLower
          value = substrings(1)
          Try
            If parameter.Equals("i") Then
              inputPointDriver = New DBMPointDriver(value)
            ElseIf parameter.Equals("c") Then
              correlationPoints.Add(
                New DBMCorrelationPoint(New DBMPointDriver(value), False))
            ElseIf parameter.Equals("cs") Then
              correlationPoints.Add(
                New DBMCorrelationPoint(New DBMPointDriver(value), True))
            ElseIf parameter.Equals("iv") Then
              CalculationInterval = Convert.ToInt32(value)
            ElseIf parameter.Equals("us") Then
              UseSundayForHolidays = Convert.ToBoolean(value)
            ElseIf parameter.Equals("p") Then
              ComparePatterns = Convert.ToInt32(value)
            ElseIf parameter.Equals("ep") Then
              EMAPreviousPeriods = Convert.ToInt32(value)
            ElseIf parameter.Equals("oi") Then
              OutlierCI = Convert.ToDouble(value)
            ElseIf parameter.Equals("bi") Then
              BandwidthCI = Convert.ToDouble(value)
            ElseIf parameter.Equals("cp") Then
              CorrelationPreviousPeriods = Convert.ToInt32(value)
            ElseIf parameter.Equals("ct") Then
              CorrelationThreshold = Convert.ToDouble(value)
            ElseIf parameter.Equals("ra") Then
              RegressionAngleRange = Convert.ToDouble(value)
            ElseIf parameter.Equals("st") Then
              startTimestamp = Convert.ToDateTime(value)
            ElseIf parameter.Equals("et") Then
              endTimestamp = Convert.ToDateTime(value)
            ElseIf parameter.Equals("f") Then
              If value.ToLower.Equals("intl") Then
                CurrentThread.CurrentCulture = InvariantCulture
              End If
            End If
          Catch ex As Exception
            DBM.Logger.LogError(ex.ToString)
            Exit Sub
          End Try
        End If
      Next commandLineArg

      If inputPointDriver IsNot Nothing And
        startTimestamp > DateTime.MinValue Then

        If endTimestamp = DateTime.MinValue Then
          ' No end timestamp, use start timestamp.
          endTimestamp = startTimestamp
        End If

        ' Header
        DBM.Logger.LogInformation(
          CSVComment & Product.Replace(NewLine, NewLine & CSVComment))
        DBM.Logger.LogInformation(
          TimestampText & Separator &
          FactorText & Separator &
          MeasurementText & Separator &
          ForecastText & Separator &
          LowerControlLimitText & Separator &
          UpperControlLimitText,
          inputPointDriver.ToString)

        ' Get results for time range.
        results = dbm.GetResults(inputPointDriver, correlationPoints,
          startTimestamp, endTimestamp)
        For Each result In results

          With result
            DBM.Logger.LogInformation(
              .Timestamp.ToString("s") & Separator &
              FormatNumber(.Factor) & Separator &
              FormatNumber(.ForecastItem.Measurement) & Separator &
              FormatNumber(.ForecastItem.Forecast) & Separator &
              FormatNumber(.ForecastItem.LowerControlLimit) & Separator &
              FormatNumber(.ForecastItem.UpperControlLimit),
              inputPointDriver.ToString)
          End With

        Next result

        DBM.Logger.LogInformation(
          CSVComment & Statistics(results).Brief, inputPointDriver.ToString)

      End If

    End Sub


  End Class


End Namespace
