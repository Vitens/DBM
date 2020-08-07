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
Imports System.Globalization
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMForecast


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Public PointDriver As DBMPointDriverAbstract


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver

    End Sub


    Public Function GetResults(StartTimestamp As DateTime,
      EndTimestamp As DateTime, NumberOfValues As Integer,
      IsInputDBMPoint As Boolean, HasCorrelationDBMPoint As Boolean,
      SubtractPoint As DBMPoint, Culture As CultureInfo) As List(Of DBMResult)

      ' Retrieves data and calculates forecast and control limits for
      ' this point. Also calculates and stores (historic) forecast errors for
      ' correlation analysis later on. Forecast results are cached and then
      ' reused when possible. This is important because, due to the use of a
      ' moving average, previously calculated results will often need to be
      ' included in later calculations.

      Dim TimeRangeInterval As Double
      Dim SnapshotTimestamp As DateTime
      Dim Result As DBMResult
      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim ForecastTimestamp, HistoryTimestamp, PatternTimestamp As DateTime
      Dim ForecastItems As New Dictionary(Of DateTime, DBMForecastItem)
      Dim Patterns(ComparePatterns), Measurements(EMAPreviousPeriods),
        Forecasts(EMAPreviousPeriods),
        LowerControlLimits(EMAPreviousPeriods),
        UpperControlLimits(EMAPreviousPeriods) As Double

      GetResults = New List(Of DBMResult)

      ' Align timestamps and determine interval.
      StartTimestamp = PreviousInterval(StartTimestamp)
      EndTimestamp = PreviousInterval(EndTimestamp)
      TimeRangeInterval = IntervalSeconds(NumberOfValues,
        (EndTimestamp-StartTimestamp).TotalSeconds)

      ' Get data from data source.
      SnapshotTimestamp = PointDriver.SnapshotTimestamp
      PointDriver.RetrieveData(StartTimestamp, EndTimestamp)
      If SubtractPoint IsNot Nothing Then
        SubtractPoint.PointDriver.RetrieveData(StartTimestamp, EndTimestamp)
      End If

      Do While EndTimestamp > StartTimestamp

        Result = New DBMResult
        Result.Timestamp = PreviousInterval(StartTimestamp)
        Result.IsFutureData = Result.Timestamp > SnapshotTimestamp

        For CorrelationCounter = 0 To CorrelationPreviousPeriods ' Corr. loop.

          ' Retrieve data and calculate forecast. Only do this for the required
          ' timestamp and only process previous timestamps for calculating
          ' correlation results if an event was found.
          If Result.ForecastItem Is Nothing Or (IsInputDBMPoint And
            Abs(Result.Factor) > 0 And HasCorrelationDBMPoint) Or
            Not IsInputDBMPoint Then

            For EMACounter = 0 To EMAPreviousPeriods ' Filter hi freq. variation

              ForecastTimestamp = Result.Timestamp.AddSeconds(
                -(EMAPreviousPeriods-EMACounter+CorrelationCounter)*
                CalculationInterval) ' Timestamp for forecast results

              ' Calculate forecast for this timestamp if there is none yet.
              If Not ForecastItems.ContainsKey(ForecastTimestamp) Then

                HistoryTimestamp = OffsetHoliday(ForecastTimestamp, Culture)

                For PatternCounter = 0 To ComparePatterns ' Data for regression.

                  If PatternCounter = ComparePatterns Then
                    PatternTimestamp = ForecastTimestamp ' Last is measurement.
                  Else
                    PatternTimestamp = HistoryTimestamp.
                      AddDays(-(ComparePatterns-PatternCounter)*7) ' History.
                  End If

                  Patterns(PatternCounter) =
                    PointDriver.DataStore.GetData(PatternTimestamp)
                  If SubtractPoint IsNot Nothing Then ' Subtract input if req'd.
                    Patterns(PatternCounter) -=
                      SubtractPoint.PointDriver.DataStore.
                      GetData(PatternTimestamp)
                  End If

                Next PatternCounter

                ' Calculate and store forecast for this timestamp.
                ForecastItems.Add(ForecastTimestamp, Forecast(Patterns))

              End If

              With ForecastItems.Item(ForecastTimestamp) ' To arrays for EMA.
                Measurements(EMACounter) = .Measurement
                Forecasts(EMACounter) = .Forecast
                LowerControlLimits(EMACounter) = .LowerControlLimit
                UpperControlLimits(EMACounter) = .UpperControlLimit
              End With

            Next EMACounter

            ' Calculate final result using filtered calculation results.
            Result.Calculate(CorrelationPreviousPeriods-CorrelationCounter,
              ExponentialMovingAverage(Measurements),
              ExponentialMovingAverage(Forecasts),
              ExponentialMovingAverage(LowerControlLimits),
              ExponentialMovingAverage(UpperControlLimits))

          End If

        Next CorrelationCounter

        GetResults.Add(Result) ' Add timestamp results.
        StartTimestamp =
          StartTimestamp.AddSeconds(TimeRangeInterval) ' Next interval.

      Loop

      Return GetResults

    End Function


  End Class


End Namespace
