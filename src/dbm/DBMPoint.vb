Option Explicit
Option Strict


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
Imports System.Globalization
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMForecast


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Private _pointDriver As DBMPointDriverAbstract


    Public Property PointDriver As DBMPointDriverAbstract
      Get
        Return _pointDriver
      End Get
      Set(value As DBMPointDriverAbstract)
        _pointDriver = value
      End Set
    End Property


    Public Sub New(pointDriver As DBMPointDriverAbstract)

      DBM.Logger.LogDebug(pointDriver.ToString, pointDriver.ToString)
      Me.PointDriver = pointDriver

    End Sub


    Public Function GetResults(startTimestamp As DateTime,
      endTimestamp As DateTime, numberOfValues As Integer,
      isInputDBMPoint As Boolean, hasCorrelationDBMPoint As Boolean,
      subtractPoint As DBMPoint, culture As CultureInfo) As List(Of DBMResult)

      ' Retrieves data and calculates forecast and control limits for
      ' this point. Also calculates and stores (historic) forecast errors for
      ' correlation analysis later on. Forecast results are cached and then
      ' reused when possible. This is important because, due to the use of a
      ' moving average, previously calculated results will often need to be
      ' included in later calculations.

      Dim timeRangeInterval As Double
      Dim snapshotTimestamp As DateTime
      Dim result As DBMResult
      Dim correlationCounter, emaCounter, patternCounter As Integer
      Dim forecastTimestamp, historyTimestamp, patternTimestamp As DateTime
      Dim forecastItems As New Dictionary(Of DateTime, DBMForecastItem)
      Dim patterns(ComparePatterns), measurements(EMAPreviousPeriods),
        forecasts(EMAPreviousPeriods),
        lowerControlLimits(EMAPreviousPeriods),
        upperControlLimits(EMAPreviousPeriods) As Double

      GetResults = New List(Of DBMResult)

      ' Align timestamps and determine interval.
      startTimestamp = PreviousInterval(startTimestamp)
      If Not IsOnInterval(endTimestamp) Then
        ' Add one extra interval if the end timestamp is not exactly on an
        ' interval boundary. This is required for linear interpolation of
        ' non-stepped data to the end of the interval.
        endTimestamp = NextInterval(endTimestamp)
      End If
      endTimestamp = NextInterval(endTimestamp)
      timeRangeInterval = IntervalSeconds(numberOfValues,
        (endTimestamp-startTimestamp).TotalSeconds)

      ' Get data from data source.
      snapshotTimestamp = Me.PointDriver.snapshotTimestamp
      Me.PointDriver.RetrieveData(startTimestamp, endTimestamp)
      If subtractPoint IsNot Nothing Then
        subtractPoint.PointDriver.RetrieveData(startTimestamp, endTimestamp)
      End If

      Do While endTimestamp > startTimestamp

        result = New DBMResult
        result.Timestamp = PreviousInterval(startTimestamp)
        result.IsFutureData = result.Timestamp > snapshotTimestamp

        For correlationCounter = 0 To CorrelationPreviousPeriods ' Corr. loop.

          ' Retrieve data and calculate forecast. Only do this for the required
          ' timestamp and only process previous timestamps for calculating
          ' correlation results if an event was found.
          If result.ForecastItem Is Nothing Or (isInputDBMPoint And
            Abs(result.Factor) > 0 And hasCorrelationDBMPoint) Or
            Not isInputDBMPoint Then

            For emaCounter = 0 To EMAPreviousPeriods ' Filter hi freq. variation

              forecastTimestamp = result.Timestamp.AddSeconds(
                -(EMAPreviousPeriods-emaCounter+correlationCounter)*
                CalculationInterval) ' Timestamp for forecast results

              ' Calculate forecast for this timestamp if there is none yet.
              If Not forecastItems.ContainsKey(forecastTimestamp) Then

                historyTimestamp = OffsetHoliday(forecastTimestamp, culture)

                For patternCounter = 0 To ComparePatterns ' Data for regression.

                  If patternCounter = ComparePatterns Then
                    patternTimestamp = forecastTimestamp ' Last is measurement.
                  Else
                    patternTimestamp = historyTimestamp.
                      AddDays(-(ComparePatterns-patternCounter)*7) ' History.
                  End If

                  patterns(patternCounter) =
                    Me.PointDriver.DataStore.GetData(patternTimestamp)
                  If subtractPoint IsNot Nothing Then ' Subtract input if req'd.
                    patterns(patternCounter) -=
                      subtractPoint.PointDriver.DataStore.
                      GetData(patternTimestamp)
                  End If

                Next patternCounter

                ' Calculate and store forecast for this timestamp.
                forecastItems.Add(forecastTimestamp, Forecast(patterns))

              End If

              With forecastItems.Item(forecastTimestamp) ' To arrays for EMA.
                measurements(emaCounter) = .Measurement
                forecasts(emaCounter) = .Forecast
                lowerControlLimits(emaCounter) = .LowerControlLimit
                upperControlLimits(emaCounter) = .UpperControlLimit
              End With

            Next emaCounter

            ' Calculate final result using filtered calculation results.
            result.Calculate(CorrelationPreviousPeriods-correlationCounter,
              ExponentialMovingAverage(measurements),
              ExponentialMovingAverage(forecasts),
              ExponentialMovingAverage(lowerControlLimits),
              ExponentialMovingAverage(upperControlLimits))

          End If

        Next correlationCounter

        GetResults.Add(result) ' Add timestamp results.
        startTimestamp =
          startTimestamp.AddSeconds(timeRangeInterval) ' Next interval.

      Loop

      Return GetResults

    End Function


  End Class


End Namespace
