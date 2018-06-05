Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2018  J.H. Fitié, Vitens N.V.
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
Imports System.DateTime
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMForecast


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Public PointDriver As DBMPointDriverAbstract
    Private PointTimeOut As DateTime
    Private Lock As New Object
    Private ForecastsSubtractPoint As DBMPoint
    Private ForecastsData As New Dictionary(Of DateTime, DBMForecastData)
    Public Shared ForecastsCacheSize As Integer =
      2*(EMAPreviousPeriods+2*CorrelationPreviousPeriods+1)


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver
      RefreshTimeOut

    End Sub


    Private Sub RefreshTimeOut

      ' Update timestamp after which point turns stale.

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        PointTimeOut = AlignTimestamp(Now, CalculationInterval).
          AddSeconds(2*CalculationInterval)

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Public Function IsStale As Boolean

      ' Returns true if this DBMPoint has not been used for at least one
      ' calculation interval. Used by the DBM class to clean up unused
      ' resources.

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        Return Now >= PointTimeOut

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Function


    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean,
      HasCorrelationDBMPoint As Boolean,
      Optional SubtractPoint As DBMPoint = Nothing) As DBMResult

      ' Retrieves data and calculates forecast and control limits for
      ' this point. Also calculates and stores (historic) forecast errors for
      ' correlation analysis later on. Forecast results are cached and then
      ' reused when possible. This is important because, due to the use of a
      ' moving average, previously calculated results will often need to be
      ' included in later calculations.

      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim ForecastTimestamp, PatternTimestamp As DateTime
      Dim ForecastData As DBMForecastData = Nothing
      Dim Patterns(ComparePatterns), Measurements(EMAPreviousPeriods),
        ForecastValues(EMAPreviousPeriods),
        LowerControlLimits(EMAPreviousPeriods),
        UpperControlLimits(EMAPreviousPeriods) As Double

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        RefreshTimeOut

        Result = New DBMResult
        Result.Timestamp = AlignTimestamp(Timestamp, CalculationInterval)

        ' Cached results can only be reused if the point that is to be
        ' subtracted is identical to the one used in the cached results.
        If SubtractPoint IsNot ForecastsSubtractPoint Then
          ForecastsSubtractPoint = SubtractPoint
          ForecastsData.Clear ' No, so clear results
        End If

        For CorrelationCounter = 0 To CorrelationPreviousPeriods ' Correl. loop

          ' Retrieve data and calculate forecast. Only do this for the required
          ' timestamp and only process previous timestamps for calculating
          ' correlation results if an event was found.
          If Result.ForecastData Is Nothing Or (IsInputDBMPoint And
            Result.Factor <> 0 And HasCorrelationDBMPoint) Or
            Not IsInputDBMPoint Then

            For EMACounter = 0 To EMAPreviousPeriods ' Filter hi freq. variation

              ForecastTimestamp = Result.Timestamp.AddSeconds(
                -(EMAPreviousPeriods-EMACounter+CorrelationCounter)*
                CalculationInterval) ' Timestamp for forecast results

              If Not ForecastsData.TryGetValue(ForecastTimestamp,
                ForecastData) Then ' Calculate forecast data if not cached

                For PatternCounter = 0 To ComparePatterns ' Data for regression.

                  PatternTimestamp = ForecastTimestamp.
                    AddDays(-(ComparePatterns-PatternCounter)*7) ' Timestamp
                  Patterns(PatternCounter) =
                    PointDriver.TryGetData(PatternTimestamp) ' Get data
                  If SubtractPoint IsNot Nothing Then ' Subtract input if req'd
                    Patterns(PatternCounter) -=
                      SubtractPoint.PointDriver.TryGetData(PatternTimestamp)
                  End If

                Next PatternCounter

                ForecastData = Forecast(Patterns)

                ' Limit number of cached forecast results per point. The size of
                ' the cache is automatically optimized for real-time continuous
                ' calculations. Cache size is limited using random eviction
                ' policy.
                If ForecastsData.Count >= ForecastsCacheSize Then
                  ForecastsData.Remove(ForecastsData.ElementAt(
                    RandomNumber(0, ForecastsData.Count-1)).Key)
                End If

                ' Add calculated forecast to cache.
                ForecastsData.Add(ForecastTimestamp, ForecastData)

              End If

              With ForecastData ' Store results in arrays for EMA calculation.
                Measurements(EMACounter) = .Measurement
                ForecastValues(EMACounter) = .ForecastValue
                LowerControlLimits(EMACounter) = .LowerControlLimit
                UpperControlLimits(EMACounter) = .UpperControlLimit
              End With

            Next EMACounter

            ' Calculate final result using filtered calculation results.
            Result.Calculate(CorrelationPreviousPeriods-CorrelationCounter,
              ExponentialMovingAverage(Measurements),
              ExponentialMovingAverage(ForecastValues),
              ExponentialMovingAverage(LowerControlLimits),
              ExponentialMovingAverage(UpperControlLimits))

          End If

        Next CorrelationCounter

        Return Result

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Function


  End Class


End Namespace
