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
Imports System.Math
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMForecast


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Public PointDriver As DBMPointDriverAbstract
    Private Lock As New Object
    Private CacheStale As New DBMStale
    Private SubtractPointsCache As New DBMCache(
      CInt(Sqrt(4^CacheSizeFactor)/2)) ' Cache of forecast results; 8 items


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver

    End Sub


    Public Sub PrepareData(StartTimestamp As DateTime, EndTimestamp As DateTime)

      ' Retrieve and store values in bulk for the passed time range from a
      ' source of data, to be used in the PointDriver.GetData method. Called
      ' from the DBM.PrepareData sub and passed on to the PointDriver.

      ' SyncLock: Access to this method has to be synchronized because the
      '           (expensive) PrepareData method is called for the PointDriver
      '           instance. Since multiple threads might simultaneously call the
      '           DBM.PrepareData method, where this method is called from for
      '           each required PointDriver, access has to be synchronized so
      '           that only the first call will actually prepare the data.
      '           Subsequent calls will then use this cached data. The same lock
      '           is shared between PrepareData and Result because these methods
      '           should not be executed simultaneously.

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        PointDriver.TryPrepareData(StartTimestamp, EndTimestamp)

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean,
      HasCorrelationDBMPoint As Boolean,
      Optional SubtractPoint As DBMPoint = Nothing) As DBMResult

      ' Retrieves data and calculates forecast and control limits for
      ' this point. Also calculates and stores (historic) forecast errors for
      ' correlation analysis later on. Forecast results are cached and then
      ' reused when possible. This is important because, due to the use of a
      ' moving average, previously calculated results will often need to be
      ' included in later calculations.

      Dim ForecastItemsCache As DBMCache ' Cached forecast results for subtr.pt.
      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim ForecastTimestamp, PatternTimestamp As DateTime
      Dim ForecastItem As DBMForecastItem = Nothing
      Dim Patterns(ComparePatterns), Measurements(EMAPreviousPeriods),
        ForecastValues(EMAPreviousPeriods),
        LowerControlLimits(EMAPreviousPeriods),
        UpperControlLimits(EMAPreviousPeriods) As Double

      ' SyncLock: Access to this method has to be synchronized. The same lock is
      '           shared between PrepareData and Result because these methods
      '           should not be executed simultaneously.

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        Result = New DBMResult
        Result.Timestamp = AlignTimestamp(Timestamp, CalculationInterval)

        ' If required, create new cache for this subtract point. The size of the
        ' cache is automatically optimized for real-time continuous
        ' calculations.
        If CacheStale.IsStale Then SubtractPointsCache.Clear ' Clear stale cache
        If Not SubtractPointsCache.HasItem(SubtractPoint) Then
          SubtractPointsCache.AddItem(SubtractPoint, New DBMCache(
            CacheSizeFactor*(EMAPreviousPeriods+
            2*CorrelationPreviousPeriods+1))) ' 208 items
        End If
        ForecastItemsCache = DirectCast(SubtractPointsCache.
          GetItem(SubtractPoint), DBMCache)

        For CorrelationCounter = 0 To CorrelationPreviousPeriods ' Correl. loop

          ' Retrieve data and calculate forecast. Only do this for the required
          ' timestamp and only process previous timestamps for calculating
          ' correlation results if an event was found.
          If Result.ForecastItem Is Nothing Or (IsInputDBMPoint And
            Result.Factor <> 0 And HasCorrelationDBMPoint) Or
            Not IsInputDBMPoint Then

            For EMACounter = 0 To EMAPreviousPeriods ' Filter hi freq. variation

              ForecastTimestamp = Result.Timestamp.AddSeconds(
                -(EMAPreviousPeriods-EMACounter+CorrelationCounter)*
                CalculationInterval) ' Timestamp for forecast results
              ForecastItem = DirectCast(ForecastItemsCache.
                GetItem(ForecastTimestamp), DBMForecastItem)

              If ForecastItem Is Nothing Then ' Not cached

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

                ForecastItem = Forecast(Patterns)
                ForecastItemsCache.AddItem(ForecastTimestamp, ForecastItem)

              End If

              With ForecastItem ' Store results in arrays for EMA calculation.
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
