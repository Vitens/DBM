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


    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean,
      HasCorrelationDBMPoint As Boolean, SubtractPoint As DBMPoint,
      Culture As CultureInfo) As DBMResult

      ' Retrieves data and calculates forecast and control limits for
      ' this point. Also calculates and stores (historic) forecast errors for
      ' correlation analysis later on.

      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim ForecastTimestamp, HistoryTimestamp, PatternTimestamp As DateTime
      Dim Patterns(ComparePatterns), Measurements(EMAPreviousPeriods),
        ForecastValues(EMAPreviousPeriods),
        LowerControlLimits(EMAPreviousPeriods),
        UpperControlLimits(EMAPreviousPeriods) As Double

      Result = New DBMResult
      Result.Timestamp = AlignTimestamp(Timestamp, CalculationInterval)

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

            HistoryTimestamp = OffsetHoliday(ForecastTimestamp, Culture)

            For PatternCounter = 0 To ComparePatterns ' Data for regression.

              If PatternCounter = ComparePatterns Then
                PatternTimestamp = ForecastTimestamp ' Last item: Measurement
              Else
                PatternTimestamp = HistoryTimestamp.
                  AddDays(-(ComparePatterns-PatternCounter)*7) ' History
              End If

              Patterns(PatternCounter) =
                PointDriver.TryGetData(PatternTimestamp) ' Get data
              If SubtractPoint IsNot Nothing Then ' Subtract input if req'd
                Patterns(PatternCounter) -=
                  SubtractPoint.PointDriver.TryGetData(PatternTimestamp)
              End If

            Next PatternCounter

            With Forecast(Patterns) ' Store results in arrays for EMA calculation.
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

    End Function


  End Class


End Namespace
