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
Imports System.DateTime
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMPrediction


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Public PointDriver As DBMPointDriverAbstract
    Private LastAccessTime As DateTime
    Private PredictionsSubtractPoint As DBMPoint
    Private PredictionsData As New Dictionary(Of DateTime, DBMPredictionData)
    Private PredictionsQueue As New Queue(Of DateTime) ' Insertion order queue
    Public Shared PredictionsCacheSize As Integer =
      EMAPreviousPeriods+2*CorrelationPreviousPeriods+1


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver
      LastAccessTime = Now

    End Sub


    Public Function IsStale As Boolean

      ' Returns true if this DBMPoint has not been used for at least one
      ' calculation interval. Used by the DBM class to clean up unused
      ' resources.

      Return Now >= AlignTimestamp(LastAccessTime,
        CalculationInterval).AddSeconds(2*CalculationInterval)

    End Function


    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean,
      HasCorrelationDBMPoint As Boolean,
      Optional SubtractPoint As DBMPoint = Nothing) As DBMResult

      ' Retrieves data and calculates prediction and control limits for
      ' this point. Also calculates and stores (historic) prediction errors for
      ' correlation analysis later on. Prediction results are cached and then
      ' reused when possible. This is important because, due to the use of a
      ' moving average, previously calculated results will often need to be
      ' included in later calculations.

      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim PredictionTimestamp, PatternTimestamp As DateTime
      Dim PredictionData As DBMPredictionData = Nothing
      Dim Patterns(ComparePatterns), MeasuredValues(EMAPreviousPeriods),
        PredictedValues(EMAPreviousPeriods),
        LowerControlLimits(EMAPreviousPeriods),
        UpperControlLimits(EMAPreviousPeriods) As Double

      LastAccessTime = Now

      Result = New DBMResult
      Result.Timestamp = AlignTimestamp(Timestamp, CalculationInterval)

      ' Cached results can only be reused if the point that is to be subtracted
      ' is identical to the one used in the cached results.
      If SubtractPoint IsNot PredictionsSubtractPoint Then
        PredictionsSubtractPoint = SubtractPoint
        PredictionsData.Clear ' No, so clear results
        PredictionsQueue.Clear
      End If

      For CorrelationCounter = 0 To CorrelationPreviousPeriods ' Correl. loop

        ' Retrieve data and calculate prediction. Only do this for the required
        ' timestamp and only process previous timestamps for calculating
        ' correlation results if an exception was found.
        If Result.PredictionData Is Nothing Or (IsInputDBMPoint And
          Result.Factor <> 0 And HasCorrelationDBMPoint) Or
          Not IsInputDBMPoint Then

          For EMACounter = 0 To EMAPreviousPeriods ' Filter high freq. variation

            PredictionTimestamp = Result.Timestamp.AddSeconds(
              -(EMAPreviousPeriods-EMACounter+CorrelationCounter)*
              CalculationInterval) ' Timestamp for prediction results

            If Not PredictionsData.TryGetValue(PredictionTimestamp,
              PredictionData) Then ' Calculate prediction data if not cached

              For PatternCounter = 0 To ComparePatterns ' Data for regression.

                PatternTimestamp = PredictionTimestamp.
                  AddDays(-(ComparePatterns-PatternCounter)*7) ' Timestamp
                Patterns(PatternCounter) =
                  PointDriver.TryGetData(PatternTimestamp) ' Get data
                If SubtractPoint IsNot Nothing Then ' Subtract input if needed.
                  Patterns(PatternCounter) -=
                    SubtractPoint.PointDriver.TryGetData(PatternTimestamp)
                End If

              Next PatternCounter

              PredictionData = Prediction(Patterns)

              ' Limit number of cached prediction results per point. The size of
              ' the cache is automatically optimized for real-time continuous
              ' calculations.
              Do While PredictionsData.Count >= PredictionsCacheSize
                ' Use the queue to remove the least recently inserted timestamp.
                PredictionsData.Remove(PredictionsQueue.Dequeue)
              Loop

              ' Add calculated prediction to cache and queue.
              PredictionsData.Add(PredictionTimestamp, PredictionData)
              PredictionsQueue.Enqueue(PredictionTimestamp)

            End If

            With PredictionData ' Store results in arrays for EMA calculation.
              MeasuredValues(EMACounter) = .MeasuredValue
              PredictedValues(EMACounter) = .PredictedValue
              LowerControlLimits(EMACounter) = .LowerControlLimit
              UpperControlLimits(EMACounter) = .UpperControlLimit
            End With

          Next EMACounter

          ' Calculate final result using filtered calculation results.
          Result.Calculate(CorrelationPreviousPeriods-CorrelationCounter,
            ExponentialMovingAverage(MeasuredValues),
            ExponentialMovingAverage(PredictedValues),
            ExponentialMovingAverage(LowerControlLimits),
            ExponentialMovingAverage(UpperControlLimits))

        End If

      Next CorrelationCounter

      Return Result

    End Function


  End Class


End Namespace
