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


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver
      LastAccessTime = Now

    End Sub


    Public Function IsStale As Boolean

      ' Returns true if this DBMPoint has not been used for at least one
      ' calculation interval. Used by the DBM class to clean up unused
      ' resources.

      Return Now >= AlignTimestamp(LastAccessTime, _
        CalculationInterval).AddSeconds(2*CalculationInterval)

    End Function


    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean, _
      HasCorrelationDBMPoint As Boolean, _
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
      Dim Patterns(ComparePatterns), MeasuredValues(EMAPreviousPeriods), _
        PredictedValues(EMAPreviousPeriods), _
        LowerControlLimits(EMAPreviousPeriods), _
        UpperControlLimits(EMAPreviousPeriods) As Double

      Me.LastAccessTime = Now

      Result = New DBMResult
      Result.Timestamp = AlignTimestamp(Timestamp, CalculationInterval)

      ' Can we reuse stored results?
      If SubtractPoint IsNot PredictionsSubtractPoint Then
        PredictionsSubtractPoint = SubtractPoint
        PredictionsData.Clear ' No, so clear results
        PredictionsQueue.Clear
      End If

      For CorrelationCounter = 0 To CorrelationPreviousPeriods
        If Result.PredictionData Is Nothing Or (IsInputDBMPoint And _
          Result.Factor <> 0 And HasCorrelationDBMPoint) Or _
          Not IsInputDBMPoint Then ' Calculate history for event or correlation.
          For EMACounter = 0 To EMAPreviousPeriods ' For filtering HF variation.
            PredictionTimestamp = Result.Timestamp.AddSeconds _
              (-(EMAPreviousPeriods-EMACounter+CorrelationCounter)* _
              CalculationInterval) ' Timestamp for prediction results
            If Not PredictionsData.TryGetValue(PredictionTimestamp, _
              PredictionData) Then
              ' Calculate prediction data
              For PatternCounter = 0 To ComparePatterns ' Data for regression.
                PatternTimestamp = PredictionTimestamp. _
                  AddDays(-(ComparePatterns-PatternCounter)*7)
                Patterns(PatternCounter) = _
                  PointDriver.TryGetData(PatternTimestamp)
                If SubtractPoint IsNot Nothing Then ' Subtract input if needed.
                  Patterns(PatternCounter) -= _
                    SubtractPoint.PointDriver.TryGetData(PatternTimestamp)
                End If
              Next PatternCounter
              PredictionData = Prediction(Patterns)
              ' Limit number of cached prediction results per point.
              ' Optimized for real-time continuous calculations.
              Do While PredictionsData.Count > _
                EMAPreviousPeriods+2*CorrelationPreviousPeriods
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
          Result.Calculate(CorrelationPreviousPeriods-CorrelationCounter, _
            ExponentialMovingAverage(MeasuredValues), _
            ExponentialMovingAverage(PredictedValues), _
            ExponentialMovingAverage(LowerControlLimits), _
            ExponentialMovingAverage(UpperControlLimits))
        End If
      Next CorrelationCounter

      Return Result

    End Function


  End Class


End Namespace
