Option Explicit
Option Strict


' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPoint


    Public DataManager As DBMDataManager
    Private Predictions As New Dictionary(Of DateTime, DBMPrediction)
    Private PredictionsSubtractPoint As DBMPoint


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      ' Each DBMPoint has a DBMDataManager which is responsible for retrieving
      ' and caching input data. The Data Manager stores and uses a
      ' DBMPointDriverAbstract object, which has a GetData method used for
      ' retrieving data.

      DataManager = New DBMDataManager(PointDriver)

    End Sub


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
      Dim Prediction As New DBMPrediction
      Dim Patterns(ComparePatterns), MeasuredValues(EMAPreviousPeriods), _
        PredictedValues(EMAPreviousPeriods), _
        LowerControlLimits(EMAPreviousPeriods), _
        UpperControlLimits(EMAPreviousPeriods) As Double

      Result = New DBMResult

      ' Can we reuse stored results?
      If SubtractPoint IsNot PredictionsSubtractPoint Then
        Predictions.Clear ' No, so clear results
        PredictionsSubtractPoint = SubtractPoint
      End If

      For CorrelationCounter = 0 To CorrelationPreviousPeriods
        If Result.Prediction Is Nothing Or (IsInputDBMPoint And _
          Result.Factor <> 0 And HasCorrelationDBMPoint) Or _
          Not IsInputDBMPoint Then ' Calculate history for event or correlation.
          For EMACounter = 0 To EMAPreviousPeriods ' For filtering HF variation.
            PredictionTimestamp = Timestamp.AddSeconds _
              (-(EMAPreviousPeriods-EMACounter+CorrelationCounter)* _
              CalculationInterval) ' Timestamp for prediction results
            If Predictions.ContainsKey(PredictionTimestamp) Then ' From cache
              Prediction = Predictions.Item(PredictionTimestamp).ShallowCopy
            Else ' Calculate prediction data
              For PatternCounter = 0 To ComparePatterns ' Data for regression.
                PatternTimestamp = PredictionTimestamp. _
                  AddDays(-(ComparePatterns-PatternCounter)*7)
                Patterns(PatternCounter) = DataManager.Value(PatternTimestamp)
                If SubtractPoint IsNot Nothing Then ' Subtract input if needed.
                  Patterns(PatternCounter) -= _
                    SubtractPoint.DataManager.Value(PatternTimestamp)
                End If
              Next PatternCounter
              Prediction.Calculate(Patterns)
              ' Limit cache size
              Do While Predictions.Count >= MaxPointPredictions ' Limit cache
                ' For the implementation of a random cache eviction policy, just
                ' taking the first element of the dictionary is random enough as
                ' a Dictionary in .NET is implemented as a hashtable.
                Predictions.Remove(Predictions.ElementAt(0).Key)
              Loop
              ' Add calculated prediction to cache.
              Predictions.Add(PredictionTimestamp, Prediction.ShallowCopy)
            End If
            With Prediction ' Store results in arrays for EMA calculation.
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
