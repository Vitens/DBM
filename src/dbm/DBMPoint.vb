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

Imports System.Collections.Generic
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMPoint

    Public DataManager As DBMDataManager
    Private Predictions As New Dictionary(Of DateTime, DBMPrediction)
    Private PredictionsSubtractPoint As DBMPoint

    Public Sub New(PointDriver As DBMPointDriver)
      DataManager = New DBMDataManager(PointDriver)
    End Sub

    Public Function Result(Timestamp As DateTime, IsInputDBMPoint As Boolean, _
      HasCorrelationDBMPoint As Boolean, _
      Optional SubtractPoint As DBMPoint = Nothing) As DBMResult
      Dim CorrelationCounter, EMACounter, PatternCounter As Integer
      Dim PredictionTimestamp As DateTime
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
          Not IsInputDBMPoint Then
          For EMACounter = 0 To EMAPreviousPeriods
            PredictionTimestamp = Timestamp.AddSeconds _
              (-(EMAPreviousPeriods-EMACounter+CorrelationCounter)* _
              CalculationInterval)
            If Predictions.ContainsKey(PredictionTimestamp) Then ' From cache
              Prediction = Predictions.Item(PredictionTimestamp).ShallowCopy
            Else ' Calculate data
              For PatternCounter = 0 To ComparePatterns
                Patterns(PatternCounter) = DataManager.Value _
                (PredictionTimestamp.AddDays _
                (-(ComparePatterns-PatternCounter)*7))
                If SubtractPoint IsNot Nothing Then
                  Patterns(PatternCounter) -= _
                    SubtractPoint.DataManager.Value _
                    (PredictionTimestamp.AddDays _
                    (-(ComparePatterns-PatternCounter)*7))
                End If
              Next PatternCounter
              Prediction.Calculate(Patterns)
              ' Limit cache size
              Do While Predictions.Count >= MaxPointPredictions
              ' Remove random cached value
                Predictions.Remove(Predictions.ElementAt _
                  (RandomNumber(0, Predictions.Count-1)).Key)
              Loop
              ' Add to cache
              Predictions.Add(PredictionTimestamp, Prediction.ShallowCopy)
            End If
            With Prediction
              MeasuredValues(EMACounter) = .MeasuredValue
              PredictedValues(EMACounter) = .PredictedValue
              LowerControlLimits(EMACounter) = .LowerControlLimit
              UpperControlLimits(EMACounter) = .UpperControlLimit
            End With
          Next EMACounter
          Result.Calculate(CorrelationCounter, _
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
