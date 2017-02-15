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

Imports Vitens.DynamicBandwidthMonitor.DBMParameters

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMResult

    ' DBMResult is the object passed by DBM.Result and contains the final results
    ' for the DBM calculation. It is also used for storing intermediate results of
    ' (correlation) calculation results from a DBMPoint object.

    Public Prediction As DBMPrediction
    Public Factor, OriginalFactor, AbsoluteErrors(), RelativeErrors(), _
      CorrelationAbsoluteErrors(), CorrelationRelativeErrors() As Double
    Public AbsoluteErrorStats, RelativeErrorStats As New DBMStatistics
    Public SuppressedBy As DBMPointDriver

    Public Sub New
      ' Initialize array sizes. To be filled from back to front.
      ReDim AbsoluteErrors(CorrelationPreviousPeriods)
      ReDim RelativeErrors(CorrelationPreviousPeriods)
      ReDim CorrelationAbsoluteErrors(CorrelationPreviousPeriods)
      ReDim CorrelationRelativeErrors(CorrelationPreviousPeriods)
    End Sub

    Public Sub Calculate(Index As Integer, MeasuredValueEMA As Double, _
      PredictedValueEMA As Double, LowerControlLimitEMA As Double, _
      UpperControlLimitEMA As Double)
      ' Calculates and stores prediction errors and initial results
      ' Absolute prediction error (for prediction error
      ' correlation calculations)
      AbsoluteErrors(Index) = PredictedValueEMA-MeasuredValueEMA
      ' Relative prediction error (for prediction error
      ' correlation calculations)
      RelativeErrors(Index) = PredictedValueEMA/MeasuredValueEMA-1
      ' Store initial (no time offset because of prediction error
      ' correlation calculations) results
      If Prediction Is Nothing Then
        ' Store EMA results in new DBMPrediction object
        Prediction = New DBMPrediction(MeasuredValueEMA, PredictedValueEMA, _
          LowerControlLimitEMA, UpperControlLimitEMA)
        ' Lower control limit exceeded
        If MeasuredValueEMA < LowerControlLimitEMA Then
          Factor = (PredictedValueEMA-MeasuredValueEMA)/ _
            (LowerControlLimitEMA-PredictedValueEMA)
        ' Upper control limit exceeded
        ElseIf MeasuredValueEMA > UpperControlLimitEMA Then
          Factor = (MeasuredValueEMA-PredictedValueEMA)/ _
            (UpperControlLimitEMA-PredictedValueEMA)
        End If
        ' Store original factor before possible suppression
        OriginalFactor = Factor
      End If
    End Sub

  End Class

End Namespace
