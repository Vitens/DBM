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
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMPrediction
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMResult


    ' DBMResult is the object passed by DBM.Result and contains the final
    ' results for the DBM calculation. It is also used for storing intermediate
    ' results of (correlation) calculation results from a DBMPoint object.


    Public Timestamp As DateTime
    Public PredictionData As DBMPredictionData
    Public Factor, AbsoluteErrors(), RelativeErrors() As Double


    Public Sub New

      ' Initialize array sizes. To be filled from back to front. When a
      ' correlation calculation is not required (no exception detected or
      ' no correlation points specified) only the last item in the arrays
      ' contains a value.

      ReDim AbsoluteErrors(CorrelationPreviousPeriods)
      ReDim RelativeErrors(CorrelationPreviousPeriods)

    End Sub


    Public Sub Calculate(Index As Integer, MeasuredValueEMA As Double, _
      PredictedValueEMA As Double, LowerControlLimitEMA As Double, _
      UpperControlLimitEMA As Double)

      ' Calculates and stores prediction errors and initial results.

      ' Prediction error (for prediction error correlation calculations).
      AbsoluteErrors(Index) = PredictedValueEMA-MeasuredValueEMA
      RelativeErrors(Index) = PredictedValueEMA/MeasuredValueEMA-1

      If PredictionData Is Nothing Then

        ' Store EMA results in new DBMPredictionData object for
        ' current timestamp.
        PredictionData = New DBMPredictionData
        With PredictionData
          .MeasuredValue = MeasuredValueEMA
          .PredictedValue = PredictedValueEMA
          .LowerControlLimit = LowerControlLimitEMA
          .UpperControlLimit = UpperControlLimitEMA
        End With

        ' Lower control limit exceeded, calculate factor.
        If MeasuredValueEMA < LowerControlLimitEMA Then
          Factor = (PredictedValueEMA-MeasuredValueEMA)/ _
            (LowerControlLimitEMA-PredictedValueEMA)
        End If

        ' Upper control limit exceeded, calculate factor.
        If MeasuredValueEMA > UpperControlLimitEMA Then
          Factor = (MeasuredValueEMA-PredictedValueEMA)/ _
            (UpperControlLimitEMA-PredictedValueEMA)
        End If

      End If

    End Sub


  End Class


End Namespace
