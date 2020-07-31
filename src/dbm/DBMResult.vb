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
Imports System.Double
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMResult


    ' DBMResult is the object passed by DBM.GetResult and contains the final
    ' results for the DBM calculation. It is also used for storing intermediate
    ' results of (correlation) calculation results from a DBMPoint object.


    Public Timestamp As DateTime
    Public ForecastItem As DBMForecastItem
    Public Factor, AbsoluteErrors(), RelativeErrors() As Double


    Public Sub New

      ' Initialize array sizes. To be filled from back to front. When a
      ' correlation calculation is not required (no event detected or no
      ' correlation points specified) only the last item in the arrays contains
      ' a value.

      ReDim AbsoluteErrors(CorrelationPreviousPeriods)
      ReDim RelativeErrors(CorrelationPreviousPeriods)

    End Sub


    Public Function HasEvent As Boolean

      ' Returns true if there is an event.

      Return Abs(Factor) > 1

    End Function


    Public Function HasSuppressedEvent As Boolean

      ' Returns true if there is a suppressed event.

      With ForecastItem
        Return Factor = 0 And
          (.Measurement < .LowerControlLimit Or
          .Measurement > .UpperControlLimit)
      End With

    End Function


    Public Sub Calculate(Index As Integer, MeasurementEMA As Double,
      ForecastEMA As Double, LowerControlLimitEMA As Double,
      UpperControlLimitEMA As Double)

      ' Calculates and stores forecast errors and initial results.

      ' Forecast error (for forecast error correlation calculations).
      AbsoluteErrors(Index) = ForecastEMA-MeasurementEMA
      RelativeErrors(Index) = ForecastEMA/MeasurementEMA-1

      If ForecastItem Is Nothing Then

        ' Store EMA results in new DBMForecastItem object for current timestamp.
        ForecastItem = New DBMForecastItem
        With ForecastItem
          .Measurement = MeasurementEMA
          .Forecast = ForecastEMA
          .LowerControlLimit = LowerControlLimitEMA
          .UpperControlLimit = UpperControlLimitEMA
        End With

        ' No factor if there is no valid measurement.
        If IsNaN(MeasurementEMA) Then
          Factor = NaN
        End If

        ' Lower control limit exceeded, calculate factor.
        If MeasurementEMA < LowerControlLimitEMA Then
          Factor = (ForecastEMA-MeasurementEMA)/
            (LowerControlLimitEMA-ForecastEMA)
        End If

        ' Upper control limit exceeded, calculate factor.
        If MeasurementEMA > UpperControlLimitEMA Then
          Factor = (MeasurementEMA-ForecastEMA)/
            (UpperControlLimitEMA-ForecastEMA)
        End If

      End If

    End Sub


  End Class


End Namespace
