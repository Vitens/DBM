Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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


    Private _timestamp As DateTime
    Private _isFutureData As Boolean
    Private _forecastItem As DBMForecastItem
    Private _factor, _absoluteErrors(), _relativeErrors() As Double


    Public Property Timestamp As DateTime
      Get
        Return _timestamp
      End Get
      Set(value As DateTime)
        _timestamp = value
      End Set
    End Property


    Public Property IsFutureData As Boolean
      Get
        Return _isFutureData
      End Get
      Set(value As Boolean)
        _isFutureData = value
      End Set
    End Property


    Public Property ForecastItem As DBMForecastItem
      Get
        Return _forecastItem
      End Get
      Set(value As DBMForecastItem)
        _forecastItem = value
      End Set
    End Property


    Public Property Factor As Double
      Get
        Return _factor
      End Get
      Set(value As Double)
        _factor = value
      End Set
    End Property


    Public Property AbsoluteErrors() As Double
      Get
        Return _absoluteErrors
      End Get
      Set(values() As Double)
        _absoluteErrors = values
      End Set
    End Property


    Public Property RelativeErrors() As Double
      Get
        Return _relativeErrors
      End Get
      Set(values() As Double)
        _relativeErrors = values
      End Set
    End Property


    Public Sub New

      ' Initialize array sizes. To be filled from back to front. When a
      ' correlation calculation is not required (no event detected or no
      ' correlation points specified) only the last item in the arrays contains
      ' a value.

      ReDim Me.AbsoluteErrors(CorrelationPreviousPeriods)
      ReDim Me.RelativeErrors(CorrelationPreviousPeriods)

    End Sub


    Public Function HasEvent As Boolean

      ' Returns true if there is an event.

      Return Abs(Me.Factor) > 1

    End Function


    Public Function HasSuppressedEvent As Boolean

      ' Returns true if there is a suppressed event.

      With Me.ForecastItem
        Return Me.Factor = 0 And
          (.Measurement < .LowerControlLimit Or
          .Measurement > .UpperControlLimit)
      End With

    End Function


    Public Function TimestampIsValid As Boolean

      ' Returns true if the timestamp is valid. When DST goes in effect, there
      ' is a missing hour in local time.

      Return Not TimeZoneInfo.Local.IsInvalidTime(Me.Timestamp)

    End Function


    Public Sub Calculate(index As Integer, measurementEMA As Double,
      forecastEMA As Double, lowerControlLimitEMA As Double,
      upperControlLimitEMA As Double)

      ' Calculates and stores forecast errors and initial results.

      ' Forecast error (for forecast error correlation calculations).
      Me.AbsoluteErrors(index) = forecastEMA-measurementEMA
      Me.RelativeErrors(index) = forecastEMA/measurementEMA-1

      If Me.ForecastItem Is Nothing Then

        ' Store EMA results in new DBMForecastItem object for current timestamp.
        Me.ForecastItem = New DBMForecastItem
        With Me.ForecastItem
          .Measurement = measurementEMA
          .Forecast = forecastEMA
          .LowerControlLimit = lowerControlLimitEMA
          .UpperControlLimit = upperControlLimitEMA
        End With

        ' No factor if there is no valid measurement.
        If IsNaN(measurementEMA) Then
          Me.Factor = NaN
        End If

        ' Lower control limit exceeded, calculate factor.
        If measurementEMA < lowerControlLimitEMA Then
          Me.Factor = (forecastEMA-measurementEMA)/
            (lowerControlLimitEMA-forecastEMA)
        End If

        ' Upper control limit exceeded, calculate factor.
        If measurementEMA > upperControlLimitEMA Then
          Me.Factor = (measurementEMA-forecastEMA)/
            (upperControlLimitEMA-forecastEMA)
        End If

      End If

    End Sub


  End Class


End Namespace
