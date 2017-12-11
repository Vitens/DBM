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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMPredictionData


    ' Stores prediction data for the DBMPrediction class.


    Public MeasuredValue, PredictedValue, LowerControlLimit, _
      UpperControlLimit As Double


    Public Sub New(Optional MeasuredValue As Double = 0, _
      Optional PredictedValue As Double = 0, _
      Optional LowerControlLimit As Double = 0, _
      Optional UpperControlLimit As Double = 0)

      ' DBMResult objects can directly store results into a DBMPredictionData
      ' object. This is the result of an EMA on predictions calculated using
      ' the Calculate method called from a DBMPoint object.

      Me.MeasuredValue = MeasuredValue
      Me.PredictedValue = PredictedValue
      Me.LowerControlLimit = LowerControlLimit
      Me.UpperControlLimit = UpperControlLimit

    End Sub


  End Class


End Namespace
