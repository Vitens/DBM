Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTimeSeries


    ' Contains time series functions.


    Public Shared Function TimeWeightedValue(Value As Double,
      NextValue As Double, Timestamp As DateTime, NextTimestamp As DateTime,
      Optional Stepped As Boolean = True) As Double

      TimeWeightedValue = Value

      ' When not stepped, using interpolation, the value to be used for the
      ' interval is the average value of the current value and the next value.
      If Not Stepped Then
        TimeWeightedValue += NextValue
        TimeWeightedValue /= 2
      End If

      TimeWeightedValue *= NextTimestamp.Subtract(Timestamp).TotalDays

      Return TimeWeightedValue

    End Function


    Public Shared Function LinearInterpolation(Timestamp As DateTime,
      StartTimestamp As DateTime, StartValue As Double,
      EndTimestamp As DateTime, EndValue As Double,
      Optional Stepped As Boolean = True) As Double

      If Stepped Then
        Return StartValue
      Else
        Return StartValue+(Timestamp-StartTimestamp)/
          (EndTimestamp-StartTimestamp)*(EndValue-StartValue)
      End If

    End Function


  End Class


End Namespace
