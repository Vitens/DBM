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
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTimeSeries


    ' Contains time series functions.


    Public Shared Function TimeWeightedValue(value As Double,
      nextValue As Double, timestamp As DateTime, nextTimestamp As DateTime,
      Optional stepped As Boolean = True) As Double

      TimeWeightedValue = value

      ' When not stepped, using interpolation, the value to be used for the
      ' interval is the average value of the current value and the next value.
      If Not stepped Then
        TimeWeightedValue += nextValue
        TimeWeightedValue /= 2
      End If

      TimeWeightedValue *= nextTimestamp.Subtract(timestamp).TotalDays

      Return TimeWeightedValue

    End Function


    Public Shared Function LinearInterpolation(timestamp As DateTime,
      startTimestamp As DateTime, startValue As Double,
      endTimestamp As DateTime, endValue As Double,
      Optional stepped As Boolean = True) As Double

      If stepped Then
        Return startValue
      Else
        Return Lerp(startValue, endValue,
          (timestamp.Ticks-startTimestamp.Ticks)/
          (endTimestamp.Ticks-startTimestamp.Ticks))
      End If

    End Function


  End Class


End Namespace
