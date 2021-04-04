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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTimeSeries


    ' Contains time series functions.


    Public Shared Function TimeWeight(Timestamp As DateTime,
      NextTimestamp As DateTime) As Double

      Return NextTimestamp.Subtract(Timestamp).TotalDays

    End Function


    Public Shared Function TimeWeightedValue(Value As Double,
      NextValue As Double, Timestamp As DateTime, NextTimestamp As DateTime,
      Stepped As Boolean) As Double

      TimeWeightedValue = Value

      ' When not stepped, using interpolation, the value to be used for the
      ' interval is the average value of the current value and the next value.
      If Not Stepped Then
        TimeWeightedValue += NextValue
        TimeWeightedValue /= 2
      End If

      TimeWeightedValue *= TimeWeight(Timestamp, NextTimestamp)

      Return TimeWeightedValue

    End Function


    Public Shared Function FindCentralValue(PreviousValue As Double,
      NextValue As Double, PreviousTimestamp As DateTime, Timestamp As DateTime,
      NextTimestamp As DateTime, Stepped As Boolean,
      TotalWeight As Double) As Double

      ' Finds the required central value v1 at given time t1 so that the
      ' time-weighted total of the three points in the time range from t0 (with
      ' value v0) to t2 (with value v2) equals the given time-weighted total w.

      If Stepped Then
        ' w = v0*(t1-t0)+v1*(t2-t1)
        ' Solve for v1:
        '   v1 = (w-v0*(t1-t0))/(t2-t1)
        Return (TotalWeight-TimeWeightedValue(PreviousValue, Nothing,
          PreviousTimestamp, Timestamp, True))/TimeWeight(Timestamp,
          NextTimestamp)
      Else
        ' w = (v0+v1)/2*(t1-t0)+(v1+v2)/2*(t2-t1)
        ' Solve for v1:
        '   v1 = (2*w-v0*(t1-t0)-v2*(t2-t1))/(t2-t0)
        Return (2*TotalWeight-TimeWeightedValue(PreviousValue, Nothing,
          PreviousTimestamp, Timestamp, True)-TimeWeightedValue(NextValue,
          Nothing, Timestamp, NextTimestamp, True))/
          TimeWeight(PreviousTimestamp, NextTimestamp)
      End If

    End Function


    Public Shared Function InterpolatedValue(PreviousValue As Double,
      NextValue As Double, PreviousTimestamp As DateTime, Timestamp As DateTime,
      NextTimestamp As DateTime, Stepped As Boolean) As Double

      ' Returns the interpolated value v1 at given time t1 between the points at
      ' times t0 (with value v0) and t2 (with value v2).

      If Stepped Then
        ' v1 = v0
        Return PreviousValue
      Else
        ' v1 = v0+(v2-v0)/(t2-t0)*(t1-t0)
        Return PreviousValue+(NextValue-PreviousValue)/
          TimeWeight(PreviousTimestamp, NextTimestamp)*
          TimeWeight(PreviousTimestamp, Timestamp)
      End If

    End Function


  End Class


End Namespace
