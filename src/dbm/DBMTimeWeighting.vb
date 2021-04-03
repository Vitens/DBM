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


  Public Class DBMTimeWeighting


    ' Contains time weighting functions.


    Public Shared Function WeightedValue(Value As Double, NextValue As Double,
      Timestamp As DateTime, NextTimestamp As DateTime,
      Stepped As Boolean) As Double

      WeightedValue = Value

      ' When not stepped, using interpolation, the value to be used for the
      ' interval is the average value of the current value and the next value.
      If Not Stepped Then
        WeightedValue += NextValue
        WeightedValue /= 2
      End If

      WeightedValue *= NextTimestamp.Subtract(Timestamp).TotalDays

      Return WeightedValue

    End Function


    Public Shared Function FindCentralValue(v0 As Double, v2 As Double,
      t0 As DateTime, t1 As DateTime, t2 As DateTime, Stepped As Boolean,
      w As Double) As Double

      ' Finds the required central value v1 at given time t1 so that the
      ' time-weighted total of the three points in the time range from t0 (with
      ' value v0) to t2 (with value v2) equals the given time-weighted total w.

      If Stepped Then
        ' w = v0*(t1-t0)+v1*(t2-t1)
        ' Solve for v1:
        '   v1 = (w-v0*(t1-t0))/(t2-t1)
        Return (w-v0*t1.Subtract(t0).TotalDays)/t2.Subtract(t1).TotalDays
      Else
        ' w = (v0+v1)/2*(t1-t0)+(v1+v2)/2*(t2-t1)
        ' Solve for v1:
        '   v1 = (2*w-v0*(t1-t0)-v2*(t2-t1))/(t2-t0)
        Return (2*w-v0*t1.Subtract(t0).TotalDays-
          v2*t2.Subtract(t1).TotalDays)/t2.Subtract(t0).TotalDays
      End If

    End Function


  End Class


End Namespace
