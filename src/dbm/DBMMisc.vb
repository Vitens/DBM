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


Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMMisc


    ' Contains miscellaneous functions.


    Public Shared Function PIAFIntervalSeconds(NumberOfValues As Integer,
      DurationSeconds As Double) As Double

      ' OSIsoft PI AF specific: Number of values desired. If 0, all intervals
      ' will be returned. If >0, that number of values will be returned. If <0,
      ' the negative value minus 1 number of values will be returned (f.ex. -25
      ' over a 24 hour period will return an hourly value). Duration is in
      ' seconds.

      If NumberOfValues < 0 Then NumberOfValues = -NumberOfValues-1
      If NumberOfValues = 1 Then
        Return DurationSeconds ' Return a single value
      Else
        Return Max(1, (DurationSeconds/CalculationInterval-1)/
          (NumberOfValues-1))*CalculationInterval ' Required interval
      End If

    End Function


  End Class


End Namespace
