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

Imports Vitens.DynamicBandwidthMonitor.DBMMath

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMPointDriver

    ' Description: Model based on hourly water usage in Leeuwarden 2016.
    ' Identifier (Point): Object (unused)
    ' Remarks: Model based on measured data with added random noise.

    Public Point As Object

    Public Sub New(Point As Object)
      Me.Point = Point ' Required, but unused.
    End Sub

    Public Function GetData(StartTimestamp As DateTime, _
      EndTimestamp As DateTime) As Double
      ' Model based on hourly water usage in Leeuwarden 2016 (+/-10% random
      ' noise). Calculated using polynomial regressions based on
      ' hourly (quintic), daily (cubic) and monthly (quartic) periodicity.
      With StartTimestamp
        Return RandomNumber(9000, 11000)/10000*790* _
          (-0.00012*.Month^4+0.0035*.Month^3-0.032*.Month^2+0.1*.Month+0.93)* _
          (0.000917*.DayOfWeek^3-0.0155*.DayOfWeek^2+0.0628*.DayOfWeek+0.956)* _
          (-0.00001221*(.Hour+.Minute/60)^5+0.0007805*(.Hour+.Minute/60)^4- _
          0.01796*(.Hour+.Minute/60)^3+0.1709*(.Hour+.Minute/60)^2- _
          0.5032*(.Hour+.Minute/60)+0.7023)
      End With
    End Function

  End Class

End Namespace
