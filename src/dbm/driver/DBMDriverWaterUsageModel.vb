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

    Public Point As Object

    Public Sub New(Point As Object)
      Me.Point = Point
    End Sub

    Public Function GetData(StartTimestamp As DateTime, _
      EndTimestamp As DateTime) As Double
      ' Model based on hourly water usage in
      ' Leeuwarden 2013 (+/-2% random noise)
      With StartTimestamp
        Return RandomNumber(9800, 10200)/10000*738.419926* _
          (-1.43724664E-05*.Hour^5+9.00975758E-04*.Hour^4- _
          2.03426424E-02*.Hour^3+1.90423208E-01*.Hour^2- _
          5.57660998E-01*.Hour+7.15188897E-01)* _
          (-5.01427884E-05*.DayOfWeek^5-2.84447601E-04*.DayOfWeek^4+ _
          8.68516927E-03*.DayOfWeek^3-4.82186003E-02*.DayOfWeek^2+ _
          1.02809577E-01*.DayOfWeek+9.51091760E-01)* _
          (-1.97603192E-05*.Month^5+3.81402902E-04*.Month^4 _
          -4.72854346E-04*.Month^3-2.42464352E-02*.Month^2 _
          +1.20188691E-01*.Month+8.80861005E-01)
      End With
    End Function

  End Class

End Namespace
