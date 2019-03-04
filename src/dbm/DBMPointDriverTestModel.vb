Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
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


  Public Class DBMPointDriverTestModel
    Inherits DBMPointDriverAbstract


    Public Sub New(Point As Object)

      MyBase.New(Point)

    End Sub


    Public Overrides Function GetData(Timestamp As DateTime) As Double

      ' Model based on hourly water usage in Leeuwarden (NL) 2016.
      ' Calculated using polynomial regressions based on hourly (quintic),
      ' daily (cubic) and monthly (quartic) periodicity.

      If TypeOf Point Is Integer Then ' Point contains offset in hours
        Timestamp = Timestamp.AddHours(DirectCast(Point, Integer))
      End If

      With Timestamp
        Return 790*(-0.00012*.Month^4+0.0035*.Month^3-0.032*.Month^2+0.1*
          .Month+0.93)*(0.000917*.DayOfWeek^3-0.0155*.DayOfWeek^2+0.0628*
          .DayOfWeek+0.956)*(-0.00001221*(.Hour+.Minute/60)^5+0.0007805*
          (.Hour+.Minute/60)^4-0.01796*(.Hour+.Minute/60)^3+0.1709*(.Hour+
          .Minute/60)^2-0.5032*(.Hour+.Minute/60)+0.7023)
      End With

    End Function


  End Class


End Namespace
