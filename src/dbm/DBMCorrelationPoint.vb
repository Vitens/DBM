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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMCorrelationPoint


    ' Contains a DBMPointDriverAbstract object and a boolean SubtractSelf which
    ' can be set to true when the input tag has to be subtracted from the
    ' correlation tag, for example when the correlation tag contains the input
    ' tag. Set to false for adjacent areas.


    Private _pointDriver As DBMPointDriverAbstract
    Private _subtractSelf As Boolean ' True if input needs to be subtracted


    Public Property PointDriver As DBMPointDriverAbstract
      Get
        Return _pointDriver
      End Get
      Set(value As DBMPointDriverAbstract)
        _pointDriver = value
      End Set
    End Property


    Public Sub New(PointDriver As DBMPointDriverAbstract,
      SubtractSelf As Boolean)

      _pointDriver = PointDriver
      _subtractSelf = SubtractSelf

    End Sub


  End Class


End Namespace
