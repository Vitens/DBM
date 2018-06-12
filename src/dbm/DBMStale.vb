Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2018  J.H. Fitié, Vitens N.V.
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
Imports System.DateTime
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStale


    Private TimeOut As DateTime ' Stale by default


    Public Function IsStale As Boolean

      ' Returns True if this object has turned stale. The timeout will be
      ' automatically updated to the next interval if required.

      IsStale = Now >= TimeOut ' True if past timeout
      If IsStale Then TimeOut = NextInterval(Now) ' Update timeout when stale

      Return IsStale

    End Function


  End Class


End Namespace
