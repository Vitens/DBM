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


Imports System.Collections.Generic


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMRtCalculator


    Public Shared PISDK As New PISDK.PISDK
    Public Shared DBM As New DBM
    Private PIServers As New List(Of DBMRtPIServer)


    Public Sub New
      ' Add default PI server when instantiating new calculator object.
      PIServers.Add(New DBMRtPIServer(PISDK.Servers.DefaultServer))
      ' (More PI servers could be added to the list here if required)
    End Sub


    Public Sub Calculate
      ' Perform calculation for each PI server.
      For Each PIServer In PIServers
        PIServer.Calculate
      Next
    End Sub


  End Class


End Namespace
