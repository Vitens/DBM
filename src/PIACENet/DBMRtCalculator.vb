Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

Public Class DBMRtCalculator

    Public Shared PISDK As New PISDK.PISDK
    Public Shared DBM As New DBM
    Private DBMRtPIServers As Collections.Generic.List(Of DBMRtPIServer)

    Public Sub New(Optional ByVal DefaultServerOnly As Boolean=False,Optional ByVal TagFilter As String="*")
        DBMRtPIServers=New Collections.Generic.List(Of DBMRtPIServer)
        For Each thisServer As PISDK.Server In PISDK.Servers
            If Not DefaultServerOnly Or thisServer Is PISDK.Servers.DefaultServer Then
                DBMRtPIServers.Add(New DBMRtPIServer(thisServer,TagFilter))
            End If
        Next
    End Sub

    Public Sub Calculate
        For Each thisDBMRtPIServer As DBMRtPIServer In DBMRtPIServers
            thisDBMRtPIServer.Calculate
        Next
    End Sub

End Class
