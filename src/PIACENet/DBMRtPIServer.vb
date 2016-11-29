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

Public Class DBMRtPIServer

    Private PIServer As PISDK.Server
    Private DBMRtPIPoints As Collections.Generic.List(Of DBMRtPIPoint)

    Public Sub New(ByVal PIServer As PISDK.Server,Optional ByVal TagFilter As String="*")
        Dim InstrTag,Fields() As String
        Me.PIServer=PIServer
        DBMRtPIPoints=New Collections.Generic.List(Of DBMRtPIPoint)
        Try
            For Each thisDBMRtPIPoint As PISDK.PIPoint In Me.PIServer.GetPointsSQL("PIpoint.Tag='" & TagFilter & "' AND PIpoint.PointSource='dbmrt' AND PIpoint.Scan=1")
                InstrTag=thisDBMRtPIPoint.PointAttributes("InstrumentTag").Value.ToString
                If Text.RegularExpressions.Regex.IsMatch(InstrTag,"^[a-zA-Z0-9_\.-]{1,}:[^:?*&]{1,}$") Then
                    Fields=Split(InstrTag.ToString,":")
                    Try
                        If DBMRtCalculator.PISDK.Servers(Fields(0)).PIPoints(Fields(1)).Name<>"" Then
                            DBMRtPIPoints.Add(New DBMRtPIPoint(DBMRtCalculator.PISDK.Servers(Fields(0)).PIPoints(Fields(1)),thisDBMRtPIPoint))
                        End If
                    Catch
                    End Try
                End If
            Next
        Catch
        End Try
    End Sub

    Public Sub CalculatePoints
        For Each thisDBMRtPIPoint As DBMRtPIPoint In DBMRtPIPoints
            Try
                thisDBMRtPIPoint.CalculatePoint
            Catch
            End Try
        Next
    End Sub

End Class
