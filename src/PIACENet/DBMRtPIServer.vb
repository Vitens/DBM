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

Namespace DBMRt

    Public Class DBMRtPIServer

        Private PIServer As PISDK.Server
        Private PIPoints As New Collections.Generic.List(Of DBMRtPIPoint)

        Public Sub New(PIServer As PISDK.Server)
            Dim InstrTag,Substrings() As String
            Me.PIServer=PIServer
            Try
                For Each PIPoint As PISDK.PIPoint In Me.PIServer.GetPointsSQL("PIpoint.Tag='*' AND PIpoint.PointSource='dbmrt' AND PIpoint.Scan=1") ' Search for DBM output PI points
                    InstrTag=PIPoint.PointAttributes("InstrumentTag").Value.ToString
                    If Text.RegularExpressions.Regex.IsMatch(InstrTag,"^[a-zA-Z0-9_\.-]{1,}:[^:?*&]{1,}$") Then ' InstrumentTag attribute should contain input PI point
                        Substrings=InstrTag.Split(New Char(){":"c}) ' Format: PI server:PI point
                        Try
                            If Not DBMRtCalculator.PISDK.Servers(Substrings(0)).PIPoints(Substrings(1)).Name.Equals(String.Empty) Then
                                PIPoints.Add(New DBMRtPIPoint(DBMRtCalculator.PISDK.Servers(Substrings(0)).PIPoints(Substrings(1)),PIPoint)) ' Add to calculation points
                            End If
                        Catch
                        End Try
                    End If
                Next
            Catch
            End Try
        End Sub

        Public Sub Calculate
            For Each PIPoint In PIPoints
                Try
                    PIPoint.Calculate
                Catch
                End Try
            Next
        End Sub

    End Class

End Namespace
