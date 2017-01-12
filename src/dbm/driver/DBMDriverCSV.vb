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

Namespace Vitens.DynamicBandwidthMonitor

    Public Class DBMPointDriver

        Public Point As Object
        Private Values As Collections.Generic.Dictionary(Of DateTime, Double)

        Public Sub New(Point As Object)
            Me.Point = Point
        End Sub

        Public Function GetData(StartTimestamp As DateTime, EndTimestamp As DateTime) As Double
            Dim StreamReader As System.IO.StreamReader
            Dim Substrings() As String
            Dim Timestamp As DateTime
            Dim Value As Double
            If Values Is Nothing Then ' No data in memory yet
                Values = New Collections.Generic.Dictionary(Of DateTime, Double)
                Try
                    StreamReader = New System.IO.StreamReader(DirectCast(Point, String))
                    Do While Not StreamReader.EndOfStream
                        Substrings = System.Text.RegularExpressions.Regex.Split(StreamReader.ReadLine, "^([^,\t]+)[,\t](.+)$") ' Comma and tab delimiters; split in 2 substrings
                        If Substrings.Length = 4 Then ' If a match is found at the beginning or the end of the input string, an empty string is included at the beginning or the end of the returned array
                            If DateTime.TryParse(Substrings(1), Timestamp) Then
                                If Double.TryParse(Substrings(2), Value) Then
                                    If Not Values.ContainsKey(Timestamp) Then
                                        Values.Add(Timestamp, Value)
                                    End If
                                End If
                            End If
                        End If
                    Loop
                    StreamReader.Close
                Catch
                End Try
            End If
            If Values.ContainsKey(StartTimestamp) Then ' In cache
                Return Values.Item(StartTimestamp) ' Return value from cache
            Else
                Return Double.NaN ' No data, return Not a Number
            End If
        End Function

    End Class

End Namespace
