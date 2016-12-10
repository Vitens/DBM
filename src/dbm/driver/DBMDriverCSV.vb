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

Namespace DBM

    Public Class DBMPointDriver

        Public Point As Object
        Private DBMCachedValues As New Collections.Generic.List(Of DBMCachedValue)

        Public Sub New(Point As Object)
            Me.Point=Point
        End Sub

        Public Function GetData(StartTimestamp As DateTime,EndTimestamp As DateTime) As Double
            Dim StreamReader As System.IO.StreamReader
            Dim Line,Fields() As String
            Dim i As Integer
            If DBMCachedValues.Count=0 Then ' No data yet
                Try
                    StreamReader=New System.IO.StreamReader(CStr(Point))
                    While Not StreamReader.EndOfStream
                        Line=StreamReader.ReadLine()
                        Try
                            Fields=Split(Line,",") ' CSV delimiter
                            DBMCachedValues.Add(New DBMCachedValue(Convert.ToDateTime(Fields(0)),Convert.ToDouble(Fields(1)))) ' timestamp,value
                        Catch
                        End Try
                    End While
                    StreamReader.Close
                Catch
                End Try
            End If
            i=DBMCachedValues.FindIndex(Function(FindCachedValue)FindCachedValue.Timestamp=StartTimestamp) ' Find timestamp in cache
            If i=-1 Then ' Not in cache
                GetData=Double.NaN ' No data, return Not a Number
            Else
                GetData=DBMCachedValues.Item(i).Value ' Return value from cache
            End If
            Return GetData
        End Function

    End Class

End Namespace
