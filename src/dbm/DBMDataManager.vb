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

Namespace DBM

    Public Class DBMDataManager

        Public PointDriver As DBMPointDriver
        Private Values As New Collections.Generic.Dictionary(Of DateTime,Double)

        Public Sub New(PointDriver As DBMPointDriver)
            Me.PointDriver=PointDriver
        End Sub

        Public Function Value(Timestamp As DateTime) As Double ' Returns value at timestamp, either from cache or using driver
            If Values.ContainsKey(Timestamp) Then ' In cache
                Value=Values.Item(Timestamp) ' Return value from cache
            Else
                If DBMUnitTests.UnitTestsRunning Then ' Do not use point driver when running unit tests
                    Value=DBMUnitTests.Data(DBMUnitTests.DataIndex) ' Return item from unit tests data array
                    DBMUnitTests.DataIndex=(DBMUnitTests.DataIndex+1) Mod DBMUnitTests.Data.Length ' Increase index
                Else
                    Try
                        Value=PointDriver.GetData(Timestamp,DateAdd(DateInterval.Second,DBMParameters.CalculationInterval,Timestamp)) ' Get data using driver
                    Catch
                        Value=Double.NaN ' Error, return Not a Number
                    End Try
                End If
                Do While Values.Count>=DBMParameters.MaxDataManagerValues ' Limit cache size
                    Values.Remove(Values.ElementAt(CInt(Math.Floor(Rnd*Values.Count))).Key) ' Remove random cached value
                Loop
                Values.Add(Timestamp,Value) ' Add to cache
            End If
            Return Value
        End Function

    End Class

End Namespace
