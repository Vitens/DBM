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

    Public Class DBMDataManager

        Public DBMPointDriver As DBMPointDriver
        Private CachedValues As New Collections.Generic.Dictionary(Of DateTime,Double)

        Public Sub New(DBMPointDriver As DBMPointDriver)
            Me.DBMPointDriver=DBMPointDriver
        End Sub

        Public Function Value(Timestamp As DateTime) As Double ' Returns value at timestamp, either from cache or using driver
            If CachedValues.ContainsKey(Timestamp) Then ' In cache
                Value=CachedValues.Item(Timestamp) ' Return value from cache
            Else
                If DBMUnitTests.UnitTestsRunning Then ' Do not use point driver when running unit tests
                    Value=DBMUnitTests.Data(DBMUnitTests.DataIndex) ' Return item from unit tests data array
                    DBMUnitTests.DataIndex=(DBMUnitTests.DataIndex+1) Mod DBMUnitTests.Data.Length ' Increase index
                Else
                    Try
                        Value=DBMPointDriver.GetData(Timestamp,DateAdd("s",DBMParameters.CalculationInterval,Timestamp)) ' Get data using driver
                    Catch
                        Value=Double.NaN ' Error, return Not a Number
                    End Try
                End If
                If CachedValues.Count>=DBMParameters.MaximumCacheSize Then ' Limit cache size
                    CachedValues.Remove(CachedValues.ElementAt(CInt(Math.Floor(Rnd()*CachedValues.Count))).Key) ' Remove random cached value
                End If
                CachedValues.Add(Timestamp,Value) ' Add to cache
            End If
            Return Value
        End Function

    End Class

End Namespace
