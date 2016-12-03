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

Public Class DBMDataManager

    Public DBMPointDriver As DBMPointDriver
    Private CacheIndex As Integer
    Private DBMCachedValues As New Collections.Generic.List(Of DBMCachedValue)

    Public Sub New(ByVal DBMPointDriver As DBMPointDriver)
        Me.DBMPointDriver=DBMPointDriver
        CacheIndex=0
    End Sub

    Public Function Value(ByVal Timestamp As DateTime) As Double ' Returns value at timestamp, either from cache or using driver
        Dim i As Integer
        i=DBMCachedValues.FindIndex(Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp) ' Find timestamp in cache
        If i=-1 Then ' Not in cache
            Try
                Value=DBMPointDriver.GetData(Timestamp,DateAdd("s",DBMConstants.CalculationInterval,Timestamp)) ' Get data using driver
            Catch
                Value=Double.NaN ' Error, return Not a Number
            End Try
            If DBMCachedValues.Count<DBMConstants.MaximumCacheSize Then ' Limit cache size (large enough for at least one day)
                DBMCachedValues.Add(New DBMCachedValue(Timestamp,Value)) ' Add to cache (new)
            Else
                DBMCachedValues(CacheIndex)=New DBMCachedValue(Timestamp,Value) ' Add to cache (existing)
                CacheIndex=(CacheIndex+1) Mod DBMCachedValues.Count ' Increase index
            End If
        Else
            Value=DBMCachedValues(i).Value ' Return value from cache
        End If
        Return Value
    End Function

End Class
