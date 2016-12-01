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

    Public DBMPointDriver As New DBMPointDriver(Nothing)
    Private CachedValues() As DBMCachedValue
    Private CacheIndex As Integer

    Public Sub New(ByVal DBMPointDriver As DBMPointDriver)
        Me.DBMPointDriver=DBMPointDriver
        InvalidateCache
        CacheIndex=0
    End Sub

    Public Function Value(ByVal Timestamp As DateTime) As Double ' Returns value at timestamp, either from cache or using driver
        Dim i As Integer
        i=Array.FindIndex(CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp) ' Find timestamp in cache
        If i=-1 Then ' Not in cache
            i=CacheIndex
            Try
                CachedValues(i)=New DBMCachedValue(Timestamp,DBMPointDriver.GetData(Timestamp,DateAdd("s",DBMConstants.CalculationInterval,Timestamp))) ' Get data using driver
            Catch
                CachedValues(i)=New DBMCachedValue(Timestamp,Double.NaN) ' Error, return Not a Number
            End Try
            CacheIndex=(CacheIndex+1) Mod CachedValues.Length ' Increase index
        End If
        Return CachedValues(i).Value
    End Function

    Public Sub InvalidateCache(Optional ByVal Timestamp As DateTime=Nothing)
        Dim i As Integer
        If Timestamp=DateTime.MinValue Then
            ReDim CachedValues(CInt((DBMConstants.ComparePatterns+1)*(DBMConstants.EMAPreviousPeriods+DBMConstants.CorrelationPreviousPeriods+24*3600/DBMConstants.CalculationInterval)-1)) ' Large enough for at least one day
            For i=0 to CachedValues.Length-1 ' Initialise cache
                CachedValues(i)=New DBMCachedValue
            Next i
        Else
            i=Array.FindIndex(CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp)
            If i>=0 Then
                CachedValues(i).Invalidate
            End If
        End If
    End Sub

End Class
