Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMDataManager

    Public DBMPointDriver As New DBMPointDriver(Nothing)
    Private CachedValues() As DBMCachedValue
    Private CacheIndex As Integer

    Public Sub New(ByVal DBMPointDriver As DBMPointDriver)
        Dim i As Integer
        Me.DBMPointDriver=DBMPointDriver
        ReDim CachedValues(CInt((DBMConstants.EMAPreviousPeriods+1+DBMConstants.CorrelationPreviousPeriods+1+24*(3600/DBMConstants.CalculationInterval))*(DBMConstants.ComparePatterns+1)-1))
        For i=0 to CachedValues.Length-1 ' Initialise cache
            CachedValues(i)=New DBMCachedValue
        Next i
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

End Class
