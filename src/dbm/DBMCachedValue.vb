Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMCachedValue

    Public Timestamp As DateTime
    Public Value As Double

    Public Sub New(ByVal Timestamp As DateTime,ByVal Value As Double)
        Me.Timestamp=Timestamp
        Me.Value=Value
    End Sub

End Class
