Option Explicit
Option Strict

Public Class DBMCachedValue

    Public Timestamp As DateTime
    Public Value As Double

    Public Sub New(ByVal Timestamp As DateTime,ByVal Value As Double)
        Me.Timestamp=Timestamp
        Me.Value=Value
    End Sub

End Class
