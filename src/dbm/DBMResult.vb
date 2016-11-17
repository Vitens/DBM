Option Explicit
Option Strict

Public Class DBMResult

    Public Factor,CurrValue,PredValue,LowContrLimit,UppContrLimit As Double

    Public Sub New(ByVal Factor As Double,ByVal CurrValue As Double,ByVal PredValue As Double,ByVal LowContrLimit As Double,ByVal UppContrLimit As Double)
        Me.Factor=Factor
        Me.CurrValue=CurrValue
        Me.PredValue=PredValue
        Me.LowContrLimit=LowContrLimit
        Me.UppContrLimit=UppContrLimit
    End Sub

End Class
