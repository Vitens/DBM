Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMFunctions

    Public Function ArrayRotateLeft(ByVal Data() As Double) As Double() ' Rotate array left, first item becomes last
        Array.Reverse(Data) ' ABCDE -> EDCBA
        Array.Reverse(Data,0,Data.Length-1) ' EDCBA -> BCDEA
        Return Data
    End Function

    Public Function ArrayRotateRight(ByVal Data() As Object) As Object() ' Rotate array right, last item becomes first
        Array.Reverse(Data,0,Data.Length-1) ' ABCDE -> DCBAE
        Array.Reverse(Data) ' DCBAE -> EABCD
        Return Data
    End Function

    Public Function ArrayMoveItemToFront(ByVal Data() As Object,ByVal Item As Integer) As Object() ' Move item to front
        Array.Reverse(Data,0,Item)
        Array.Reverse(Data,0,Item+1)
        Return Data
    End Function

End Class
