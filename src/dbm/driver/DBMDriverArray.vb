Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Public Class DBMDriver

    Public Shared Data(-1) As Double
    Public Shared DataIndex As Integer

    Public Sub New(Optional ByVal Data() As Object=Nothing)
        Dim i As Integer
        ReDim DBMDriver.Data(Data.Length-1)
        For i=0 To Data.Length-1
            DBMDriver.Data(i)=CDbl(Data(i))
        Next i
        DBMDriver.DataIndex=0
    End Sub

End Class

Public Class DBMPointDriver

    Public Point As String

    Public Sub New(ByVal Point As String)
        Me.Point=Point
    End Sub

    Public Function GetData(ByVal StartTimestamp As DateTime,ByVal EndTimestamp As DateTime) As Double
        GetData=DBMDriver.Data(DBMDriver.DataIndex) ' Return item from data array passed to DBMDriver
        DBMDriver.DataIndex=(DBMDriver.DataIndex+1) Mod DBMDriver.Data.Length ' Increase index
        Return GetData
    End Function

End Class
