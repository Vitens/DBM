Option Explicit
Option Strict

Public Class DBMDriver

    Public Shared Data() As Object

    Public Sub New(Optional ByVal Data() As Object=Nothing)
        DBMDriver.Data=Data
    End Sub

End Class

Public Class DBMPointDriver

    Public Point As String

    Public Sub New(Optional ByVal Point As String=Nothing)
        Me.Point=Point
    End Sub

    Public Sub PreCalculate(Optional ByVal Timestamp As DateTime=Nothing)
    End Sub

    Public Function GetData(ByVal StartTimestamp As DateTime,ByVal EndTimestamp As DateTime) As Double
        GetData=CDbl(DBMDriver.Data(0))
        DBMDriver.Data=DBMDriver.Data.Skip(1).ToArray
        Return GetData
    End Function

    Public Sub PostCalculate(Optional ByVal Timestamp As DateTime=Nothing)
    End Sub

End Class
