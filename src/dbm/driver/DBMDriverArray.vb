Option Explicit
Option Strict

Public Class DBMDriver

    Public Shared Data(-1) As Double

    Public Sub New(Optional ByVal Data() As Object=Nothing)
        For Each Value In Data
            ReDim Preserve DBMDriver.Data(DBMDriver.Data.Length)
            DBMDriver.Data(DBMDriver.Data.Length-1)=CDbl(Value)
        Next
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
        GetData=DBMDriver.Data(0)
        DBMDriver.Data=DBMDriver.Data.Skip(1).ToArray
        Return GetData
    End Function

    Public Sub PostCalculate(Optional ByVal Timestamp As DateTime=Nothing)
    End Sub

End Class
