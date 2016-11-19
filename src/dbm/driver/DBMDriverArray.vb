Option Explicit
Option Strict

Public Class DBMDriver

    Public Shared UnitTestData() As Double

    Public Sub New(Optional ByVal Data() As Double=Nothing)
        UnitTestData=Data
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
        GetData=DBMDriver.UnitTestData(0)
        DBMDriver.UnitTestData=DBMDriver.UnitTestData.Skip(1).ToArray
        Return GetData
    End Function

    Public Sub PostCalculate(Optional ByVal Timestamp As DateTime=Nothing)
    End Sub

End Class
