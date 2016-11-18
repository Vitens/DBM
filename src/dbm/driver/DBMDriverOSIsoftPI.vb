Option Explicit
Option Strict

Public Class DBMDriver

    Public Sub New(Optional ByVal Data() As Double=Nothing)
    End Sub

End Class

Public Class DBMPointDriver

    Public Point As PISDK.PIPoint

    Public Sub New(Optional ByVal Point As PISDK.PIPoint=Nothing)
        Me.Point=Point
    End Sub

    Public Function GetData(ByVal StartTimestamp As DateTime,ByVal EndTimestamp As DateTime) As Double
        Return CDbl(Point.Data.Summary(StartTimestamp,EndTimestamp,PISDK.ArchiveSummaryTypeConstants.astAverage,PISDK.CalculationBasisConstants.cbTimeWeighted).Value)
    End Function

End Class
