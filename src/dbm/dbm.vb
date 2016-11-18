Option Explicit
Option Strict

Imports System.Reflection
<assembly:AssemblyTitle("DBM")>
<assembly:AssemblyCompany("Vitens N.V.")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor")>
<assembly:AssemblyCopyright("J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyVersion("1.2.*")>

Public Class DBM

    Private DBMDriver As DBMDriver
    Private DBMPoints(-1) As DBMPoint

    Public Sub New(Optional ByVal Data() As Double=Nothing)
        DBMDriver=New DBMDriver(Data)
    End Sub

    Private Function DBMPointIndex(ByVal Point As DBMPointDriver) As Integer
        DBMPointIndex=Array.FindIndex(DBMPoints,Function(FindDBMPoint)FindDBMPoint.DBMPointDriver.Point Is Point.Point)
        If DBMPointIndex=-1 Then ' Point not found
            ReDim Preserve DBMPoints(DBMPoints.Length)
            DBMPointIndex=DBMPoints.Length-1
            DBMPoints(DBMPointIndex)=New DBMPoint(Point)
        End If
        Return DBMPointIndex
    End Function

    Public Function Calculate(ByVal InputPoint As DBMPointDriver,ByVal CorrelationPoint As DBMPointDriver,ByVal Timestamp As DateTime,Optional ByVal SubstractInputPointFromCorrelationPoint As Boolean=False) As DBMResult
        Dim InputDBMPointIndex,CorrelationDBMPointIndex As Integer
        Dim InputDBMResult,CorrelationDBMResult As DBMResult
        Dim AbsErrorStats,RelErrorStats As New DBMStatistics
        InputDBMPointIndex=DBMPointIndex(InputPoint)
        InputDBMResult=DBMPoints(InputDBMPointIndex).Calculate(Timestamp,True,Not IsNothing(CorrelationPoint)) ' Calculate for input point
        If InputDBMResult.Factor<>0 And Not IsNothing(CorrelationPoint) Then ' If an event is found and a correlation point is available
            CorrelationDBMPointIndex=DBMPointIndex(CorrelationPoint)
            If SubstractInputPointFromCorrelationPoint Then ' If pattern of correlation point contains input point
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointIndex)) ' Calculate for correlation point, substract input point
            Else
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True) ' Calculate for correlation point
            End If
            AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).AbsoluteError,DBMPoints(InputDBMPointIndex).AbsoluteError) ' Absolute error compared to prediction
            RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).RelativeError,DBMPoints(InputDBMPointIndex).RelativeError) ' Relative error compared to prediction
            If RelErrorStats.ModifiedCorrelation>DBMConstants.CorrelationThreshold Then ' Suppress event due to correlation of relative error
                InputDBMResult.Factor=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-DBMConstants.CorrelationThreshold Then ' Suppress event due to anticorrelation of absolute error (unmeasured supply)
                InputDBMResult.Factor=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return InputDBMResult
    End Function

End Class
