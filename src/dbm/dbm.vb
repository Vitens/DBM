Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' J.H. Fitié, Vitens N.V.

Imports System.Reflection
<assembly:AssemblyTitle("DBM")>
<assembly:AssemblyCompany("Vitens N.V.")>
<assembly:AssemblyProduct("Dynamic Bandwidth Monitor")>
<assembly:AssemblyCopyright("J.H. Fitié, Vitens N.V.")>
<assembly:AssemblyVersion("1.2.1.*")>

Public Class DBM

    Private DBMDriver As DBMDriver
    Private DBMPoints(-1) As DBMPoint

    Public Sub New(Optional ByVal Data() As Object=Nothing)
        DBMDriver=New DBMDriver(Data)
    End Sub

    Private Function DBMPointDriverIndex(ByVal DBMPointDriver As DBMPointDriver) As Integer
        DBMPointDriverIndex=Array.FindIndex(DBMPoints,Function(FindDBMPoint)FindDBMPoint.DBMDataManager.DBMPointDriver.Point Is DBMPointDriver.Point)
        If DBMPointDriverIndex=-1 Then ' PointDriver not found
            ReDim Preserve DBMPoints(DBMPoints.Length) ' Add to array
            DBMPointDriverIndex=DBMPoints.Length-1
            DBMPoints(DBMPointDriverIndex)=New DBMPoint(DBMPointDriver)
        End If
        Return DBMPointDriverIndex
    End Function

    Public Function Calculate(ByVal InputDBMPointDriver As DBMPointDriver,ByVal CorrelationDBMPointDriver As DBMPointDriver,ByVal Timestamp As DateTime,Optional ByVal SubstractInputPointFromCorrelationPoint As Boolean=False) As DBMResult
        Dim InputDBMPointDriverIndex,CorrelationDBMPointDriverIndex As Integer
        Dim CorrelationDBMResult As DBMResult
        Dim AbsErrorStats,RelErrorStats As New DBMStatistics
        InputDBMPointDriverIndex=DBMPointDriverIndex(InputDBMPointDriver)
        Calculate=DBMPoints(InputDBMPointDriverIndex).Calculate(Timestamp,True,Not IsNothing(CorrelationDBMPointDriver)) ' Calculate for input point
        If Calculate.Factor<>0 And Not IsNothing(CorrelationDBMPointDriver) Then ' If an event is found and a correlation point is available
            CorrelationDBMPointDriverIndex=DBMPointDriverIndex(CorrelationDBMPointDriver)
            If SubstractInputPointFromCorrelationPoint Then ' If pattern of correlation point contains input point
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointDriverIndex)) ' Calculate for correlation point, substract input point
            Else
                CorrelationDBMResult=DBMPoints(CorrelationDBMPointDriverIndex).Calculate(Timestamp,False,True) ' Calculate for correlation point
            End If
            AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).AbsoluteError,DBMPoints(InputDBMPointDriverIndex).AbsoluteError) ' Absolute error compared to prediction
            RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointDriverIndex).RelativeError,DBMPoints(InputDBMPointDriverIndex).RelativeError) ' Relative error compared to prediction
            If RelErrorStats.ModifiedCorrelation>DBMConstants.CorrelationThreshold Then ' Suppress event due to correlation of relative error
                Calculate.Factor=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-DBMConstants.CorrelationThreshold Then ' Suppress event due to anticorrelation of absolute error (unmeasured supply)
                Calculate.Factor=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return Calculate
    End Function

End Class
