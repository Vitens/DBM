Option Explicit
Option Strict

Public Class DBM

    #If OfflineUnitTests Then
    Public Shared UnitTestData() As Double
    #End If

    Private DBMPoints(-1) As DBMPoint

    #If OfflineUnitTests Then
    Public Sub New(Optional ByVal Data() As Double=Nothing)
        If Not IsNothing(Data) Then
            UnitTestData=Data
        End If
    End Sub
    #End If

    #If OfflineUnitTests Then
    Private Function DBMPointIndex(ByVal Point As String) As Integer
    #Else
    Private Function DBMPointIndex(ByVal Point As PISDK.PIPoint) As Integer
    #End If
        DBMPointIndex=Array.FindIndex(DBMPoints,Function(FindDBMPoint)FindDBMPoint.Point Is Point)
        If DBMPointIndex=-1 Then
            ReDim Preserve DBMPoints(DBMPoints.Length)
            DBMPointIndex=DBMPoints.Length-1
            DBMPoints(DBMPointIndex)=New DBMPoint(Point)
        End If
        Return DBMPointIndex
    End Function

    #If OfflineUnitTests Then
    Public Function Calculate(ByVal InputPoint As String,ByVal CorrelationPoint As String,ByVal Timestamp As DateTime,Optional ByVal SubstractInputPointFromCorrelationPoint As Boolean=False) As Double
    #Else
    Public Function Calculate(ByVal InputPoint As PISDK.PIPoint,ByVal CorrelationPoint As PISDK.PIPoint,ByVal Timestamp As DateTime,Optional ByVal SubstractInputPointFromCorrelationPoint As Boolean=False) As Double
    #End If
        Dim InputDBMPointIndex,CorrelationDBMPointIndex As Integer
        Dim AbsErrorStats,RelErrorStats As New Statistics
        InputDBMPointIndex=DBMPointIndex(InputPoint)
        DBMPoints(InputDBMPointIndex).Calculate(Timestamp,True,Not IsNothing(CorrelationPoint))
        Calculate=DBMPoints(InputDBMPointIndex).Factor
        If Calculate<>0 And Not IsNothing(CorrelationPoint) Then ' If an event is found and a correlation point is available
            CorrelationDBMPointIndex=DBMPointIndex(CorrelationPoint)
            If SubstractInputPointFromCorrelationPoint Then
                DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointIndex)) ' Pattern of correlation point contains input point
            Else
                DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True)
            End If
            AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).AbsoluteError,DBMPoints(InputDBMPointIndex).AbsoluteError)
            RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).RelativeError,DBMPoints(InputDBMPointIndex).RelativeError)
            If RelErrorStats.ModifiedCorrelation>DBMConstants.CorrelationThreshold Then ' Suppress event due to correlation of relative error
                Calculate=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-DBMConstants.CorrelationThreshold Then ' Suppress event due to anticorrelation of absolute error (unmeasured supply)
                Calculate=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return Math.Round(Calculate,3)
    End Function

End Class
