Option Explicit
Option Strict

Public Class DBMPoint

    Public DBMPointDriver As New DBMPointDriver
    Private CachedValues() As DBMCachedValue
    Public AbsoluteError(),RelativeError() As Double
    Private DBMFunctions As New DBMFunctions

    Public Sub New(ByVal DBMPointDriver As DBMPointDriver)
        Dim i As Integer
        Me.DBMPointDriver=DBMPointDriver
        ReDim Me.CachedValues(CInt((DBMConstants.EMAPreviousPeriods+1+DBMConstants.CorrelationPreviousPeriods+1+24*(3600/DBMConstants.CalculationInterval))*(DBMConstants.ComparePatterns+1)-1))
        For i=0 to Me.CachedValues.Length-1 ' Initialise cache
            Me.CachedValues(i)=New DBMCachedValue(Nothing,Nothing)
        Next i
        ReDim Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods)
        ReDim Me.RelativeError(DBMConstants.CorrelationPreviousPeriods)
    End Sub

    Private Function Value(ByVal Timestamp As DateTime) As Double
        Dim i As Integer
        i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp) ' Find timestamp in cache
        If i>=0 Then ' Found timestamp in cache
            DBMFunctions.ArrayMoveItemToFront(Me.CachedValues,i) ' Move to first item in cache
        Else
            DBMFunctions.ArrayRotateRight(Me.CachedValues) ' Remove last item from cache
            Try
                Me.CachedValues(0)=New DBMCachedValue(Timestamp,Me.DBMPointDriver.GetData(Timestamp,DateAdd("s",DBMConstants.CalculationInterval,Timestamp))) ' Get data using driver
            Catch
                Me.CachedValues(0)=New DBMCachedValue(Timestamp,Double.NaN) ' Error, return Not a Number
            End Try
        End If
        Return Me.CachedValues(0).Value
    End Function

    Public Function Calculate(ByVal Timestamp As DateTime,ByVal IsInputDBMPoint As Boolean,ByVal HasCorrelationDBMPoint As Boolean,Optional ByRef SubstractDBMPoint As DBMPoint=Nothing) As DBMResult
        Dim CorrelationCounter,EMACounter,PatternCounter As Integer
        Dim Pattern(DBMConstants.ComparePatterns),CurrValueEMA(DBMConstants.EMAPreviousPeriods),PredValueEMA(DBMConstants.EMAPreviousPeriods),LowContrLimitEMA(DBMConstants.EMAPreviousPeriods),UppContrLimitEMA(DBMConstants.EMAPreviousPeriods) As Double
        Dim DBMStatistics As New DBMStatistics
        Dim DBMMath As New DBMMath
        Calculate.Factor=0 ' No event
        For CorrelationCounter=0 To DBMConstants.CorrelationPreviousPeriods
            If CorrelationCounter=0 Or (IsInputDBMPoint And Calculate.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                For EMACounter=DBMConstants.EMAPreviousPeriods To 0 Step -1
                    If CorrelationCounter=0 Or (CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods) Then
                        If CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods Then ' Reuse calculation results when moving back for correlation calculation
                            CurrValueEMA=DBMFunctions.ArrayRotateLeft(CurrValueEMA)
                            PredValueEMA=DBMFunctions.ArrayRotateLeft(PredValueEMA)
                            LowContrLimitEMA=DBMFunctions.ArrayRotateLeft(LowContrLimitEMA)
                            UppContrLimitEMA=DBMFunctions.ArrayRotateLeft(UppContrLimitEMA)
                        End If
                        For PatternCounter=DBMConstants.ComparePatterns To 0 Step -1
                            Pattern(DBMConstants.ComparePatterns-PatternCounter)=Me.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*DBMConstants.CalculationInterval,Timestamp)))
                            If Not IsNothing(SubstractDBMPoint) Then
                                Pattern(DBMConstants.ComparePatterns-PatternCounter)-=SubstractDBMPoint.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*DBMConstants.CalculationInterval,Timestamp)))
                            End If
                        Next PatternCounter
                        DBMStatistics.Calculate(DBMMath.RemoveOutliers(Pattern.Take(Pattern.Length-1).ToArray),Nothing) ' Calculate statistics for data after removing outliers
                        CurrValueEMA(EMACounter)=Pattern(DBMConstants.ComparePatterns)
                        PredValueEMA(EMACounter)=DBMConstants.ComparePatterns*DBMStatistics.Slope+DBMStatistics.Intercept ' Extrapolate linear regression
                        LowContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)-DBMMath.ControlLimitRejectionCriterion(DBMStatistics.Count)*DBMStatistics.StDevSLinReg
                        UppContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)+DBMMath.ControlLimitRejectionCriterion(DBMStatistics.Count)*DBMStatistics.StDevSLinReg
                    End If
                Next EMACounter
                Calculate.CurrValue=DBMMath.CalculateExpMovingAvg(CurrValueEMA)
                Calculate.PredValue=DBMMath.CalculateExpMovingAvg(PredValueEMA)
                Calculate.LowContrLimit=DBMMath.CalculateExpMovingAvg(LowContrLimitEMA)
                Calculate.UppContrLimit=DBMMath.CalculateExpMovingAvg(UppContrLimitEMA)
                Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=Calculate.PredValue-Calculate.CurrValue ' Absolute error compared to prediction
                Me.RelativeError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=Calculate.PredValue/Calculate.CurrValue-1 ' Relative error compared to prediction
                If CorrelationCounter=0 Then
                    If Calculate.CurrValue<Calculate.LowContrLimit Then ' Lower control limit exceeded
                        Calculate.Factor=(Calculate.PredValue-Calculate.CurrValue)/(Calculate.LowContrLimit-Calculate.PredValue)
                    End If
                    If Calculate.CurrValue>Calculate.UppContrLimit Then ' Upper control limit exceeded
                        Calculate.Factor=(Calculate.CurrValue-Calculate.PredValue)/(Calculate.UppContrLimit-Calculate.PredValue)
                    End If
                End If
            End If
        Next CorrelationCounter
        Return Calculate
    End Function

End Class
