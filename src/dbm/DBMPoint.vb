Option Explicit
Option Strict

Public Class DBMPoint

    #If OfflineUnitTests Then
    Public Point As String
    #Else
    Public Point As PISDK.PIPoint
    #End If
    Private CachedValues() As DBMCachedValue
    Public Factor,CurrValue,PredValue,LowContrLimit,UppContrLimit,AbsoluteError(),RelativeError() As Double

    #If OfflineUnitTests Then
    Public Sub New(ByVal Point As String)
    #Else
    Public Sub New(ByVal Point As PISDK.PIPoint)
    #End If
        Dim i As Integer
        Me.Point=Point
        ReDim Me.CachedValues(CInt((DBMConstants.EMAPreviousPeriods+1+DBMConstants.CorrelationPreviousPeriods+1+24*(3600/DBMConstants.CalculationInterval))*(DBMConstants.ComparePatterns+1)-1))
        ReDim Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods)
        ReDim Me.RelativeError(DBMConstants.CorrelationPreviousPeriods)
        For i=0 to Me.CachedValues.Length-1 ' Initialise cache
            Me.CachedValues(i)=New DBMCachedValue(Nothing,Nothing)
        Next i
    End Sub

    Public Function ArrRemoveFirstValue(ByVal Data() As Double) As Double()
        Array.Reverse(Data) ' ABCDE -> EDCBA
        Array.Reverse(Data,0,Data.Length-1) ' EDCBA -> BCDEA
        Return Data
    End Function

    Private Function Value(ByVal Timestamp As DateTime) As Double
        Dim i As Integer
        i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp) ' Find timestamp in cache
        If i>=0 Then ' Found timestamp in cache
            Array.Reverse(Me.CachedValues,0,i)
            Array.Reverse(Me.CachedValues,0,i+1) ' Move item to beginning of cache
        Else
            Array.Reverse(Me.CachedValues,0,Me.CachedValues.Length-1) ' Remove last item from cache, ABCDE -> DCBAE
            Array.Reverse(Me.CachedValues) ' DCBAE -> EABCD
            Try
                #If OfflineUnitTests Then
                Me.CachedValues(0)=New DBMCachedValue(Timestamp,DBM.UnitTestData(0))
                DBM.UnitTestData=DBM.UnitTestData.Skip(1).ToArray
                #Else
                Me.CachedValues(0)=New DBMCachedValue(Timestamp,CDbl(Me.Point.Data.Summary(Timestamp,DateAdd("s",DBMConstants.CalculationInterval,Timestamp),PISDK.ArchiveSummaryTypeConstants.astAverage,PISDK.CalculationBasisConstants.cbTimeWeighted).Value)) ' Get data from OSIsoft PI
                #End If
            Catch
                Me.CachedValues(0)=New DBMCachedValue(Timestamp,Double.NaN) ' Error, return Not a Number
            End Try
        End If
        Return Me.CachedValues(0).Value
    End Function

    Public Sub Calculate(ByVal Timestamp As DateTime,ByVal IsInputDBMPoint As Boolean,ByVal HasCorrelationDBMPoint As Boolean,Optional ByRef SubstractDBMPoint As DBMPoint=Nothing)
        Dim CorrelationCounter,EMACounter,PatternCounter As Integer
        Dim Pattern(DBMConstants.ComparePatterns),CurrValueEMA(DBMConstants.EMAPreviousPeriods),PredValueEMA(DBMConstants.EMAPreviousPeriods),LowContrLimitEMA(DBMConstants.EMAPreviousPeriods),UppContrLimitEMA(DBMConstants.EMAPreviousPeriods) As Double
        Dim DBMStatistics As New DBMStatistics
        Dim DBMMath As New DBMMath
        Me.Factor=0 ' No event
        For CorrelationCounter=0 To DBMConstants.CorrelationPreviousPeriods
            If CorrelationCounter=0 Or (IsInputDBMPoint And Me.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                For EMACounter=DBMConstants.EMAPreviousPeriods To 0 Step -1
                    If CorrelationCounter=0 Or (CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods) Then
                        If CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods Then ' Reuse calculation results when moving back for correlation calculation
                            CurrValueEMA=ArrRemoveFirstValue(CurrValueEMA)
                            PredValueEMA=ArrRemoveFirstValue(PredValueEMA)
                            LowContrLimitEMA=ArrRemoveFirstValue(LowContrLimitEMA)
                            UppContrLimitEMA=ArrRemoveFirstValue(UppContrLimitEMA)
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
                Me.CurrValue=DBMMath.CalculateExpMovingAvg(CurrValueEMA)
                Me.PredValue=DBMMath.CalculateExpMovingAvg(PredValueEMA)
                Me.LowContrLimit=DBMMath.CalculateExpMovingAvg(LowContrLimitEMA)
                Me.UppContrLimit=DBMMath.CalculateExpMovingAvg(UppContrLimitEMA)
                Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=Me.PredValue-Me.CurrValue ' Absolute error compared to prediction
                Me.RelativeError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=Me.PredValue/Me.CurrValue-1 ' Relative error compared to prediction
                If CorrelationCounter=0 Then
                    If Me.CurrValue<Me.LowContrLimit Then ' Lower control limit exceeded
                        Me.Factor=(Me.PredValue-Me.CurrValue)/(Me.LowContrLimit-Me.PredValue)
                    End If
                    If Me.CurrValue>Me.UppContrLimit Then ' Upper control limit exceeded
                        Me.Factor=(Me.CurrValue-Me.PredValue)/(Me.UppContrLimit-Me.PredValue)
                    End If
                End If
            End If
        Next CorrelationCounter
    End Sub

End Class
