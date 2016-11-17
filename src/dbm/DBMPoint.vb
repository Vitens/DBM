Option Explicit
Option Strict

Public Class DBMPoint

    Private Structure CachedValue

        Public Timestamp As DateTime
        Public Value As Double

        Public Sub New(ByVal Timestamp As DateTime,ByVal Value As Double)
            Me.Timestamp=Timestamp
            Me.Value=Value
        End Sub

    End Structure

    #If OfflineUnitTests Then
    Public Point As String
    #Else
    Public Point As PISDK.PIPoint
    #End If
    Private CachedValues() As CachedValue
    Public Factor,AbsoluteError(),RelativeError() As Double

    #If OfflineUnitTests Then
    Public Sub New(ByVal Point As String)
    #Else
    Public Sub New(ByVal Point As PISDK.PIPoint)
    #End If
        Me.Point=Point
        ReDim Me.CachedValues(CInt((DBMConstants.EMAPreviousPeriods+1+DBMConstants.CorrelationPreviousPeriods+1+24*(3600/DBMConstants.CalculationInterval))*(DBMConstants.ComparePatterns+1)-1))
        ReDim Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods)
        ReDim Me.RelativeError(DBMConstants.CorrelationPreviousPeriods)
    End Sub

    Public Function ArrRemoveFirstValue(ByVal Data() As Double) As Double()
        Array.Reverse(Data) ' ABCDE -> EDCBA
        Array.Reverse(Data,0,Data.Length-1) ' EDCBA -> BCDEA
        Return Data
    End Function

    Private Function Value(ByVal Timestamp As DateTime) As Double
        Dim i As Integer
        i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp)
        If i>=0 Then ' Found value in cache
            Array.Reverse(Me.CachedValues,0,i)
            Array.Reverse(Me.CachedValues,0,i+1) ' Move value to beginning of cache
        Else
            Array.Reverse(Me.CachedValues,0,Me.CachedValues.Length-1) ' Remove last value from cache, ABCDE -> DCBAE
            Array.Reverse(Me.CachedValues) ' DCBAE -> EABCD
            Try
                #If OfflineUnitTests Then
                Me.CachedValues(0)=New CachedValue(Timestamp,DBM.UnitTestData(0))
                DBM.UnitTestData=DBM.UnitTestData.Skip(1).ToArray
                #Else
                Me.CachedValues(0)=New CachedValue(Timestamp,CDbl(Me.Point.Data.Summary(Timestamp,DateAdd("s",DBMConstants.CalculationInterval,Timestamp),PISDK.ArchiveSummaryTypeConstants.astAverage,PISDK.CalculationBasisConstants.cbTimeWeighted).Value))
                #End If
            Catch
                Me.CachedValues(0)=New CachedValue(Timestamp,Double.NaN)
            End Try
        End If
        Return Me.CachedValues(0).Value
    End Function

    Public Sub Calculate(ByVal Timestamp As DateTime,ByVal IsInputDBMPoint As Boolean,ByVal HasCorrelationDBMPoint As Boolean,Optional ByRef SubstractDBMPoint As DBMPoint=Nothing)
        Dim CorrelationCounter,EMACounter,PatternCounter As Integer
        Dim Pattern(DBMConstants.ComparePatterns),CurrValueEMA(DBMConstants.EMAPreviousPeriods),PredValueEMA(DBMConstants.EMAPreviousPeriods),LowContrLimitEMA(DBMConstants.EMAPreviousPeriods),UppContrLimitEMA(DBMConstants.EMAPreviousPeriods) As Double
        Dim CurrValue,PredValue,LowContrLimit,UppContrLimit As Double
        Dim DBMStatistics As New DBMStatistics
        Dim DBMMath As New DBMMath
        Me.Factor=0 ' No event
        For CorrelationCounter=0 To DBMConstants.CorrelationPreviousPeriods
            If CorrelationCounter=0 Or (IsInputDBMPoint And Me.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                For EMACounter=DBMConstants.EMAPreviousPeriods To 0 Step -1
                    If CorrelationCounter=0 Or (CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods) Then
                        If CorrelationCounter>0 And EMACounter=DBMConstants.EMAPreviousPeriods Then
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
                        DBMStatistics.Calculate(DBMMath.RemoveOutliers(Pattern.Take(Pattern.Length-1).ToArray),Nothing)
                        CurrValueEMA(EMACounter)=Pattern(DBMConstants.ComparePatterns)
                        PredValueEMA(EMACounter)=DBMConstants.ComparePatterns*DBMStatistics.Slope+DBMStatistics.Intercept
                        LowContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)-DBMMath.ControlLimitRejectionCriterion(DBMStatistics.Count)*DBMStatistics.StDevSLinReg
                        UppContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)+DBMMath.ControlLimitRejectionCriterion(DBMStatistics.Count)*DBMStatistics.StDevSLinReg
                    End If
                Next EMACounter
                CurrValue=DBMMath.CalculateExpMovingAvg(CurrValueEMA)
                PredValue=DBMMath.CalculateExpMovingAvg(PredValueEMA)
                LowContrLimit=DBMMath.CalculateExpMovingAvg(LowContrLimitEMA)
                UppContrLimit=DBMMath.CalculateExpMovingAvg(UppContrLimitEMA)
                Me.AbsoluteError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=PredValue-CurrValue
                Me.RelativeError(DBMConstants.CorrelationPreviousPeriods-CorrelationCounter)=PredValue/CurrValue-1
                If CorrelationCounter=0 Then
                    If CurrValue<LowContrLimit Then ' Lower control limit exceeded
                        Me.Factor=(PredValue-CurrValue)/(LowContrLimit-PredValue)
                    End If
                    If CurrValue>UppContrLimit Then ' Upper control limit exceeded
                        Me.Factor=(CurrValue-PredValue)/(UppContrLimit-PredValue)
                    End If
                End If
            End If
        Next CorrelationCounter
    End Sub

End Class
