Option Explicit
Option Strict

Public Class DBMPoint

    Public Const CalculationInterval As Integer =           300 ' seconds; 300 = 5 minutes
    Public Const ComparePatterns As Integer =               12 ' weeks
    Public Const EMAPreviousPeriods As Integer =            6 ' previous periods; 6 = 35 minutes, current value inclusive
    Public Const CorrelationPreviousPeriods As Integer =    CInt(2*3600/CalculationInterval-1) ' previous periods; 23 = 2 hours, current value inclusive

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
        ReDim Me.CachedValues(CInt((EMAPreviousPeriods+1+CorrelationPreviousPeriods+1+24*(3600/CalculationInterval))*(ComparePatterns+1)-1))
        ReDim Me.AbsoluteError(CorrelationPreviousPeriods)
        ReDim Me.RelativeError(CorrelationPreviousPeriods)
    End Sub

    Private Function Value(ByVal Timestamp As DateTime) As Double
        Dim i As Integer
        i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp)
        If i>=0 Then ' Found value in cache
            Array.Reverse(Me.CachedValues,0,i)
            Array.Reverse(Me.CachedValues,0,i+1) ' Move value to beginning of cache
        Else
            Array.Reverse(Me.CachedValues,0,Me.CachedValues.Length-1)
            Array.Reverse(Me.CachedValues) ' Remove last value from cache
            Try
                #If OfflineUnitTests Then
                Me.CachedValues(0)=New CachedValue(Timestamp,DBM.UnitTestData(0))
                DBM.UnitTestData=DBM.UnitTestData.Skip(1).ToArray
                #Else
                Me.CachedValues(0)=New CachedValue(Timestamp,CDbl(Me.Point.Data.Summary(Timestamp,DateAdd("s",CalculationInterval,Timestamp),PISDK.ArchiveSummaryTypeConstants.astAverage,PISDK.CalculationBasisConstants.cbTimeWeighted).Value))
                #End If
            Catch
                Me.CachedValues(0)=New CachedValue(Timestamp,Double.NaN)
            End Try
        End If
        Return Me.CachedValues(0).Value
    End Function

    Public Sub Calculate(ByVal Timestamp As DateTime,ByVal IsInputDBMPoint As Boolean,ByVal HasCorrelationDBMPoint As Boolean,Optional ByRef SubstractDBMPoint As DBMPoint=Nothing)
        Dim CorrelationCounter,EMACounter,PatternCounter As Integer
        Dim Pattern(ComparePatterns),CurrValueEMA(EMAPreviousPeriods),PredValueEMA(EMAPreviousPeriods),LowContrLimitEMA(EMAPreviousPeriods),UppContrLimitEMA(EMAPreviousPeriods) As Double
        Dim CurrValue,PredValue,LowContrLimit,UppContrLimit As Double
        Dim Stats As New Statistics
        Me.Factor=0 ' No event
        For CorrelationCounter=0 To CorrelationPreviousPeriods
            If CorrelationCounter=0 Or (IsInputDBMPoint And Me.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                For EMACounter=EMAPreviousPeriods To 0 Step -1
                    For PatternCounter=ComparePatterns To 0 Step -1
                        Pattern(ComparePatterns-PatternCounter)=Me.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*CalculationInterval,Timestamp)))
                        If Not IsNothing(SubstractDBMPoint) Then
                            Pattern(ComparePatterns-PatternCounter)-=SubstractDBMPoint.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*CalculationInterval,Timestamp)))
                        End If
                    Next PatternCounter
                    Stats.Calculate(Stats.RemoveOutliers(Pattern.Take(Pattern.Length-1).ToArray),Nothing)
                    CurrValueEMA(EMACounter)=Pattern(ComparePatterns)
                    PredValueEMA(EMACounter)=ComparePatterns*Stats.Slope+Stats.Intercept
                    LowContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)-Stats.ControlLimitRejectionCriterion(Stats.Count)*Stats.StDevSLinReg
                    UppContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)+Stats.ControlLimitRejectionCriterion(Stats.Count)*Stats.StDevSLinReg
                Next EMACounter
                CurrValue=Stats.CalculateExpMovingAvg(CurrValueEMA)
                PredValue=Stats.CalculateExpMovingAvg(PredValueEMA)
                LowContrLimit=Stats.CalculateExpMovingAvg(LowContrLimitEMA)
                UppContrLimit=Stats.CalculateExpMovingAvg(UppContrLimitEMA)
                Me.AbsoluteError(CorrelationPreviousPeriods-CorrelationCounter)=PredValue-CurrValue
                Me.RelativeError(CorrelationPreviousPeriods-CorrelationCounter)=PredValue/CurrValue-1
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
