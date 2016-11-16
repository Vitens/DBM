Option Explicit
Option Strict

Public Class DBM

    Public Const CalculationInterval As Integer =           300 ' seconds; 300 = 5 minutes
    Public Const ComparePatterns As Integer =               12 ' weeks
    Public Const EMAPreviousPeriods As Integer =            6 ' previous periods; 6 = 35 minutes, current value inclusive
    Public Const CorrelationPreviousPeriods As Integer =    CInt(2*3600/CalculationInterval-1) ' previous periods; 23 = 2 hours, current value inclusive
    Public Const CorrelationThreshold As Double =           0.83666 ' absolute correlation lower limit for detecting (anti)correlation

    #If OfflineUnitTests Then
    Dim Shared UnitTestData() As Double
    #End If

    Public Structure DBMPoint

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
        Private CachedValues() As DBMPoint.CachedValue
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
                    Me.CachedValues(0)=New CachedValue(Timestamp,UnitTestData(0))
                    UnitTestData=UnitTestData.Skip(1).ToArray
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
            Dim Stats As Statistics
            Me.Factor=0 ' No event
            For CorrelationCounter=0 To CorrelationPreviousPeriods
                If CorrelationCounter=0 Or (IsInputDBMPoint And Me.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                    For EMACounter=EMAPreviousPeriods To 0 Step -1
                        For PatternCounter=ComparePatterns To 0 Step -1
                            Pattern(ComparePatterns-PatternCounter)=Me.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*CalculationInterval,Timestamp)))
                            If Not IsNothing(SubstractDBMPoint.Point) Then
                                Pattern(ComparePatterns-PatternCounter)-=SubstractDBMPoint.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*CalculationInterval,Timestamp)))
                            End If
                        Next PatternCounter
                        Stats.Calculate(Stats.RemoveOutliers(Pattern.Take(Pattern.Length-1).ToArray),Nothing)
                        CurrValueEMA(EMAPreviousPeriods-EMACounter)=Pattern(ComparePatterns)
                        PredValueEMA(EMAPreviousPeriods-EMACounter)=ComparePatterns*Stats.Slope+Stats.Intercept
                        LowContrLimitEMA(EMAPreviousPeriods-EMACounter)=PredValueEMA(EMAPreviousPeriods-EMACounter)-Stats.ControlLimitRejectionCriterion(Stats.n)*Stats.StDevSLinReg
                        UppContrLimitEMA(EMAPreviousPeriods-EMACounter)=PredValueEMA(EMAPreviousPeriods-EMACounter)+Stats.ControlLimitRejectionCriterion(Stats.n)*Stats.StDevSLinReg
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

    End Structure

    Private Structure Statistics

        Public Slope,Intercept,StDevSLinReg,ModifiedCorrelation As Double
        Public n As Integer

        Private Function MeanAbsDevScaleFactor As Double ' Scale factor k
            Return 1.253314137316 ' SQRT(PI()/2)
        End Function

        Private Function MedianAbsDevScaleFactor(ByVal n As Integer) As Double ' Scale factor k
            Select Case n
                Case <=3
                    MedianAbsDevScaleFactor=1.224744871392 ' n<30 Student's t-distribution: 1/T.INV(75%,n-1)
                Case 4
                    MedianAbsDevScaleFactor=1.307373551629
                Case 5
                    MedianAbsDevScaleFactor=1.350079568894
                Case 6
                    MedianAbsDevScaleFactor=1.376108579000
                Case 7
                    MedianAbsDevScaleFactor=1.393615186740
                Case 8
                    MedianAbsDevScaleFactor=1.406189357483
                Case 9
                    MedianAbsDevScaleFactor=1.415655367895
                Case 10
                    MedianAbsDevScaleFactor=1.423037547092
                Case 11
                    MedianAbsDevScaleFactor=1.428955079918
                Case 12
                    MedianAbsDevScaleFactor=1.433804142754
                Case 13
                    MedianAbsDevScaleFactor=1.437849945108
                Case 14
                    MedianAbsDevScaleFactor=1.441276685628
                Case 15
                    MedianAbsDevScaleFactor=1.444216273612
                Case 16
                    MedianAbsDevScaleFactor=1.446765645460
                Case 17
                    MedianAbsDevScaleFactor=1.448997629771
                Case 18
                    MedianAbsDevScaleFactor=1.450968000282
                Case 19
                    MedianAbsDevScaleFactor=1.452720190409
                Case 20
                    MedianAbsDevScaleFactor=1.454288526282
                Case 21
                    MedianAbsDevScaleFactor=1.455700494239
                Case 22
                    MedianAbsDevScaleFactor=1.456978363159
                Case 23
                    MedianAbsDevScaleFactor=1.458140366059
                Case 24
                    MedianAbsDevScaleFactor=1.459201574556
                Case 25
                    MedianAbsDevScaleFactor=1.460174555449
                Case 26
                    MedianAbsDevScaleFactor=1.461069870224
                Case 27
                    MedianAbsDevScaleFactor=1.461896459630
                Case 28
                    MedianAbsDevScaleFactor=1.462661943036
                Case 29
                    MedianAbsDevScaleFactor=1.463372853805
                Case >=30
                    MedianAbsDevScaleFactor=1.482602218506 ' n>=30 Standard normal distribution: 1/NORM.S.INV(75%)
            End Select
            Return MedianAbsDevScaleFactor
        End Function

        Public Function ControlLimitRejectionCriterion(ByVal n As Integer) As Double
            Select Case n
                Case <=3
                    ControlLimitRejectionCriterion=9.924843200918 ' n<30 Student's t-distribution: T.INV.2T(1%,n-1) (P=99%)
                Case 4
                    ControlLimitRejectionCriterion=5.840909309733
                Case 5
                    ControlLimitRejectionCriterion=4.604094871350
                Case 6
                    ControlLimitRejectionCriterion=4.032142983555
                Case 7
                    ControlLimitRejectionCriterion=3.707428021325
                Case 8
                    ControlLimitRejectionCriterion=3.499483297350
                Case 9
                    ControlLimitRejectionCriterion=3.355387331333
                Case 10
                    ControlLimitRejectionCriterion=3.249835541592
                Case 11
                    ControlLimitRejectionCriterion=3.169272672617
                Case 12
                    ControlLimitRejectionCriterion=3.105806515539
                Case 13
                    ControlLimitRejectionCriterion=3.054539589393
                Case 14
                    ControlLimitRejectionCriterion=3.012275838717
                Case 15
                    ControlLimitRejectionCriterion=2.976842734371
                Case 16
                    ControlLimitRejectionCriterion=2.946712883475
                Case 17
                    ControlLimitRejectionCriterion=2.920781622425
                Case 18
                    ControlLimitRejectionCriterion=2.898230519677
                Case 19
                    ControlLimitRejectionCriterion=2.878440472739
                Case 20
                    ControlLimitRejectionCriterion=2.860934606465
                Case 21
                    ControlLimitRejectionCriterion=2.845339709786
                Case 22
                    ControlLimitRejectionCriterion=2.831359558023
                Case 23
                    ControlLimitRejectionCriterion=2.818756060600
                Case 24
                    ControlLimitRejectionCriterion=2.807335683770
                Case 25
                    ControlLimitRejectionCriterion=2.796939504774
                Case 26
                    ControlLimitRejectionCriterion=2.787435813677
                Case 27
                    ControlLimitRejectionCriterion=2.778714533330
                Case 28
                    ControlLimitRejectionCriterion=2.770682957122
                Case 29
                    ControlLimitRejectionCriterion=2.763262455461
                Case >=30
                    ControlLimitRejectionCriterion=2.575829303549 ' n>=30 Standard normal distribution: NORM.S.INV(1-1%/2) (P=99%)
            End Select
            Return ControlLimitRejectionCriterion
        End Function

        Private Function CalculateMean(ByVal Data() As Double) As Double
            For Each Value As Double In Data
                CalculateMean+=Value/Data.Length
            Next
            Return CalculateMean
        End Function

        Private Function CalculateMedian(ByVal Data() As Double) As Double
            Array.Sort(Data)
            If Data.Length Mod 2=0 Then
                CalculateMedian=(Data(Data.Length\2)+Data(Data.Length\2-1))/2
            Else
                CalculateMedian=Data(Data.Length\2)
            End If
            Return CalculateMedian
        End Function

        Private Function CalculateMeanAbsDev(ByVal Mean As Double,ByVal Data() As Double) As Double
            Dim i As Integer
            For i=0 to Data.Length-1
                Data(i)=Math.Abs(Data(i)-Mean)
            Next i
            CalculateMeanAbsDev=CalculateMean(Data)
            Return CalculateMeanAbsDev
        End Function

        Private Function CalculateMedianAbsDev(ByVal Median As Double,ByVal Data() As Double) As Double
            Dim i As Integer
            For i=0 to Data.Length-1
                Data(i)=Math.Abs(Data(i)-Median)
            Next i
            CalculateMedianAbsDev=CalculateMedian(Data)
            Return CalculateMedianAbsDev
        End Function

        Public Function RemoveOutliers(ByVal Data() As Double) As Double()
            Dim Mean,Median,MeanAbsDev,MedianAbsDev As Double
            Dim i As Integer
            Mean=CalculateMean(Data.ToArray)
            Median=CalculateMedian(Data.ToArray)
            MeanAbsDev=CalculateMeanAbsDev(Mean,Data.ToArray)
            MedianAbsDev=CalculateMedianAbsDev(Median,Data.ToArray)
            For i=0 to Data.Length-1
                If MedianAbsDev=0 Then ' Use Mean Absolute Deviation instead of Median Absolute Deviation to detect outliers
                    If Math.Abs(Data(i)-Mean)>MeanAbsDev*MeanAbsDevScaleFactor*ControlLimitRejectionCriterion(ComparePatterns) Then ' If value is an outlier
                        Data(i)=Double.NaN ' Exclude outlier
                    End If
                Else ' Use Median Absolute Deviation to detect outliers
                    If Math.Abs(Data(i)-Median)>MedianAbsDev*MedianAbsDevScaleFactor(ComparePatterns)*ControlLimitRejectionCriterion(ComparePatterns) Then ' If value is an outlier
                        Data(i)=Double.NaN ' Exclude outlier
                    End If
                End If
            Next i
            Return Data
        End Function

        Public Function CalculateExpMovingAvg(ByVal Data() As Double) As Double ' Filter high frequency variation
            Dim Weight,TotalWeight As Double
            Weight=1
            TotalWeight=0
            For Each Value As Double In Data
                CalculateExpMovingAvg+=Value*Weight
                TotalWeight+=Weight
                Weight/=1-2/((Data.Length)+1)
            Next
            CalculateExpMovingAvg/=TotalWeight
            Return CalculateExpMovingAvg
        End Function

        Public Sub Calculate(ByVal DataY() As Double,Optional ByVal DataX() As Double=Nothing)
            Dim SumX,SumXX,SumY,SumYY,SumXY As Double
            Dim i As Integer
            Me.n=0
            For i=0 To DataY.Length-1
                If Not Double.IsNaN(DataY(i)) Then
                    If DataX Is Nothing Then
                        SumX+=i
                        SumXX+=i^2
                        SumXY+=i*DataY(i)
                    Else
                        SumX+=DataX(i)
                        SumXX+=DataX(i)^2
                        SumXY+=DataX(i)*DataY(i)
                    End If
                    SumY+=DataY(i)
                    SumYY+=DataY(i)^2
                    Me.n+=1
                End If
            Next i
            Me.Slope=(Me.n*SumXY-SumX*SumY)/(Me.n*SumXX-SumX^2)
            Me.Intercept=(SumX*SumXY-SumY*SumXX)/(SumX^2-Me.n*SumXX)
            Me.StDevSLinReg=0
            For i=0 to DataY.Length-1
                If Not Double.IsNaN(DataY(i)) Then
                    Me.StDevSLinReg+=(DataY(i)-i*Me.Slope-Me.Intercept)^2
                End If
            Next i
            Me.StDevSLinReg=Math.Sqrt(Me.StDevSLinReg/(Me.n-2)) ' n-2 is used because two parameters (slope and intercept) were estimated in order to estimate the sum of squares
            Me.ModifiedCorrelation=SumXY/Math.Sqrt(SumXX)/Math.Sqrt(SumYY) ' Average is not removed, as expected average is zero
        End Sub

    End Structure

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
        Dim AbsErrorStats,RelErrorStats As Statistics
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
            If RelErrorStats.ModifiedCorrelation>CorrelationThreshold Then ' Suppress event due to correlation of relative error
                Calculate=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-CorrelationThreshold Then ' Suppress event due to anticorrelation of absolute error (unmeasured supply)
                Calculate=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return Math.Round(Calculate,3)
    End Function

End Class
