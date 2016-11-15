Option Explicit
Option Strict

Public Class DBM

    Public Const DBMVersion As String =                     "DBM v1.16 161102 J.H. Fitié, Vitens N.V."
    Public Const CalculationInterval As Integer =           300 ' seconds; 300 = 5 minutes
    Public Const ComparePatterns As Integer =               12 ' weeks
    Public Const EMAPreviousPeriods As Integer =            6 ' previous periods; 6 = 35 minutes, current value inclusive
    Public Const CorrelationPreviousPeriods As Integer =    CInt(2*3600/CalculationInterval-1) ' previous periods; 23 = 2 hours, current value inclusive
    Public Const CorrelationThreshold As Double =           0.83666 ' absolute correlation lower limit for detecting (anti)correlation

    #If OfflineUnitTests Then
    Dim Shared UnitTestData() As Double
    #End If

    Public Structure DBMPoint

        Public Structure CachedValue

            Public Timestamp As DateTime
            Public Value As Double

            Public Sub New(ByVal Timestamp As DateTime,ByVal Value As Double)
                Me.Timestamp=Timestamp
                Me.Value=Value
            End Sub

            Public Sub Invalidate
                Me.Timestamp=Nothing
                Me.Value=Nothing
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
            InvalidateCache
            ReDim Me.AbsoluteError(CorrelationPreviousPeriods)
            ReDim Me.RelativeError(CorrelationPreviousPeriods)
        End Sub

        Public Function Value(ByVal Timestamp As DateTime) As Double
            Dim i As Integer
            i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp)
            If i>=0 Then
                Array.Reverse(Me.CachedValues,0,i)
                Array.Reverse(Me.CachedValues,0,i+1)
            Else
                Array.Reverse(Me.CachedValues,0,Me.CachedValues.Length-1)
                Array.Reverse(Me.CachedValues)
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

        Public Sub InvalidateCache(Optional ByVal Timestamp As DateTime=Nothing)
            Dim i As Integer
            If Timestamp=DateTime.MinValue Then
                ReDim Me.CachedValues(CInt((EMAPreviousPeriods+1+CorrelationPreviousPeriods+1+24*(3600/CalculationInterval))*(ComparePatterns+1)-1))
            Else
                i=Array.FindIndex(Me.CachedValues,Function(FindCachedValue)FindCachedValue.Timestamp=Timestamp)
                If i>=0 Then
                    Me.CachedValues(i).Invalidate
                    Array.Reverse(Me.CachedValues,i,Me.CachedValues.Length-i)
                    Array.Reverse(Me.CachedValues,i,Me.CachedValues.Length-i-1)
                End If
            End If
        End Sub

        Private Function MedianADScaleFactor(ByVal n As Integer) As Double
            Return 1.224744871392+CDbl(IIf(n>3,1,0))*0.082628680237+CDbl(IIf(n>4,1,0))*0.042706017265+CDbl(IIf(n>5,1,0))*0.026029010106+CDbl(IIf(n>6,1,0))*0.01750660774+CDbl(IIf(n>7,1,0))*0.012574170743+CDbl(IIf(n>8,1,0))*0.009466010412+CDbl(IIf(n>9,1,0))*0.007382179197+CDbl(IIf(n>10,1,0))*0.005917532826+CDbl(IIf(n>11,1,0))*0.004849062836+CDbl(IIf(n>12,1,0))*0.004045802354+CDbl(IIf(n>13,1,0))*0.00342674052+CDbl(IIf(n>14,1,0))*0.002939587984+CDbl(IIf(n>15,1,0))*0.002549371848+CDbl(IIf(n>16,1,0))*0.002231984311+CDbl(IIf(n>17,1,0))*0.001970370511+CDbl(IIf(n>18,1,0))*0.001752190127+CDbl(IIf(n>19,1,0))*0.001568335873+CDbl(IIf(n>20,1,0))*0.001411967957+CDbl(IIf(n>21,1,0))*0.00127786892+CDbl(IIf(n>22,1,0))*0.0011620029+CDbl(IIf(n>23,1,0))*0.001061208497+CDbl(IIf(n>24,1,0))*0.000972980893+CDbl(IIf(n>25,1,0))*0.000895314775+CDbl(IIf(n>26,1,0))*0.000826589406+CDbl(IIf(n>27,1,0))*0.000765483406+CDbl(IIf(n>28,1,0))*0.000710910769+CDbl(IIf(n>29,1,0))*0.019229364701
        End Function

        Private Function ControlLimitRejectionCriterion(ByVal n As Integer) As Double
            Return 9.924843200918-CDbl(IIf(n>3,1,0))*4.083933891185-CDbl(IIf(n>4,1,0))*1.236814438383-CDbl(IIf(n>5,1,0))*0.571951887795-CDbl(IIf(n>6,1,0))*0.32471496223-CDbl(IIf(n>7,1,0))*0.207944723975-CDbl(IIf(n>8,1,0))*0.144095966017-CDbl(IIf(n>9,1,0))*0.105551789741-CDbl(IIf(n>10,1,0))*0.080562868975-CDbl(IIf(n>11,1,0))*0.063466157078-CDbl(IIf(n>12,1,0))*0.051266926146-CDbl(IIf(n>13,1,0))*0.042263750676-CDbl(IIf(n>14,1,0))*0.035433104346-CDbl(IIf(n>15,1,0))*0.030129850896-CDbl(IIf(n>16,1,0))*0.02593126105-CDbl(IIf(n>17,1,0))*0.022551102748-CDbl(IIf(n>18,1,0))*0.019790046938-CDbl(IIf(n>19,1,0))*0.017505866274-CDbl(IIf(n>20,1,0))*0.015594896679-CDbl(IIf(n>21,1,0))*0.013980151763-CDbl(IIf(n>22,1,0))*0.012603497423-CDbl(IIf(n>23,1,0))*0.01142037683-CDbl(IIf(n>24,1,0))*0.010396178996-CDbl(IIf(n>25,1,0))*0.009503691097-CDbl(IIf(n>26,1,0))*0.008721280347-CDbl(IIf(n>27,1,0))*0.008031576208-CDbl(IIf(n>28,1,0))*0.007420501661-CDbl(IIf(n>29,1,0))*0.187433151912
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

        Public Sub Calculate(ByVal Timestamp As DateTime,ByVal IsInputDBMPoint As Boolean,ByVal HasCorrelationDBMPoint As Boolean,Optional ByRef SubstractDBMPoint As DBMPoint=Nothing)
            Dim CorrelationCounter,EMACounter,PatternCounter,n As Integer
            Dim EMAWeight,EMATotalWeight,Median,Mean,MedianAD,MeanAD,VarS,StDevS,CurrEMA,PredEMA,UCLEMA,LCLEMA As Double
            Dim Pattern(ComparePatterns),Data(ComparePatterns-1) As Double
            Dim Stats As Statistics
            Me.Factor=0
            For CorrelationCounter=CorrelationPreviousPeriods To 0 Step -1
                If CorrelationCounter=CorrelationPreviousPeriods Or (IsInputDBMPoint And Me.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                    EMAWeight=1
                    EMATotalWeight=0
                    CurrEMA=0
                    PredEMA=0
                    UCLEMA=0
                    LCLEMA=0
                    For EMACounter=0 To EMAPreviousPeriods
                        Mean=0
                        For PatternCounter=0 To ComparePatterns
                            Pattern(PatternCounter)=Me.Value(DateAdd("d",-(ComparePatterns-PatternCounter)*7,DateAdd("s",-((EMAPreviousPeriods-EMACounter)+(CorrelationPreviousPeriods-CorrelationCounter))*CalculationInterval,Timestamp)))
                            If Not IsNothing(SubstractDBMPoint.Point) Then
                                Pattern(PatternCounter)-=SubstractDBMPoint.Value(DateAdd("d",-(ComparePatterns-PatternCounter)*7,DateAdd("s",-((EMAPreviousPeriods-EMACounter)+(CorrelationPreviousPeriods-CorrelationCounter))*CalculationInterval,Timestamp)))
                            End If
                            If PatternCounter<ComparePatterns Then
                                Mean+=Pattern(PatternCounter)/ComparePatterns
                                Data(PatternCounter)=Pattern(PatternCounter)
                            End If
                        Next PatternCounter
                        Median=CalculateMedian(Data)
                        MeanAD=0
                        For PatternCounter=0 To ComparePatterns-1
                            MeanAD+=Math.Abs(Pattern(PatternCounter)-Mean)/ComparePatterns
                            Data(PatternCounter)=Math.Abs(Pattern(PatternCounter)-Median)
                        Next PatternCounter
                        MedianAD=CalculateMedian(Data)
                        n=0
                        For PatternCounter=0 To ComparePatterns-1
                            If MedianAD=0 Then
                                If Math.Abs(Pattern(PatternCounter)-Mean)>MeanAD*Math.Sqrt(Math.PI/2)*ControlLimitRejectionCriterion(ComparePatterns) Then
                                    Pattern(PatternCounter)=Double.NaN
                                Else
                                    n+=1
                                End If
                            Else
                                If Math.Abs(Pattern(PatternCounter)-Median)>MedianAD*MedianADScaleFactor(ComparePatterns)*ControlLimitRejectionCriterion(ComparePatterns) Then
                                    Pattern(PatternCounter)=Double.NaN
                                Else
                                    n+=1
                                End If
                            End If
                        Next PatternCounter
                        Stats.Calculate(Pattern,Nothing,True)
                        VarS=0
                        StDevS=0
                        For PatternCounter=0 To ComparePatterns-1
                            If Not Double.IsNaN(Pattern(PatternCounter)) Then
                                VarS+=(Pattern(PatternCounter)-PatternCounter*Stats.Slope-Stats.Intercept)^2/(n-2)
                            End If
                        Next PatternCounter
                        If VarS<>0 Then
                            StDevS=Math.Sqrt(VarS)
                        End If
                        CurrEMA+=(Pattern(ComparePatterns))*EMAWeight
                        PredEMA+=(ComparePatterns*Stats.Slope+Stats.Intercept)*EMAWeight
                        UCLEMA+=(ComparePatterns*Stats.Slope+Stats.Intercept+ControlLimitRejectionCriterion(n)*StDevS)*EMAWeight
                        LCLEMA+=(ComparePatterns*Stats.Slope+Stats.Intercept-ControlLimitRejectionCriterion(n)*StDevS)*EMAWeight
                        EMATotalWeight+=EMAWeight
                        EMAWeight/=1-2/((EMAPreviousPeriods+1)+1)
                    Next EMACounter
                    CurrEMA/=EMATotalWeight
                    PredEMA/=EMATotalWeight
                    Me.AbsoluteError(CorrelationCounter)=PredEMA-CurrEMA
                    Me.RelativeError(CorrelationCounter)=PredEMA/CurrEMA-1
                    UCLEMA/=EMATotalWeight
                    LCLEMA/=EMATotalWeight
                    If CorrelationCounter=CorrelationPreviousPeriods Then
                        If CurrEMA<LCLEMA Then
                            Me.Factor=(PredEMA-CurrEMA)/(LCLEMA-PredEMA)
                        End If
                        If CurrEMA>UCLEMA Then
                            Me.Factor=(CurrEMA-PredEMA)/(UCLEMA-PredEMA)
                        End If
                    End If
                End If
            Next CorrelationCounter
        End Sub

    End Structure

    Private Structure Statistics

        Public Slope,Intercept,ModifiedCorrelation As Double

        Public Sub Calculate(ByVal DataY() As Double,Optional ByVal DataX() As Double=Nothing,Optional ByVal ExcludeLastValue As Boolean=False)
            Dim SumX,SumXX,SumY,SumYY,SumXY As Double
            Dim i,n As Integer
            For i=0 To CInt(IIf(ExcludeLastValue,DataY.Length-2,DataY.Length-1))
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
                    n+=1
                End If
            Next i
            Me.Slope=(n*SumXY-SumX*SumY)/(n*SumXX-SumX^2)
            Me.Intercept=(SumX*SumXY-SumY*SumXX)/(SumX^2-n*SumXX)
            Me.ModifiedCorrelation=SumXY/Math.Sqrt(SumXX)/Math.Sqrt(SumYY)
        End Sub

    End Structure

    Public DBMPoints(-1) As DBMPoint

    #If OfflineUnitTests Then
    Public Sub New(Optional ByVal Data() As Double=Nothing)
        If Not IsNothing(Data) Then
            UnitTestData=Data
        End If
    End Sub
    #End If

    #If OfflineUnitTests Then
    Public Function DBMPointIndex(ByVal Point As String) As Integer
    #Else
    Public Function DBMPointIndex(ByVal Point As PISDK.PIPoint) As Integer
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
        If Calculate<>0 And Not IsNothing(CorrelationPoint) Then
            CorrelationDBMPointIndex=DBMPointIndex(CorrelationPoint)
            If SubstractInputPointFromCorrelationPoint Then
                DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True,DBMPoints(InputDBMPointIndex))
            Else
                DBMPoints(CorrelationDBMPointIndex).Calculate(Timestamp,False,True)
            End If
            AbsErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).AbsoluteError,DBMPoints(InputDBMPointIndex).AbsoluteError)
            RelErrorStats.Calculate(DBMPoints(CorrelationDBMPointIndex).RelativeError,DBMPoints(InputDBMPointIndex).RelativeError)
            If RelErrorStats.ModifiedCorrelation>CorrelationThreshold Then
                Calculate=RelErrorStats.ModifiedCorrelation
            End If
            If Not SubstractInputPointFromCorrelationPoint And AbsErrorStats.ModifiedCorrelation<-CorrelationThreshold Then
                Calculate=AbsErrorStats.ModifiedCorrelation
            End If
        End If
        Return Math.Round(Calculate,3)
    End Function

End Class
