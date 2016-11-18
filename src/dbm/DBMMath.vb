Option Explicit
Option Strict

Public Class DBMMath

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
        Dim Median,MedianAbsDev,Mean,MeanAbsDev As Double
        Dim i As Integer
        Median=CalculateMedian(Data.ToArray)
        MedianAbsDev=CalculateMedianAbsDev(Median,Data.ToArray)
        For i=0 to Data.Length-1
            If MedianAbsDev=0 Then ' Use Mean Absolute Deviation instead of Median Absolute Deviation to detect outliers
                Mean=CalculateMean(Data.ToArray)
                MeanAbsDev=CalculateMeanAbsDev(Mean,Data.ToArray)
                If Math.Abs(Data(i)-Mean)>MeanAbsDev*MeanAbsDevScaleFactor*ControlLimitRejectionCriterion(Data.Length) Then ' If value is an outlier
                    Data(i)=Double.NaN ' Exclude outlier
                End If
            Else ' Use Median Absolute Deviation to detect outliers
                If Math.Abs(Data(i)-Median)>MedianAbsDev*MedianAbsDevScaleFactor(Data.Length)*ControlLimitRejectionCriterion(Data.Length) Then ' If value is an outlier
                    Data(i)=Double.NaN ' Exclude outlier
                End If
            End If
        Next i
        Return Data
    End Function

    Public Function CalculateExpMovingAvg(ByVal Data() As Double) As Double ' Filter high frequency variation
        Dim Weight,TotalWeight As Double
        Weight=1 ' Initial weight
        TotalWeight=0
        For Each Value As Double In Data ' Most significant value first
            CalculateExpMovingAvg+=Value*Weight
            TotalWeight+=Weight
            Weight*=1-2/((Data.Length)+1) ' Decrease weight for older values
        Next
        CalculateExpMovingAvg/=TotalWeight
        Return CalculateExpMovingAvg
    End Function

End Class
