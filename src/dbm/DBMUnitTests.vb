Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
'
' This file is part of DBM.
'
' DBM is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' DBM is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with DBM.  If not, see <http://www.gnu.org/licenses/>.

Namespace DBM

    Public Class DBMUnitTests

        Private Shared Function ArrayMultiplier(Data() As Double) As Double
            ArrayMultiplier=1
            For Each Value In Data
                If Not Double.IsNaN(Value) Then
                    ArrayMultiplier=ArrayMultiplier*2+Value
                End If
            Next
            Return ArrayMultiplier
        End Function

        Public Shared Function Results As Boolean
            Dim i As Integer
            Dim DBMStatistics As New DBMStatistics
            Results=True

            Results=Results And ArrayMultiplier({6,4,7,1,1,4,2,4})=1552
            Results=Results And ArrayMultiplier({8,4,7,3,2,6,5,7})=1865
            Results=Results And ArrayMultiplier({1,5,4,7,7,8,5,1})=1043

            Results=Results And DBM.KeepOrSuppressEvent(5,0,0,False)=5
            Results=Results And DBM.KeepOrSuppressEvent(-5,0,0,False)=-5
            Results=Results And DBM.KeepOrSuppressEvent(5,-0.8,0,False)=5
            Results=Results And DBM.KeepOrSuppressEvent(5,-0.9,0,False)=-0.9
            Results=Results And DBM.KeepOrSuppressEvent(5,-0.9,0,True)=5
            Results=Results And DBM.KeepOrSuppressEvent(5,0,0.8,False)=5
            Results=Results And DBM.KeepOrSuppressEvent(5,0,0.9,False)=0.9
            Results=Results And DBM.KeepOrSuppressEvent(5,0,0.9,True)=0.9
            Results=Results And DBM.KeepOrSuppressEvent(-0.9,-0.85,0,False)=-0.9
            Results=Results And DBM.KeepOrSuppressEvent(-0.9,-0.95,0,False)=-0.95
            Results=Results And DBM.KeepOrSuppressEvent(0.9,0,0.85,False)=0.9
            Results=Results And DBM.KeepOrSuppressEvent(0.9,0,0.95,False)=0.95
            Results=Results And DBM.KeepOrSuppressEvent(0.9,0,0.95,True)=0.95
            Results=Results And DBM.KeepOrSuppressEvent(-0.9,-0.99,0,False)=-0.99
            Results=Results And DBM.KeepOrSuppressEvent(-0.9,0.99,0,False)=-0.9
            Results=Results And DBM.KeepOrSuppressEvent(-0.9,0.99,0,True)=-0.9
            Results=Results And DBM.KeepOrSuppressEvent(0.99,-0.9,0,False)=-0.9
            Results=Results And DBM.KeepOrSuppressEvent(0.99,-0.9,0,True)=0.99
            Results=Results And DBM.KeepOrSuppressEvent(-0.99,0,0,True)=-0.99
            Results=Results And DBM.KeepOrSuppressEvent(0.99,0,0,True)=0.99

            Results=Results And ArrayMultiplier(DBMFunctions.ArrayRotateLeft({1,2,3,4,5}))=ArrayMultiplier({2,3,4,5,1})
            Results=Results And ArrayMultiplier(DBMFunctions.ArrayRotateLeft({2,3,4,5,1}))=ArrayMultiplier({3,4,5,1,2})
            Results=Results And ArrayMultiplier(DBMFunctions.ArrayRotateLeft({3,4,5,1,2}))=ArrayMultiplier({4,5,1,2,3})
            Results=Results And ArrayMultiplier(DBMFunctions.ArrayRotateLeft({4,5,1,2,3}))=ArrayMultiplier({5,1,2,3,4})
            Results=Results And ArrayMultiplier(DBMFunctions.ArrayRotateLeft({5,1,2,3,4}))=ArrayMultiplier({1,2,3,4,5})

            Results=Results And Math.Round(DBMMath.NormSInv(0.7451),4)=0.6591
            Results=Results And Math.Round(DBMMath.NormSInv(0.4188),4)=-0.205
            Results=Results And Math.Round(DBMMath.NormSInv(0.1385),4)=-1.0871
            Results=Results And Math.Round(DBMMath.NormSInv(0.8974),4)=1.2669
            Results=Results And Math.Round(DBMMath.NormSInv(0.7663),4)=0.7267
            Results=Results And Math.Round(DBMMath.NormSInv(0.2248),4)=-0.7561
            Results=Results And Math.Round(DBMMath.NormSInv(0.9372),4)=1.5317
            Results=Results And Math.Round(DBMMath.NormSInv(0.4135),4)=-0.2186
            Results=Results And Math.Round(DBMMath.NormSInv(0.2454),4)=-0.689
            Results=Results And Math.Round(DBMMath.NormSInv(0.2711),4)=-0.6095
            Results=Results And Math.Round(DBMMath.NormSInv(0.2287),4)=-0.7431
            Results=Results And Math.Round(DBMMath.NormSInv(0.6517),4)=0.3899
            Results=Results And Math.Round(DBMMath.NormSInv(0.8663),4)=1.1091
            Results=Results And Math.Round(DBMMath.NormSInv(0.9275),4)=1.4574
            Results=Results And Math.Round(DBMMath.NormSInv(0.7089),4)=0.5502
            Results=Results And Math.Round(DBMMath.NormSInv(0.1234),4)=-1.1582
            Results=Results And Math.Round(DBMMath.NormSInv(0.0837),4)=-1.3806
            Results=Results And Math.Round(DBMMath.NormSInv(0.6243),4)=0.3168
            Results=Results And Math.Round(DBMMath.NormSInv(0.0353),4)=-1.808
            Results=Results And Math.Round(DBMMath.NormSInv(0.9767),4)=1.9899

            Results=Results And Math.Round(DBMMath.TInv2T(0.3353,16),4)=0.9934
            Results=Results And Math.Round(DBMMath.TInv2T(0.4792,12),4)=0.7303
            Results=Results And Math.Round(DBMMath.TInv2T(0.4384,9),4)=0.8108
            Results=Results And Math.Round(DBMMath.TInv2T(0.0905,6),4)=2.0152
            Results=Results And Math.Round(DBMMath.TInv2T(0.63,16),4)=0.4911
            Results=Results And Math.Round(DBMMath.TInv2T(0.1533,11),4)=1.5339
            Results=Results And Math.Round(DBMMath.TInv2T(0.6297,12),4)=0.4948
            Results=Results And Math.Round(DBMMath.TInv2T(0.1512,4),4)=1.7714
            Results=Results And Math.Round(DBMMath.TInv2T(0.4407,18),4)=0.7884
            Results=Results And Math.Round(DBMMath.TInv2T(0.6169,15),4)=0.5108
            Results=Results And Math.Round(DBMMath.TInv2T(0.6077,18),4)=0.5225
            Results=Results And Math.Round(DBMMath.TInv2T(0.4076,20),4)=0.8459
            Results=Results And Math.Round(DBMMath.TInv2T(0.1462,18),4)=1.5187
            Results=Results And Math.Round(DBMMath.TInv2T(0.3421,6),4)=1.0315
            Results=Results And Math.Round(DBMMath.TInv2T(0.6566,6),4)=0.4676
            Results=Results And Math.Round(DBMMath.TInv2T(0.2986,1),4)=1.9733
            Results=Results And Math.Round(DBMMath.TInv2T(0.2047,14),4)=1.3303
            Results=Results And Math.Round(DBMMath.TInv2T(0.5546,2),4)=0.7035
            Results=Results And Math.Round(DBMMath.TInv2T(0.0862,6),4)=2.0504
            Results=Results And Math.Round(DBMMath.TInv2T(0.6041,10),4)=0.5354

            Results=Results And Math.Round(DBMMath.TInv(0.4097,8),4)=-0.2359
            Results=Results And Math.Round(DBMMath.TInv(0.174,19),4)=-0.9623
            Results=Results And Math.Round(DBMMath.TInv(0.6545,15),4)=0.4053
            Results=Results And Math.Round(DBMMath.TInv(0.7876,5),4)=0.8686
            Results=Results And Math.Round(DBMMath.TInv(0.2995,3),4)=-0.5861
            Results=Results And Math.Round(DBMMath.TInv(0.0184,2),4)=-5.0679
            Results=Results And Math.Round(DBMMath.TInv(0.892,1),4)=2.8333
            Results=Results And Math.Round(DBMMath.TInv(0.7058,18),4)=0.551
            Results=Results And Math.Round(DBMMath.TInv(0.3783,2),4)=-0.3549
            Results=Results And Math.Round(DBMMath.TInv(0.2774,15),4)=-0.6041
            Results=Results And Math.Round(DBMMath.TInv(0.0406,8),4)=-1.9945
            Results=Results And Math.Round(DBMMath.TInv(0.1271,4),4)=-1.3303
            Results=Results And Math.Round(DBMMath.TInv(0.241,18),4)=-0.718
            Results=Results And Math.Round(DBMMath.TInv(0.0035,1),4)=-90.942
            Results=Results And Math.Round(DBMMath.TInv(0.1646,10),4)=-1.0257
            Results=Results And Math.Round(DBMMath.TInv(0.279,11),4)=-0.6041
            Results=Results And Math.Round(DBMMath.TInv(0.8897,4),4)=1.4502
            Results=Results And Math.Round(DBMMath.TInv(0.5809,13),4)=0.2083
            Results=Results And Math.Round(DBMMath.TInv(0.3776,11),4)=-0.3197
            Results=Results And Math.Round(DBMMath.TInv(0.5267,15),4)=0.0681

            Results=Results And Math.Round(DBMMath.MeanAbsDevScaleFactor,4)=1.2533

            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(1),4)=1
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(2),4)=1.2247
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(4),4)=1.3501
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(6),4)=1.3936
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(8),4)=1.4157
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(10),4)=1.429
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(12),4)=1.4378
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(14),4)=1.4442
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(16),4)=1.449
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(18),4)=1.4527
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(20),4)=1.4557
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(22),4)=1.4581
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(24),4)=1.4602
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(26),4)=1.4619
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(28),4)=1.4634
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(30),4)=1.4826
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(32),4)=1.4826
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(34),4)=1.4826
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(36),4)=1.4826
            Results=Results And Math.Round(DBMMath.MedianAbsDevScaleFactor(38),4)=1.4826

            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,1),4)=63.6567
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,2),4)=9.9248
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,4),4)=4.6041
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,6),4)=3.7074
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,8),4)=3.3554
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,10),4)=3.1693
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,12),4)=3.0545
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,14),4)=2.9768
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,16),4)=2.9208
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,20),4)=2.8453
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,22),4)=2.8188
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,24),4)=2.7969
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,28),4)=2.7633
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,30),4)=2.5758
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,32),4)=2.5758
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.99,34),4)=2.5758
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.95,30),4)=1.96
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.90,30),4)=1.6449
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.95,25),4)=2.0595
            Results=Results And Math.Round(DBMMath.ControlLimitRejectionCriterion(0.90,20),4)=1.7247

            Results=Results And Math.Round(DBMMath.CalculateMean({60}),4)=60
            Results=Results And Math.Round(DBMMath.CalculateMean({72}),4)=72
            Results=Results And Math.Round(DBMMath.CalculateMean({32,95}),4)=63.5
            Results=Results And Math.Round(DBMMath.CalculateMean({81,75}),4)=78
            Results=Results And Math.Round(DBMMath.CalculateMean({67,76,25}),4)=56
            Results=Results And Math.Round(DBMMath.CalculateMean({22,70,73}),4)=55
            Results=Results And Math.Round(DBMMath.CalculateMean({31,73,83,81}),4)=67
            Results=Results And Math.Round(DBMMath.CalculateMean({18,58,47,47}),4)=42.5
            Results=Results And Math.Round(DBMMath.CalculateMean({22,67,45,31,76}),4)=48.2
            Results=Results And Math.Round(DBMMath.CalculateMean({67,90,74,32,62}),4)=65
            Results=Results And Math.Round(DBMMath.CalculateMean({78,0,98,65,69,57}),4)=61.1667
            Results=Results And Math.Round(DBMMath.CalculateMean({49,35,74,25,28,92}),4)=50.5
            Results=Results And Math.Round(DBMMath.CalculateMean({7,64,22,7,42,34,30}),4)=29.4286
            Results=Results And Math.Round(DBMMath.CalculateMean({13,29,39,33,96,43,17}),4)=38.5714
            Results=Results And Math.Round(DBMMath.CalculateMean({59,78,53,7,18,44,63,40}),4)=45.25
            Results=Results And Math.Round(DBMMath.CalculateMean({77,71,99,39,50,94,67,30}),4)=65.875
            Results=Results And Math.Round(DBMMath.CalculateMean({91,69,63,5,44,93,89,45,50}),4)=61
            Results=Results And Math.Round(DBMMath.CalculateMean({12,84,12,94,52,17,1,13,37}),4)=35.7778
            Results=Results And Math.Round(DBMMath.CalculateMean({18,14,54,40,73,77,4,91,53,10}),4)=43.4
            Results=Results And Math.Round(DBMMath.CalculateMean({80,30,1,92,44,61,18,72,63,41}),4)=50.2

            Results=Results And DBMMath.CalculateMedian({57})=57
            Results=Results And DBMMath.CalculateMedian({46})=46
            Results=Results And DBMMath.CalculateMedian({79,86})=82.5
            Results=Results And DBMMath.CalculateMedian({46,45})=45.5
            Results=Results And DBMMath.CalculateMedian({21,81,13})=21
            Results=Results And DBMMath.CalculateMedian({58,79,68})=68
            Results=Results And DBMMath.CalculateMedian({60,42,59,23})=50.5
            Results=Results And DBMMath.CalculateMedian({33,23,36,47})=34.5
            Results=Results And DBMMath.CalculateMedian({99,10,41,40,55})=41
            Results=Results And DBMMath.CalculateMedian({19,3,9,14,31})=14
            Results=Results And DBMMath.CalculateMedian({2,85,33,10,38,56})=35.5
            Results=Results And DBMMath.CalculateMedian({42,65,57,92,56,59})=58
            Results=Results And DBMMath.CalculateMedian({53,17,18,73,34,96,9})=34
            Results=Results And DBMMath.CalculateMedian({23,23,43,74,8,51,88})=43
            Results=Results And DBMMath.CalculateMedian({72,46,39,7,83,96,18,50})=48
            Results=Results And DBMMath.CalculateMedian({50,25,28,99,79,97,30,16})=40
            Results=Results And DBMMath.CalculateMedian({7,22,58,32,5,90,46,91,66})=46
            Results=Results And DBMMath.CalculateMedian({81,64,5,23,48,18,19,87,15})=23
            Results=Results And DBMMath.CalculateMedian({33,82,42,33,81,56,13,13,54,6})=37.5
            Results=Results And DBMMath.CalculateMedian({55,40,75,23,53,85,59,9,72,44})=54

            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({19}),4)=0
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({86}),4)=0
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({7,24}),4)=8.5
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({12,96}),4)=42
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({19,74,70}),4)=23.5556
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({96,93,65}),4)=13.1111
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({47,29,24,11}),4)=10.25
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({3,43,53,80}),4)=21.75
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({17,27,98,85,51}),4)=28.72
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({2,82,63,1,49}),4)=30.32
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({9,25,41,85,82,55}),4)=24.5
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({5,74,53,97,81,21}),4)=28.8333
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({26,81,9,18,39,97,21}),4)=27.102
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({5,83,31,24,55,22,87}),4)=26.6939
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({22,84,6,79,89,71,34,56}),4)=25.8438
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({33,39,6,88,69,11,76,65}),4)=26.125
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({31,52,12,60,52,44,47,81,34}),4)=13.9012
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({64,63,54,94,25,80,97,45,51}),4)=17.8519
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({47,22,52,22,10,38,94,85,54,41}),4)=19.9
            Results=Results And Math.Round(DBMMath.CalculateMeanAbsDev({7,12,84,29,41,8,18,15,16,84}),4)=22.96

            Results=Results And DBMMath.CalculateMedianAbsDev({2})=0
            Results=Results And DBMMath.CalculateMedianAbsDev({37})=0
            Results=Results And DBMMath.CalculateMedianAbsDev({87,37})=25
            Results=Results And DBMMath.CalculateMedianAbsDev({13,74})=30.5
            Results=Results And DBMMath.CalculateMedianAbsDev({39,52,93})=13
            Results=Results And DBMMath.CalculateMedianAbsDev({90,24,47})=23
            Results=Results And DBMMath.CalculateMedianAbsDev({11,51,20,62})=20
            Results=Results And DBMMath.CalculateMedianAbsDev({74,35,9,95})=30
            Results=Results And DBMMath.CalculateMedianAbsDev({32,46,15,90,66})=20
            Results=Results And DBMMath.CalculateMedianAbsDev({91,19,50,55,44})=6
            Results=Results And DBMMath.CalculateMedianAbsDev({2,64,87,65,61,97})=13
            Results=Results And DBMMath.CalculateMedianAbsDev({35,66,73,74,71,93})=4
            Results=Results And DBMMath.CalculateMedianAbsDev({54,81,80,36,11,36,45})=9
            Results=Results And DBMMath.CalculateMedianAbsDev({14,69,40,68,75,10,69})=7
            Results=Results And DBMMath.CalculateMedianAbsDev({40,51,28,21,91,95,66,3})=22.5
            Results=Results And DBMMath.CalculateMedianAbsDev({57,87,94,46,51,27,10,7})=30
            Results=Results And DBMMath.CalculateMedianAbsDev({3,89,62,84,86,37,14,72,33})=25
            Results=Results And DBMMath.CalculateMedianAbsDev({48,6,14,2,74,89,15,8,83})=13
            Results=Results And DBMMath.CalculateMedianAbsDev({2,22,91,84,43,96,55,3,9,11})=26.5
            Results=Results And DBMMath.CalculateMedianAbsDev({6,96,82,26,47,84,34,39,60,99})=28

            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2282.4634,3233.9927,2313.853,2145.928,2809.7371,2213.207,2193.105,2557.4885,3927.1062,2304.9473,2086.3416,2872.4075}))=ArrayMultiplier({2282.4634,Double.NaN,2313.853,2145.928,2809.7371,2213.207,2193.105,2557.4885,Double.NaN,2304.9473,2086.3416,2872.4075})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2922.467,2654.3137,2531.8938,2554.7935,2894.1758,2526.7852,2051.6868,3083.7932,2542.0398,2378.6038,2610.0007,3567.543}))=ArrayMultiplier({2922.467,2654.3137,2531.8938,2554.7935,2894.1758,2526.7852,2051.6868,3083.7932,2542.0398,2378.6038,2610.0007,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2223.6946,2770.1624,2125.7544,3948.9927,2184.2341,2238.6421,2170.0227,2967.0674,2177.3738,3617.1328,2460.8193,3315.8684}))=ArrayMultiplier({2223.6946,2770.1624,2125.7544,Double.NaN,2184.2341,2238.6421,2170.0227,2967.0674,2177.3738,Double.NaN,2460.8193,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3355.1553,3624.3154,3317.6895,3610.0039,3990.751,2950.4382,2140.5908,3237.4917,3319.7139,2829.2725,3406.9199,3230.0078}))=ArrayMultiplier({3355.1553,3624.3154,3317.6895,3610.0039,3990.751,2950.4382,Double.NaN,3237.4917,3319.7139,2829.2725,3406.9199,3230.0078})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2969.7808,3899.0913,2637.4045,2718.73,2960.9597,2650.6521,2707.4294,2034.5339,2935.9111,3458.7085,2584.53,3999.4238}))=ArrayMultiplier({2969.7808,Double.NaN,2637.4045,2718.73,2960.9597,2650.6521,2707.4294,2034.5339,2935.9111,3458.7085,2584.53,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2774.8018,2755.0251,2756.6152,3800.0625,2900.0671,2784.0134,3955.2947,2847.0908,2329.7837,3282.4614,2597.1582,3009.8796}))=ArrayMultiplier({2774.8018,2755.0251,2756.6152,Double.NaN,2900.0671,2784.0134,Double.NaN,2847.0908,2329.7837,3282.4614,2597.1582,3009.8796})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3084.8821,3394.1196,3131.3245,2799.9587,2528.3088,3015.4998,2912.2029,2022.2645,3666.5674,3685.1973,3149.6931,3070.0479}))=ArrayMultiplier({3084.8821,3394.1196,3131.3245,2799.9587,2528.3088,3015.4998,2912.2029,Double.NaN,3666.5674,3685.1973,3149.6931,3070.0479})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3815.72,3063.9106,3535.0366,2349.564,2597.2661,3655.3076,3452.7407,2020.7682,3810.7046,3833.8396,3960.6016,3866.8149}))=ArrayMultiplier({3815.72,3063.9106,3535.0366,Double.NaN,2597.2661,3655.3076,3452.7407,Double.NaN,3810.7046,3833.8396,3960.6016,3866.8149})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2812.0613,3726.7427,2090.9749,2548.4485,3900.5151,3545.854,3880.2229,3940.9585,3942.2234,3263.0137,3701.8882,2056.5291}))=ArrayMultiplier({2812.0613,3726.7427,Double.NaN,2548.4485,3900.5151,3545.854,3880.2229,3940.9585,3942.2234,3263.0137,3701.8882,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3798.4775,2959.3879,2317.3547,2596.3599,2075.6292,2563.9685,2695.5081,2386.2161,2433.1106,2810.3716,2499.7554,3843.103}))=ArrayMultiplier({Double.NaN,2959.3879,2317.3547,2596.3599,2075.6292,2563.9685,2695.5081,2386.2161,2433.1106,2810.3716,2499.7554,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2245.7856,2012.4834,2473.0103,2684.5693,2645.4729,2851.019,2344.6099,2408.1492,3959.5967,3954.0583,2399.2617,2652.8855}))=ArrayMultiplier({2245.7856,2012.4834,2473.0103,2684.5693,2645.4729,2851.019,2344.6099,2408.1492,Double.NaN,Double.NaN,2399.2617,2652.8855})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2004.5355,2743.0693,3260.7441,2382.8906,2365.9385,2243.333,3506.5352,3905.7717,3516.5337,2133.8328,2308.1809,2581.4009}))=ArrayMultiplier({2004.5355,2743.0693,3260.7441,2382.8906,2365.9385,2243.333,3506.5352,Double.NaN,3516.5337,2133.8328,2308.1809,2581.4009})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3250.5376,3411.313,2037.264,3709.5815,3417.1167,3996.0493,3529.637,3992.7163,2786.95,3728.834,3304.4272,2248.9119}))=ArrayMultiplier({3250.5376,3411.313,Double.NaN,3709.5815,3417.1167,3996.0493,3529.637,3992.7163,2786.95,3728.834,3304.4272,2248.9119})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2398.3125,2742.4028,2720.752,2628.8442,2750.1482,2724.4932,2161.6875,2644.4163,2188.2952,2455.4622,3332.5503,2540.5198}))=ArrayMultiplier({2398.3125,2742.4028,2720.752,2628.8442,2750.1482,2724.4932,2161.6875,2644.4163,2188.2952,2455.4622,Double.NaN,2540.5198})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3991.7854,3607.98,2686.032,2546.969,3053.8796,3138.9824,2441.1689,2737.1245,2616.7139,2550.5774,2406.0913,2743.2361}))=ArrayMultiplier({Double.NaN,3607.98,2686.032,2546.969,3053.8796,3138.9824,2441.1689,2737.1245,2616.7139,2550.5774,2406.0913,2743.2361})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2361.5334,3636.4312,2187.593,2281.5432,2132.3833,2056.792,2227.7795,2757.1753,3416.9126,2568.927,2094.2065,3449.3984}))=ArrayMultiplier({2361.5334,Double.NaN,2187.593,2281.5432,2132.3833,2056.792,2227.7795,2757.1753,Double.NaN,2568.927,2094.2065,Double.NaN})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2249.7119,2411.8374,3041.5498,2679.1458,2561.1577,2405.7229,2775.2253,2832.1233,2540.2134,3654.5903,3970.5173,2920.5637}))=ArrayMultiplier({2249.7119,2411.8374,3041.5498,2679.1458,2561.1577,2405.7229,2775.2253,2832.1233,2540.2134,3654.5903,Double.NaN,2920.5637})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({2038.1091,2248.3057,2427.1646,2337.2427,2642.043,3497.5393,3996.3579,2178.979,3968.8848,3460.8613,2774.8486,2338.1362}))=ArrayMultiplier({2038.1091,2248.3057,2427.1646,2337.2427,2642.043,3497.5393,Double.NaN,2178.979,Double.NaN,3460.8613,2774.8486,2338.1362})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3010.9485,2517.2876,2057.7188,2133.0801,3192.0308,2035.0759,3821.248,2391.8086,2267.896,3751.3276,2340.9497,2327.333}))=ArrayMultiplier({3010.9485,2517.2876,2057.7188,2133.0801,3192.0308,2035.0759,Double.NaN,2391.8086,2267.896,Double.NaN,2340.9497,2327.333})
            Results=Results And ArrayMultiplier(DBMMath.RemoveOutliers({3009.2004,3779.3677,3500.9705,3114.8486,3312.1548,3478.4153,3080.0972,2848.6597,3197.1477,3064.1587,3100.6992,2953.3582}))=ArrayMultiplier({3009.2004,Double.NaN,3500.9705,3114.8486,3312.1548,3478.4153,3080.0972,2848.6597,3197.1477,3064.1587,3100.6992,2953.3582})

            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({70.5547}),4)=70.5547
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({53.3424,57.9519}),4)=54.4948
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({28.9562,30.1948}),4)=29.2658
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({77.474,1.4018}),4)=58.456
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({76.0724,81.449,70.9038}),4)=76.8702
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({4.5353,41.4033,86.2619}),4)=26.7442
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({79.048,37.3536,96.1953}),4)=69.5849
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({87.1446,5.6237,94.9557}),4)=64.9688
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({36.4019,52.4868,76.7112}),4)=46.7561
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({5.3505,59.2458,46.87,29.8165}),4)=29.509
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({62.2697,64.7821,26.3793,27.9342}),4)=53.6164
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({82.9802,82.4602,58.9163,98.6093}),4)=80.4071
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({91.0964,22.6866,69.5116,98.0003,24.3931}),4)=65.6332
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({53.3873,10.637,99.9415,67.6176,1.5704}),4)=48.0787
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({57.5184,10.0052,10.3023,79.8884,28.448}),4)=37.6429
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({4.5649,29.5773,38.2011,30.097,94.8571}),4)=26.456
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({97.9829,40.1374,27.828,16.0442,16.2822}),4)=55.6939
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({64.6587,41.0073,41.2767,71.273,32.6206,63.3179}),4)=53.1265
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({20.7561,18.6014,58.3359,8.0715,45.7971,90.573}),4)=31.4677
            Results=Results And Math.Round(DBMMath.CalculateExpMovingAvg({26.1368,78.5212,37.8903,28.9665,91.9377,63.1742}),4)=48.6925

            For i=0 To 19
                If i=0 Then DBMStatistics.Calculate({3411,3067,3159,2579,2604,3549,2028,3521,3629,3418,2091,2828})
                If i=1 Then DBMStatistics.Calculate({3725,3581,2747,3924,3743,2112,3899,2728,3050,3534,2107,3185})
                If i=2 Then DBMStatistics.Calculate({2937,2596,3245,3296,2528,2559,3660,3649,3178,3972,3822,2454})
                If i=3 Then DBMStatistics.Calculate({3390,3960,2488,3068,2213,3999,3352,2031,3150,2200,2206,3598})
                If i=4 Then DBMStatistics.Calculate({2569,2091,2592,2764,2602,3897,3960,2803,2557,2321,2326,3293})
                If i=5 Then DBMStatistics.Calculate({2820,2826,3425,2652,3266,2415,2372,3167,2161,2916,3811,2523})
                If i=6 Then DBMStatistics.Calculate({3570,2758,2579,3839,3263,3255,2857,2196,3122,3389,3827,3670})
                If i=7 Then DBMStatistics.Calculate({2045,3087,3832,2861,3356,3005,3027,2926,2707,2810,2539,2111})
                If i=8 Then DBMStatistics.Calculate({2488,3958,2122,2781,2730,2980,2311,2949,2515,3258,3084,2313})
                If i=9 Then DBMStatistics.Calculate({3877,3309,3012,2781,2215,3568,2919,3507,3192,3665,2038,2421})
                If i=10 Then DBMStatistics.Calculate({2148,2211,2663,2256,2000,3074,3314,3088,3655,2164,2384,3358},{3,5,10,20,28,32,41,46,57,66,74,76})
                If i=11 Then DBMStatistics.Calculate({2908,2714,2300,3409,3858,3060,2179,3515,2804,2924,2984,2415},{10,18,28,35,44,50,51,62,63,66,74,80})
                If i=12 Then DBMStatistics.Calculate({2659,2191,3180,2340,3855,2196,2888,2546,3745,3501,2546,3347},{9,11,13,18,21,26,28,37,42,47,56,61})
                If i=13 Then DBMStatistics.Calculate({2513,2180,2062,2645,3580,2595,2471,2961,2509,2681,2090,2965},{6,14,23,26,32,35,37,41,45,48,58,68})
                If i=14 Then DBMStatistics.Calculate({2412,3729,3177,3510,3856,2662,3086,2161,3269,2820,3921,2229},{7,12,23,29,30,32,36,38,41,49,55,61})
                If i=15 Then DBMStatistics.Calculate({3847,3240,2695,2298,2960,2439,3987,2261,2058,2691,3095,3846},{4,9,15,18,20,26,36,45,49,58,64,71})
                If i=16 Then DBMStatistics.Calculate({3076,2813,3694,3652,3345,3444,3994,2680,2990,2826,3391,2358},{7,14,24,28,37,40,49,51,55,61,69,77})
                If i=17 Then DBMStatistics.Calculate({2846,3086,3629,3082,2855,3018,2456,3238,2980,3362,3773,2741},{6,16,23,29,35,40,49,60,64,73,75,78})
                If i=18 Then DBMStatistics.Calculate({2605,2586,2301,3060,2447,3169,2727,3752,2956,2381,3368,3495},{6,13,24,30,38,47,57,59,69,77,86,96})
                If i=19 Then DBMStatistics.Calculate({3228,3564,2323,3616,2405,3914,2132,2123,3586,2759,2927,2239},{10,15,21,22,24,32,34,43,46,53,55,63})
                Results=Results And Math.Round(DBMStatistics.Slope,4)={-24.1399,-67.5699,51.3427,-56.9825,27.3182,-2.6573,32.1923,-46.8462,-11.1224,-61.5455,9.4424,-0.1602,10.7659,4.8889,-4.6572,-0.4548,-6.6652,2.2442,8.6539,-12.0462}(i)
                Results=Results And Math.Round(DBMStatistics.Intercept,4)={3123.1026,3566.2179,2875.6154,3284.6538,2664.3333,2877.4487,3016.6923,3116.4872,2851.9231,3380.5,2332.5305,2930.2549,2585.1154,2427.9251,3229.618,2967.1468,3472.9635,2986.3475,2469.7771,3320.9411}(i)
                Results=Results And Math.Round(DBMStatistics.StDevSLinReg,4)={582.4218,633.706,535.0359,720.9024,619.3358,506.8629,525.9328,483.3154,527.9273,573.7699,676.6992,523.5618,678.6414,472.6452,663.1077,698.3863,563.5002,399.9662,638.4709,781.7412}(i)
                Results=Results And Math.Round(DBMStatistics.Correlation,4)={-0.1548,-0.374,0.3411,-0.2864,0.1645,-0.0198,0.2255,-0.3441,-0.0794,-0.3759,0.4296,-0.0071,0.3208,0.2029,-0.1209,-0.0154,-0.3016,0.1484,0.5294,-0.3121}(i)
                Results=Results And Math.Round(DBMStatistics.ModifiedCorrelation,4)={0.819,0.7932,0.8652,0.7909,0.8474,0.8345,0.8554,0.8061,0.8274,0.7962,0.8665,0.9037,0.8892,0.908,0.887,0.8274,0.8714,0.8915,0.9047,0.8561}(i)
                Results=Results And Math.Round(DBMStatistics.Determination,4)={0.024,0.1398,0.1164,0.082,0.0271,0.0004,0.0509,0.1184,0.0063,0.1413,0.1845,0.0001,0.1029,0.0412,0.0146,0.0002,0.091,0.022,0.2803,0.0974}(i)
            Next i

            Return Results
        End Function

    End Class

End Namespace
