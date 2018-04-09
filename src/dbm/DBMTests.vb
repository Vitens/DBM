Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014, 2015, 2016, 2017, 2018  J.H. Fitié, Vitens N.V.
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


Imports System
Imports System.Collections.Generic
Imports System.DateTime
Imports System.Double
Imports System.Math
Imports System.TimeSpan
Imports Vitens.DynamicBandwidthMonitor.DBM
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMPoint
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTests


    Private Shared Function Hash(Values() As Double, _
      Optional Digits As Integer = 15) As Double

      ' Simple hash function for checking array contents.

      Dim Value As Double

      Hash = 1
      For Each Value In Values
        If Not IsNaN(Value) Then
          Hash = (Hash+Round(Value, Digits)+1)/3
        End If
      Next

      Return Hash

    End Function


    Public Shared Function UnitTestsPassed As Boolean

      ' Unit tests, returns True if all tests pass.

      Dim i As Integer
      Dim StatisticsData As New DBMStatisticsData

      UnitTestsPassed = True

      UnitTestsPassed = UnitTestsPassed And _
        Round(Hash({6, 4, 7, 1, 1, 4, 2, 4}), 4) = 2.2326 And _
        Round(Hash({8, 4, 7, 3, 2, 6, 5, 7}), 4) = 3.6609 And _
        Round(Hash({1, 5, 4, 7, 7, 8, 5, 1}), 4) = 1.8084

      UnitTestsPassed = UnitTestsPassed And _
        HasCorrelation(0.9, 45) = True And _
        HasCorrelation(0.85, 45) = True And _
        HasCorrelation(0.8, 45) = False And _
        HasCorrelation(0.9, 0) = False And _
        HasCorrelation(0.9, 35) = True And _
        HasCorrelation(0.9, 25) = False And _
        HasCorrelation(0.8, 35) = False And _
        HasCorrelation(-0.9, 45) = False And _
        HasCorrelation(0.9, -45) = False And _
        HasCorrelation(0, 0) = False

      UnitTestsPassed = UnitTestsPassed And _
        HasAnticorrelation(-0.9, -45, False) = True And _
        HasAnticorrelation(-0.85, -45, True) = False And _
        HasAnticorrelation(-0.8, -45, False) = False And _
        HasAnticorrelation(-0.9, 0, False) = False And _
        HasAnticorrelation(-0.9, -35, False) = True And _
        HasAnticorrelation(-0.9, -25, False) = False And _
        HasAnticorrelation(-0.8, -35, False) = False And _
        HasAnticorrelation(0.9, -45, False) = False And _
        HasAnticorrelation(-0.9, 45, False) = False And _
        HasAnticorrelation(0, 0, False) = False

      UnitTestsPassed = UnitTestsPassed And _
        Suppress(5, 0, 0, 0, 0, False) = 5 And _
        Suppress(5, -0.8, 0, 0, 0, False) = 5 And _
        Suppress(5, -0.9, -45, 0, 0, False) = 0 And _
        Suppress(5, -0.9, 0, 0, 0, True) = 5 And _
        Suppress(5, 0, 0, 0.8, 0, False) = 5 And _
        Suppress(5, 0, 0, 0.9, 45, False) = 0 And _
        Suppress(5, 0, 0, 0.9, 45, True) = 0 And _
        Suppress(-0.9, -0.85, 0, 0, 0, False) = -0.9 And _
        Suppress(-0.9, -0.95, -45, 0, 0, False) = 0 And _
        Suppress(0.9, 0, 0, 0.85, 45, False) = 0 And _
        Suppress(0.9, 0, 0, 0.95, 1, True) = 0.9 And _
        Suppress(-0.9, -0.99, -45, 0, 0, False) = 0 And _
        Suppress(-0.9, 0.99, 0, 0, 0, False) = -0.9 And _
        Suppress(-0.9, 0.99, 0, 0, 0, True) = -0.9 And _
        Suppress(0.99, -0.9, -45, 0, 0, False) = 0 And _
        Suppress(0.99, -0.9, 0, 0, 0, True) = 0.99 And _
        Suppress(-0.99, 0, 0, 0, 0, True) = -0.99 And _
        Suppress(0.99, 0, 0, 0, 0, True) = 0.99 And _
        Suppress(-0.98, -0.99, -1, 0, 0, False) = -0.98 And _
        Suppress(-0.98, -0.99, -45, 0, 0, False) = 0

      UnitTestsPassed = UnitTestsPassed And _
        Round(NormSInv(0.7451), 4) = 0.6591 And _
        Round(NormSInv(0.4188), 4) = -0.205 And _
        Round(NormSInv(0.1385), 4) = -1.0871 And _
        Round(NormSInv(0.8974), 4) = 1.2669 And _
        Round(NormSInv(0.7663), 4) = 0.7267 And _
        Round(NormSInv(0.2248), 4) = -0.7561 And _
        Round(NormSInv(0.9372), 4) = 1.5317 And _
        Round(NormSInv(0.4135), 4) = -0.2186 And _
        Round(NormSInv(0.2454), 4) = -0.689 And _
        Round(NormSInv(0.2711), 4) = -0.6095 And _
        Round(NormSInv(0.2287), 4) = -0.7431 And _
        Round(NormSInv(0.6517), 4) = 0.3899 And _
        Round(NormSInv(0.8663), 4) = 1.1091 And _
        Round(NormSInv(0.9275), 4) = 1.4574 And _
        Round(NormSInv(0.7089), 4) = 0.5502 And _
        Round(NormSInv(0.1234), 4) = -1.1582 And _
        Round(NormSInv(0.0837), 4) = -1.3806 And _
        Round(NormSInv(0.6243), 4) = 0.3168 And _
        Round(NormSInv(0.0353), 4) = -1.808 And _
        Round(NormSInv(0.9767), 4) = 1.9899

      UnitTestsPassed = UnitTestsPassed And _
        Round(TInv2T(0.3353, 16), 4) = 0.9934 And _
        Round(TInv2T(0.4792, 12), 4) = 0.7303 And _
        Round(TInv2T(0.4384, 9), 4) = 0.8108 And _
        Round(TInv2T(0.0905, 6), 4) = 2.0152 And _
        Round(TInv2T(0.63, 16), 4) = 0.4911 And _
        Round(TInv2T(0.1533, 11), 4) = 1.5339 And _
        Round(TInv2T(0.6297, 12), 4) = 0.4948 And _
        Round(TInv2T(0.1512, 4), 4) = 1.7714 And _
        Round(TInv2T(0.4407, 18), 4) = 0.7884 And _
        Round(TInv2T(0.6169, 15), 4) = 0.5108 And _
        Round(TInv2T(0.6077, 18), 4) = 0.5225 And _
        Round(TInv2T(0.4076, 20), 4) = 0.8459 And _
        Round(TInv2T(0.1462, 18), 4) = 1.5187 And _
        Round(TInv2T(0.3421, 6), 4) = 1.0315 And _
        Round(TInv2T(0.6566, 6), 4) = 0.4676 And _
        Round(TInv2T(0.2986, 1), 4) = 1.9733 And _
        Round(TInv2T(0.2047, 14), 4) = 1.3303 And _
        Round(TInv2T(0.5546, 2), 4) = 0.7035 And _
        Round(TInv2T(0.0862, 6), 4) = 2.0504 And _
        Round(TInv2T(0.6041, 10), 4) = 0.5354

      UnitTestsPassed = UnitTestsPassed And _
        Round(TInv(0.4097, 8), 4) = -0.2359 And _
        Round(TInv(0.174, 19), 4) = -0.9623 And _
        Round(TInv(0.6545, 15), 4) = 0.4053 And _
        Round(TInv(0.7876, 5), 4) = 0.8686 And _
        Round(TInv(0.2995, 3), 4) = -0.5861 And _
        Round(TInv(0.0184, 2), 4) = -5.0679 And _
        Round(TInv(0.892, 1), 4) = 2.8333 And _
        Round(TInv(0.7058, 18), 4) = 0.551 And _
        Round(TInv(0.3783, 2), 4) = -0.3549 And _
        Round(TInv(0.2774, 15), 4) = -0.6041 And _
        Round(TInv(0.0406, 8), 4) = -1.9945 And _
        Round(TInv(0.1271, 4), 4) = -1.3303 And _
        Round(TInv(0.241, 18), 4) = -0.718 And _
        Round(TInv(0.0035, 1), 4) = -90.942 And _
        Round(TInv(0.1646, 10), 4) = -1.0257 And _
        Round(TInv(0.279, 11), 4) = -0.6041 And _
        Round(TInv(0.8897, 4), 4) = 1.4502 And _
        Round(TInv(0.5809, 13), 4) = 0.2083 And _
        Round(TInv(0.3776, 11), 4) = -0.3197 And _
        Round(TInv(0.5267, 15), 4) = 0.0681

      UnitTestsPassed = UnitTestsPassed And _
        Round(MeanAbsoluteDeviationScaleFactor, 4) = 1.2533

      UnitTestsPassed = UnitTestsPassed And _
        Round(MedianAbsoluteDeviationScaleFactor(1), 4) = 1 And _
        Round(MedianAbsoluteDeviationScaleFactor(2), 4) = 1.2247 And _
        Round(MedianAbsoluteDeviationScaleFactor(4), 4) = 1.3501 And _
        Round(MedianAbsoluteDeviationScaleFactor(6), 4) = 1.3936 And _
        Round(MedianAbsoluteDeviationScaleFactor(8), 4) = 1.4157 And _
        Round(MedianAbsoluteDeviationScaleFactor(10), 4) = 1.429 And _
        Round(MedianAbsoluteDeviationScaleFactor(12), 4) = 1.4378 And _
        Round(MedianAbsoluteDeviationScaleFactor(14), 4) = 1.4442 And _
        Round(MedianAbsoluteDeviationScaleFactor(16), 4) = 1.449 And _
        Round(MedianAbsoluteDeviationScaleFactor(18), 4) = 1.4527 And _
        Round(MedianAbsoluteDeviationScaleFactor(20), 4) = 1.4557 And _
        Round(MedianAbsoluteDeviationScaleFactor(22), 4) = 1.4581 And _
        Round(MedianAbsoluteDeviationScaleFactor(24), 4) = 1.4602 And _
        Round(MedianAbsoluteDeviationScaleFactor(26), 4) = 1.4619 And _
        Round(MedianAbsoluteDeviationScaleFactor(28), 4) = 1.4634 And _
        Round(MedianAbsoluteDeviationScaleFactor(30), 4) = 1.4826 And _
        Round(MedianAbsoluteDeviationScaleFactor(32), 4) = 1.4826 And _
        Round(MedianAbsoluteDeviationScaleFactor(34), 4) = 1.4826 And _
        Round(MedianAbsoluteDeviationScaleFactor(36), 4) = 1.4826 And _
        Round(MedianAbsoluteDeviationScaleFactor(38), 4) = 1.4826

      UnitTestsPassed = UnitTestsPassed And _
        Round(ControlLimitRejectionCriterion(0.99, 1), 4) = 63.6567 And _
        Round(ControlLimitRejectionCriterion(0.99, 2), 4) = 9.9248 And _
        Round(ControlLimitRejectionCriterion(0.99, 4), 4) = 4.6041 And _
        Round(ControlLimitRejectionCriterion(0.99, 6), 4) = 3.7074 And _
        Round(ControlLimitRejectionCriterion(0.99, 8), 4) = 3.3554 And _
        Round(ControlLimitRejectionCriterion(0.99, 10), 4) = 3.1693 And _
        Round(ControlLimitRejectionCriterion(0.99, 12), 4) = 3.0545 And _
        Round(ControlLimitRejectionCriterion(0.99, 14), 4) = 2.9768 And _
        Round(ControlLimitRejectionCriterion(0.99, 16), 4) = 2.9208 And _
        Round(ControlLimitRejectionCriterion(0.99, 20), 4) = 2.8453 And _
        Round(ControlLimitRejectionCriterion(0.99, 22), 4) = 2.8188 And _
        Round(ControlLimitRejectionCriterion(0.99, 24), 4) = 2.7969 And _
        Round(ControlLimitRejectionCriterion(0.99, 28), 4) = 2.7633 And _
        Round(ControlLimitRejectionCriterion(0.99, 30), 4) = 2.5758 And _
        Round(ControlLimitRejectionCriterion(0.99, 32), 4) = 2.5758 And _
        Round(ControlLimitRejectionCriterion(0.99, 34), 4) = 2.5758 And _
        Round(ControlLimitRejectionCriterion(0.95, 30), 4) = 1.96 And _
        Round(ControlLimitRejectionCriterion(0.90, 30), 4) = 1.6449 And _
        Round(ControlLimitRejectionCriterion(0.95, 25), 4) = 2.0595 And _
        Round(ControlLimitRejectionCriterion(0.90, 20), 4) = 1.7247

      UnitTestsPassed = UnitTestsPassed And _
        NonNaNCount({60}) = 1 And _
        NonNaNCount({60, 70}) = 2 And _
        NonNaNCount({60, 70, NaN}) = 2 And _
        NonNaNCount({60, 70, NaN, 20}) = 3 And _
        NonNaNCount({60, 70, NaN, 20, NaN}) = 3 And _
        NonNaNCount({60, 70, NaN, 20, NaN, NaN}) = 3 And _
        NonNaNCount({70, NaN, 20, NaN, NaN, 5}) = 3 And _
        NonNaNCount({NaN, 20, NaN, NaN, 5, 10}) = 3 And _
        NonNaNCount({20, NaN, NaN, 5, 10, 15}) = 4 And _
        NonNaNCount({NaN, NaN, 5, 10, 15, 20}) = 4 And _
        NonNaNCount({NaN, 5, 10, 15, 20, 25}) = 5 And _
        NonNaNCount({5, 10, 15, 20, 25, 30}) = 6 And _
        NonNaNCount({10, 15, 20, 25, 30}) = 5 And _
        NonNaNCount({15, 20, 25, 30}) = 4 And _
        NonNaNCount({20, 25, 30}) = 3 And _
        NonNaNCount({25, 30}) = 2 And _
        NonNaNCount({30}) = 1 And _
        NonNaNCount({}) = 0 And _
        NonNaNCount({NaN}) = 0 And _
        NonNaNCount({NaN, NaN}) = 0

      UnitTestsPassed = UnitTestsPassed And _
        Round(Mean({60}), 4) = 60 And _
        Round(Mean({72}), 4) = 72 And _
        Round(Mean({32, 95}), 4) = 63.5 And _
        Round(Mean({81, 75}), 4) = 78 And _
        Round(Mean({67, 76, 25}), 4) = 56 And _
        Round(Mean({25, NaN, 27}), 4) = 26 And _
        Round(Mean({31, 73, 83, 81}), 4) = 67 And _
        Round(Mean({18, 58, 47, 47}), 4) = 42.5 And _
        Round(Mean({10, NaN, 20, 30, NaN}), 4) = 20 And _
        Round(Mean({67, 90, 74, 32, 62}), 4) = 65 And _
        Round(Mean({78, 0, 98, 65, 69, 57}), 4) = 61.1667 And _
        Round(Mean({49, 35, 74, 25, 28, 92}), 4) = 50.5 And _
        Round(Mean({7, 64, 22, 7, 42, 34, 30}), 4) = 29.4286 And _
        Round(Mean({13, 29, 39, 33, 96, 43, 17}), 4) = 38.5714 And _
        Round(Mean({59, 78, 53, 7, 18, 44, 63, 40}), 4) = 45.25 And _
        Round(Mean({77, 71, 99, 39, 50, 94, 67, 30}), 4) = 65.875 And _
        Round(Mean({91, 69, 63, 5, 44, 93, 89, 45, 50}), 4) = 61 And _
        Round(Mean({12, 84, 12, 94, 52, 17, 1, 13, 37}), 4) = 35.7778 And _
        Round(Mean({18, 14, 54, 40, 73, 77, 4, 91, 53, 10}), 4) = 43.4 And _
        Round(Mean({80, 30, 1, 92, 44, 61, 18, 72, 63, 41}), 4) = 50.2

      UnitTestsPassed = UnitTestsPassed And _
        Median({57}) = 57 And _
        Median({46}) = 46 And _
        Median({79, 86}) = 82.5 And _
        Median({46, 45}) = 45.5 And _
        Median({10, NaN, 20}) = 15 And _
        Median({58, 79, 68}) = 68 And _
        Median({NaN, 10, 30, NaN}) = 20 And _
        Median({30, 10, NaN, 15}) = 15 And _
        IsNaN(Median({NaN, NaN, NaN, NaN, NaN})) And _
        Median({19, 3, 9, 14, 31}) = 14 And _
        Median({2, 85, 33, 10, 38, 56}) = 35.5 And _
        Median({42, 65, 57, 92, 56, 59}) = 58 And _
        Median({53, 17, 18, 73, 34, 96, 9}) = 34 And _
        Median({23, 23, 43, 74, 8, 51, 88}) = 43 And _
        Median({72, 46, 39, 7, 83, 96, 18, 50}) = 48 And _
        Median({50, 25, 28, 99, 79, 97, 30, 16}) = 40 And _
        Median({7, 22, 58, 32, 5, 90, 46, 91, 66}) = 46 And _
        Median({81, 64, 5, 23, 48, 18, 19, 87, 15}) = 23 And _
        Median({33, 82, 42, 33, 81, 56, 13, 13, 54, 6}) = 37.5 And _
        Median({55, 40, 75, 23, 53, 85, 59, 9, 72, 44}) = 54

      UnitTestsPassed = UnitTestsPassed And _
        Hash(AbsoluteDeviation({100, 44, 43, 45}, 44)) = _
        Hash({56, 0, 1, 1}) And _
        Hash(AbsoluteDeviation({76, 70, 84, 39}, 77)) = _
        Hash({1, 7, 7, 38}) And _
        Hash(AbsoluteDeviation({15, 74, 54, 23}, 93)) = _
        Hash({78, 19, 39, 70}) And _
        Hash(AbsoluteDeviation({48, 12, 85, 50}, 9)) = _
        Hash({39, 3, 76, 41}) And _
        Hash(AbsoluteDeviation({7, 24, 28, 41}, 69)) = _
        Hash({62, 45, 41, 28}) And _
        Hash(AbsoluteDeviation({55, 46, 81, 50}, 88)) = _
        Hash({33, 42, 7, 38}) And _
        Hash(AbsoluteDeviation({84, 14, 23, 54}, 52)) = _
        Hash({32, 38, 29, 2}) And _
        Hash(AbsoluteDeviation({92, 92, 45, 21}, 49)) = _
        Hash({43, 43, 4, 28}) And _
        Hash(AbsoluteDeviation({3, 9, 70, 51}, 36)) = _
        Hash({33, 27, 34, 15}) And _
        Hash(AbsoluteDeviation({14, 7, 8, 53}, 37)) = _
        Hash({23, 30, 29, 16}) And _
        Hash(AbsoluteDeviation({66, 37, 23, 80}, 3)) = _
        Hash({63, 34, 20, 77}) And _
        Hash(AbsoluteDeviation({93, 45, 18, 50}, 75)) = _
        Hash({18, 30, 57, 25}) And _
        Hash(AbsoluteDeviation({2, 73, 13, 0}, 37)) = _
        Hash({35, 36, 24, 37}) And _
        Hash(AbsoluteDeviation({76, 77, 56, 11}, 77)) = _
        Hash({1, 0, 21, 66}) And _
        Hash(AbsoluteDeviation({90, 36, 72, 54}, 81)) = _
        Hash({9, 45, 9, 27}) And _
        Hash(AbsoluteDeviation({5, 62, 50, 23}, 19)) = _
        Hash({14, 43, 31, 4}) And _
        Hash(AbsoluteDeviation({79, 51, 54, 24}, 60)) = _
        Hash({19, 9, 6, 36}) And _
        Hash(AbsoluteDeviation({26, 53, 38, 42}, 60)) = _
        Hash({34, 7, 22, 18}) And _
        Hash(AbsoluteDeviation({4, 100, 32, 23}, 32)) = _
        Hash({28, 68, 0, 9}) And _
        Hash(AbsoluteDeviation({15, 94, 58, 67}, 78)) = _
        Hash({63, 16, 20, 11})

      UnitTestsPassed = UnitTestsPassed And _
        Round(MeanAbsoluteDeviation _
        ({19}), 4) = 0 And _
        Round(MeanAbsoluteDeviation _
        ({86}), 4) = 0 And _
        Round(MeanAbsoluteDeviation _
        ({7, 24}), 4) = 8.5 And _
        Round(MeanAbsoluteDeviation _
        ({12, 96}), 4) = 42 And _
        Round(MeanAbsoluteDeviation _
        ({19, 74, 70}), 4) = 23.5556 And _
        Round(MeanAbsoluteDeviation _
        ({96, 93, 65}), 4) = 13.1111 And _
        Round(MeanAbsoluteDeviation _
        ({47, 29, 24, 11}), 4) = 10.25 And _
        Round(MeanAbsoluteDeviation _
        ({3, 43, 53, 80}), 4) = 21.75 And _
        Round(MeanAbsoluteDeviation _
        ({17, 27, 98, 85, 51}), 4) = 28.72 And _
        Round(MeanAbsoluteDeviation _
        ({2, 82, 63, 1, 49}), 4) = 30.32 And _
        Round(MeanAbsoluteDeviation _
        ({9, 25, 41, 85, 82, 55}), 4) = 24.5 And _
        Round(MeanAbsoluteDeviation _
        ({5, 74, 53, 97, 81, 21}), 4) = 28.8333 And _
        Round(MeanAbsoluteDeviation _
        ({26, 81, 9, 18, 39, 97, 21}), 4) = 27.102 And _
        Round(MeanAbsoluteDeviation _
        ({5, 83, 31, 24, 55, 22, 87}), 4) = 26.6939 And _
        Round(MeanAbsoluteDeviation _
        ({22, 84, 6, 79, 89, 71, 34, 56}), 4) = 25.8438 And _
        Round(MeanAbsoluteDeviation _
        ({33, 39, 6, 88, 69, 11, 76, 65}), 4) = 26.125 And _
        Round(MeanAbsoluteDeviation _
        ({31, 52, 12, 60, 52, 44, 47, 81, 34}), 4) = 13.9012 And _
        Round(MeanAbsoluteDeviation _
        ({64, 63, 54, 94, 25, 80, 97, 45, 51}), 4) = 17.8519 And _
        Round(MeanAbsoluteDeviation _
        ({47, 22, 52, 22, 10, 38, 94, 85, 54, 41}), 4) = 19.9 And _
        Round(MeanAbsoluteDeviation _
        ({7, 12, 84, 29, 41, 8, 18, 15, 16, 84}), 4) = 22.96

      UnitTestsPassed = UnitTestsPassed And _
        MedianAbsoluteDeviation _
        ({2}) = 0 And _
        MedianAbsoluteDeviation _
        ({37}) = 0 And _
        MedianAbsoluteDeviation _
        ({87, 37}) = 25 And _
        MedianAbsoluteDeviation _
        ({13, 74}) = 30.5 And _
        MedianAbsoluteDeviation _
        ({39, 52, 93}) = 13 And _
        MedianAbsoluteDeviation _
        ({90, 24, 47}) = 23 And _
        MedianAbsoluteDeviation _
        ({11, 51, 20, 62}) = 20 And _
        MedianAbsoluteDeviation _
        ({74, 35, 9, 95}) = 30 And _
        MedianAbsoluteDeviation _
        ({32, 46, 15, 90, 66}) = 20 And _
        MedianAbsoluteDeviation _
        ({91, 19, 50, 55, 44}) = 6 And _
        MedianAbsoluteDeviation _
        ({2, 64, 87, 65, 61, 97}) = 13 And _
        MedianAbsoluteDeviation _
        ({35, 66, 73, 74, 71, 93}) = 4 And _
        MedianAbsoluteDeviation _
        ({54, 81, 80, 36, 11, 36, 45}) = 9 And _
        MedianAbsoluteDeviation _
        ({14, 69, 40, 68, 75, 10, 69}) = 7 And _
        MedianAbsoluteDeviation _
        ({40, 51, 28, 21, 91, 95, 66, 3}) = 22.5 And _
        MedianAbsoluteDeviation _
        ({57, 87, 94, 46, 51, 27, 10, 7}) = 30 And _
        MedianAbsoluteDeviation _
        ({3, 89, 62, 84, 86, 37, 14, 72, 33}) = 25 And _
        MedianAbsoluteDeviation _
        ({48, 6, 14, 2, 74, 89, 15, 8, 83}) = 13 And _
        MedianAbsoluteDeviation _
        ({2, 22, 91, 84, 43, 96, 55, 3, 9, 11}) = 26.5 And _
        MedianAbsoluteDeviation _
        ({6, 96, 82, 26, 47, 84, 34, 39, 60, 99}) = 28

      UnitTestsPassed = UnitTestsPassed And _
        Not UseMeanAbsoluteDeviation({6, 5, 5, 9}) And _
        Not UseMeanAbsoluteDeviation({2, 2, 10, 9}) And _
        Not UseMeanAbsoluteDeviation({4, 0, 10, 6}) And _
        Not UseMeanAbsoluteDeviation({6, 10, 1, 1}) And _
        UseMeanAbsoluteDeviation({3, 0, 0, 0}) And _
        Not UseMeanAbsoluteDeviation({10, 0, 8, 5}) And _
        Not UseMeanAbsoluteDeviation({7, 4, 5, 3}) And _
        UseMeanAbsoluteDeviation({5, 1, 5, 5}) And _
        Not UseMeanAbsoluteDeviation({2, 7, 3, 8}) And _
        Not UseMeanAbsoluteDeviation({2, 6, 10, 1}) And _
        Not UseMeanAbsoluteDeviation({1, 6, 3, 5}) And _
        Not UseMeanAbsoluteDeviation({3, 9, 7, 3}) And _
        UseMeanAbsoluteDeviation({5, 5, 8, 5}) And _
        Not UseMeanAbsoluteDeviation({5, 10, 5, 4}) And _
        Not UseMeanAbsoluteDeviation({0, 2, 4, 1}) And _
        Not UseMeanAbsoluteDeviation({7, 3, 0, 10}) And _
        UseMeanAbsoluteDeviation({4, 4, 4, 0}) And _
        Not UseMeanAbsoluteDeviation({5, 7, 4, 5}) And _
        UseMeanAbsoluteDeviation({2, 2, 2, 2}) And _
        UseMeanAbsoluteDeviation({9, 4, 4, 4})

      UnitTestsPassed = UnitTestsPassed And _
        CentralTendency({3, 0, 0, 0}) = 0.75 And _
        CentralTendency({2, 10, 0, 1}) = 1.5 And _
        CentralTendency({7, 7, 7, 1}) = 5.5 And _
        CentralTendency({5, 7, 2, 8}) = 6 And _
        CentralTendency({9, 3, 4, 5}) = 4.5 And _
        CentralTendency({3, 3, 3, 3}) = 3 And _
        CentralTendency({8, 4, 2, 10}) = 6 And _
        CentralTendency({2, 1, 10, 10}) = 6 And _
        CentralTendency({3, 3, 6, 2}) = 3 And _
        CentralTendency({9, 9, 6, 5}) = 7.5 And _
        CentralTendency({2, 8, 8, 9}) = 8 And _
        CentralTendency({7, 7, 4, 1}) = 5.5 And _
        CentralTendency({5, 5, 5, 0}) = 3.75 And _
        CentralTendency({4, 2, 3, 7}) = 3.5 And _
        CentralTendency({2, 1, 5, 1}) = 1.5 And _
        CentralTendency({9, 4, 5, 0}) = 4.5 And _
        CentralTendency({1, 1, 7, 1}) = 2.5 And _
        CentralTendency({1, 5, 9, 5}) = 5 And _
        CentralTendency({3, 5, 1, 9}) = 4 And _
        CentralTendency({0, 0, 0, 0}) = 0

      UnitTestsPassed = UnitTestsPassed And _
        Round(ControlLimit({8, 2, 0, 10}, 0.99), 4) = 30.5449 And _
        Round(ControlLimit({2, 4, 8, 7}, 0.99), 4) = 15.2724 And _
        Round(ControlLimit({5, 8, 0, 2}, 0.99), 4) = 19.0906 And _
        Round(ControlLimit({8, 1, 0, 3}, 0.99), 4) = 11.4543 And _
        Round(ControlLimit({10, 7, 1, 3}, 0.99), 4) = 22.9087 And _
        Round(ControlLimit({6, 2, 1, 9}, 0.99), 4) = 19.0906 And _
        Round(ControlLimit({4, 9, 9, 3}, 0.99), 4) = 19.0906 And _
        Round(ControlLimit({10, 7, 2, 8}, 0.99), 4) = 11.4543 And _
        Round(ControlLimit({6, 0, 10, 1}, 0.99), 4) = 22.9087 And _
        Round(ControlLimit({10, 3, 4, 2}, 0.99), 4) = 7.6362 And _
        Round(ControlLimit({6, 4, 4, 4}, 0.99), 4) = 5.4904 And _
        Round(ControlLimit({1, 0, 9, 9}, 0.99), 4) = 30.5449 And _
        Round(ControlLimit({0, 3, 6, 2}, 0.99), 4) = 11.4543 And _
        Round(ControlLimit({9, 7, 4, 6}, 0.99), 4) = 11.4543 And _
        Round(ControlLimit({6, 6, 4, 1}, 0.99), 4) = 7.6362 And _
        Round(ControlLimit({7, 3, 4, 1}, 0.99), 4) = 11.4543 And _
        Round(ControlLimit({6, 4, 4, 10}, 0.99), 4) = 7.6362 And _
        Round(ControlLimit({10, 5, 5, 5}, 0.99), 4) = 13.7259 And _
        Round(ControlLimit({8, 5, 5, 5}, 0.98), 4) = 6.4023 And _
        Round(ControlLimit({8, 4, 0, 0}, 0.95), 4) = 8.3213

      UnitTestsPassed = UnitTestsPassed And _
        Hash(RemoveOutliers({100, 100, 100, 100, 100, 100, 100, 100, 999})) = _
        Hash({100, 100, 100, 100, 100, 100, 100, 100, NaN}) And _
        Hash(RemoveOutliers({100, 101, 102, 103, 104, 105, 106, 107, 999})) = _
        Hash({100, 101, 102, 103, 104, 105, 106, 107, NaN}) And _
        Hash(RemoveOutliers({2223.6946, 2770.1624, 2125.7544, 3948.9927, _
        2184.2341, 2238.6421, 2170.0227, 2967.0674, 2177.3738, 3617.1328, _
        2460.8193, 3315.8684})) = Hash({2223.6946, 2770.1624, 2125.7544, NaN, _
        2184.2341, 2238.6421, 2170.0227, 2967.0674, 2177.3738, NaN, _
        2460.8193, NaN}) And _
        Hash(RemoveOutliers({3355.1553, 3624.3154, 3317.6895, 3610.0039, _
        3990.751, 2950.4382, 2140.5908, 3237.4917, 3319.7139, 2829.2725, _
        3406.9199, 3230.0078})) = Hash({3355.1553, 3624.3154, 3317.6895, _
        3610.0039, 3990.751, 2950.4382, NaN, 3237.4917, 3319.7139, 2829.2725, _
        3406.9199, 3230.0078}) And _
        Hash(RemoveOutliers({2969.7808, 3899.0913, 2637.4045, 2718.73, _
        2960.9597, 2650.6521, 2707.4294, 2034.5339, 2935.9111, 3458.7085, _
        2584.53, 3999.4238})) = Hash({2969.7808, NaN, 2637.4045, 2718.73, _
        2960.9597, 2650.6521, 2707.4294, 2034.5339, 2935.9111, 3458.7085, _
        2584.53, NaN}) And _
        Hash(RemoveOutliers({2774.8018, 2755.0251, 2756.6152, 3800.0625, _
        2900.0671, 2784.0134, 3955.2947, 2847.0908, 2329.7837, 3282.4614, _
        2597.1582, 3009.8796})) = Hash({2774.8018, 2755.0251, 2756.6152, NaN, _
        2900.0671, 2784.0134, NaN, 2847.0908, 2329.7837, 3282.4614, _
        2597.1582, 3009.8796}) And _
        Hash(RemoveOutliers({3084.8821, 3394.1196, 3131.3245, 2799.9587, _
        2528.3088, 3015.4998, 2912.2029, 2022.2645, 3666.5674, 3685.1973, _
        3149.6931, 3070.0479})) = Hash({3084.8821, 3394.1196, 3131.3245, _
        2799.9587, 2528.3088, 3015.4998, 2912.2029, NaN, 3666.5674, _
        3685.1973, 3149.6931, 3070.0479}) And _
        Hash(RemoveOutliers({3815.72, 3063.9106, 3535.0366, 2349.564, _
        2597.2661, 3655.3076, 3452.7407, 2020.7682, 3810.7046, 3833.8396, _
        3960.6016, 3866.8149})) = Hash({3815.72, 3063.9106, 3535.0366, NaN, _
        2597.2661, 3655.3076, 3452.7407, NaN, 3810.7046, 3833.8396, _
        3960.6016, 3866.8149}) And _
        Hash(RemoveOutliers({2812.0613, 3726.7427, 2090.9749, 2548.4485, _
        3900.5151, 3545.854, 3880.2229, 3940.9585, 3942.2234, 3263.0137, _
        3701.8882, 2056.5291})) = Hash({2812.0613, 3726.7427, NaN, 2548.4485, _
        3900.5151, 3545.854, 3880.2229, 3940.9585, 3942.2234, 3263.0137, _
        3701.8882, NaN}) And _
        Hash(RemoveOutliers({3798.4775, 2959.3879, 2317.3547, 2596.3599, _
        2075.6292, 2563.9685, 2695.5081, 2386.2161, 2433.1106, 2810.3716, _
        2499.7554, 3843.103})) = Hash({NaN, 2959.3879, 2317.3547, 2596.3599, _
        2075.6292, 2563.9685, 2695.5081, 2386.2161, 2433.1106, 2810.3716, _
        2499.7554, NaN}) And _
        Hash(RemoveOutliers({2245.7856, 2012.4834, 2473.0103, 2684.5693, _
        2645.4729, 2851.019, 2344.6099, 2408.1492, 3959.5967, 3954.0583, _
        2399.2617, 2652.8855})) = Hash({2245.7856, 2012.4834, 2473.0103, _
        2684.5693, 2645.4729, 2851.019, 2344.6099, 2408.1492, NaN, NaN, _
        2399.2617, 2652.8855}) And _
        Hash(RemoveOutliers({2004.5355, 2743.0693, 3260.7441, 2382.8906, _
        2365.9385, 2243.333, 3506.5352, 3905.7717, 3516.5337, 2133.8328, _
        2308.1809, 2581.4009})) = Hash({2004.5355, 2743.0693, 3260.7441, _
        2382.8906, 2365.9385, 2243.333, 3506.5352, NaN, 3516.5337, 2133.8328, _
        2308.1809, 2581.4009}) And _
        Hash(RemoveOutliers({3250.5376, 3411.313, 2037.264, 3709.5815, _
        3417.1167, 3996.0493, 3529.637, 3992.7163, 2786.95, 3728.834, _
        3304.4272, 2248.9119})) = Hash({3250.5376, 3411.313, NaN, 3709.5815, _
        3417.1167, 3996.0493, 3529.637, 3992.7163, 2786.95, 3728.834, _
        3304.4272, 2248.9119}) And _
        Hash(RemoveOutliers({2398.3125, 2742.4028, 2720.752, 2628.8442, _
        2750.1482, 2724.4932, 2161.6875, 2644.4163, 2188.2952, 2455.4622, _
        3332.5503, 2540.5198})) = Hash({2398.3125, 2742.4028, 2720.752, _
        2628.8442, 2750.1482, 2724.4932, 2161.6875, 2644.4163, 2188.2952, _
        2455.4622, NaN, 2540.5198}) And _
        Hash(RemoveOutliers({3991.7854, 3607.98, 2686.032, 2546.969, _
        3053.8796, 3138.9824, 2441.1689, 2737.1245, 2616.7139, 2550.5774, _
        2406.0913, 2743.2361})) = Hash({NaN, 3607.98, 2686.032, 2546.969, _
        3053.8796, 3138.9824, 2441.1689, 2737.1245, 2616.7139, 2550.5774, _
        2406.0913, 2743.2361}) And _
        Hash(RemoveOutliers({2361.5334, 3636.4312, 2187.593, 2281.5432, _
        2132.3833, 2056.792, 2227.7795, 2757.1753, 3416.9126, 2568.927, _
        2094.2065, 3449.3984})) = Hash({2361.5334, NaN, 2187.593, 2281.5432, _
        2132.3833, 2056.792, 2227.7795, 2757.1753, NaN, 2568.927, 2094.2065, _
        NaN}) And _
        Hash(RemoveOutliers({2249.7119, 2411.8374, 3041.5498, 2679.1458, _
        2561.1577, 2405.7229, 2775.2253, 2832.1233, 2540.2134, 3654.5903, _
        3970.5173, 2920.5637})) = Hash({2249.7119, 2411.8374, 3041.5498, _
        2679.1458, 2561.1577, 2405.7229, 2775.2253, 2832.1233, 2540.2134, _
        3654.5903, NaN, 2920.5637}) And _
        Hash(RemoveOutliers({2038.1091, 2248.3057, 2427.1646, 2337.2427, _
        2642.043, 3497.5393, 3996.3579, 2178.979, 3968.8848, 3460.8613, _
        2774.8486, 2338.1362})) = Hash({2038.1091, 2248.3057, 2427.1646, _
        2337.2427, 2642.043, 3497.5393, NaN, 2178.979, NaN, 3460.8613, _
        2774.8486, 2338.1362}) And _
        Hash(RemoveOutliers({3010.9485, 2517.2876, 2057.7188, 2133.0801, _
        3192.0308, 2035.0759, 3821.248, 2391.8086, 2267.896, 3751.3276, _
        2340.9497, 2327.333})) = Hash({3010.9485, 2517.2876, 2057.7188, _
        2133.0801, 3192.0308, 2035.0759, NaN, 2391.8086, 2267.896, NaN, _
        2340.9497, 2327.333}) And _
        Hash(RemoveOutliers({NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN, _
        999})) = Hash({NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN, NaN})

      UnitTestsPassed = UnitTestsPassed And _
        Round(ExponentialMovingAverage _
        ({70.5547}), 4) = 70.5547 And _
        Round(ExponentialMovingAverage _
        ({53.3424, 57.9519}), 4) = 56.7995 And _
        Round(ExponentialMovingAverage _
        ({28.9562, 30.1948}), 4) = 29.8852 And _
        Round(ExponentialMovingAverage _
        ({77.474, 1.4018}), 4) = 20.4199 And _
        Round(ExponentialMovingAverage _
        ({76.0724, 81.449, 70.9038}), 4) = 74.6551 And _
        Round(ExponentialMovingAverage _
        ({4.5353, 41.4033, 86.2619}), 4) = 61.7699 And _
        Round(ExponentialMovingAverage _
        ({79.048, 37.3536, 96.1953}), 4) = 76.9338 And _
        Round(ExponentialMovingAverage _
        ({87.1446, 5.6237, 94.9557}), 4) = 68.3164 And _
        Round(ExponentialMovingAverage _
        ({36.4019, 52.4868, 76.7112}), 4) = 64.0315 And _
        Round(ExponentialMovingAverage _
        ({5.3505, 59.2458, 46.87, 29.8165}), 4) = 36.959 And _
        Round(ExponentialMovingAverage _
        ({62.2697, 64.7821, 26.3793, 27.9342}), 4) = 37.0099 And _
        Round(ExponentialMovingAverage _
        ({82.9802, 82.4602, 58.9163, 98.6093}), 4) = 83.4414 And _
        Round(ExponentialMovingAverage _
        ({91.0964, 22.6866, 69.5116, 98.0003, 24.3931}), 4) = 55.7929 And _
        Round(ExponentialMovingAverage _
        ({53.3873, 10.637, 99.9415, 67.6176, 1.5704}), 4) = 40.2177 And _
        Round(ExponentialMovingAverage _
        ({57.5184, 10.0052, 10.3023, 79.8884, 28.448}), 4) = 38.6235 And _
        Round(ExponentialMovingAverage _
        ({4.5649, 29.5773, 38.2011, 30.097, 94.8571}), 4) = 54.345 And _
        Round(ExponentialMovingAverage _
        ({97.9829, 40.1374, 27.828, 16.0442, 16.2822}), 4) = 27.0999 And _
        Round(ExponentialMovingAverage({64.6587, _
        41.0073, 41.2767, 71.273, 32.6206, 63.3179}), 4) = 52.9531 And _
        Round(ExponentialMovingAverage({20.7561, _
        18.6014, 58.3359, 8.0715, 45.7971, 90.573}), 4) = 51.847 And _
        Round(ExponentialMovingAverage _
        ({26.1368, 78.5212, 37.8903, 28.9665, 91.9377, 63.1742}), 4) = 60.2045

      UnitTestsPassed = UnitTestsPassed And _
        Round(SlopeToAngle(-4.5806), 4) = -77.6849 And _
        Round(SlopeToAngle(-4.2541), 4) = -76.7718 And _
        Round(SlopeToAngle(1.7964), 4) = 60.8967 And _
        Round(SlopeToAngle(-3.2474), 4) = -72.8844 And _
        Round(SlopeToAngle(4.7917), 4) = 78.2119 And _
        Round(SlopeToAngle(2.1792), 4) = 65.3504 And _
        Round(SlopeToAngle(0.4736), 4) = 25.3422 And _
        Round(SlopeToAngle(-2.0963), 4) = -64.4974 And _
        Round(SlopeToAngle(-3.2077), 4) = -72.6851 And _
        Round(SlopeToAngle(-1.5425), 4) = -57.0447 And _
        Round(SlopeToAngle(-0.5587), 4) = -29.1921 And _
        Round(SlopeToAngle(1.2829), 4) = 52.0642 And _
        Round(SlopeToAngle(3.9501), 4) = 75.7936 And _
        Round(SlopeToAngle(2.5841), 4) = 68.8445 And _
        Round(SlopeToAngle(-3.4547), 4) = -73.8563 And _
        Round(SlopeToAngle(3.2931), 4) = 73.1083 And _
        Round(SlopeToAngle(3.2042), 4) = 72.6674 And _
        Round(SlopeToAngle(3.1088), 4) = 72.1687 And _
        Round(SlopeToAngle(-1.6831), 4) = -59.2837 And _
        Round(SlopeToAngle(-2.0031), 4) = -63.4704

      UnitTestsPassed = UnitTestsPassed And _
        AlignTimestamp(New DateTime(2016, 4, 4, 16, 33, 2), 60) = _
        New DateTime(2016, 4, 4, 16, 33, 0) And _
        AlignTimestamp(New DateTime(2015, 7, 15, 2, 29, 58), 60) =  _
        New DateTime(2015, 7, 15, 2, 29, 0) And _
        AlignTimestamp(New DateTime(2016, 4, 1, 22, 5, 17), 60) =  _
        New DateTime(2016, 4, 1, 22, 5, 0) And _
        AlignTimestamp(New DateTime(2013, 12, 1, 21, 47, 35), 60) =  _
        New DateTime(2013, 12, 1, 21, 47, 0) And _
        AlignTimestamp(New DateTime(2016, 11, 22, 0, 22, 17), 60) =  _
        New DateTime(2016, 11, 22, 0, 22, 0) And _
        AlignTimestamp(New DateTime(2016, 10, 11, 19, 11, 41), 300) =  _
        New DateTime(2016, 10, 11, 19, 10, 0) And _
        AlignTimestamp(New DateTime(2013, 10, 26, 4, 24, 53), 300) =  _
        New DateTime(2013, 10, 26, 4, 20, 0) And _
        AlignTimestamp(New DateTime(2014, 5, 2, 2, 52, 41), 300) =  _
        New DateTime(2014, 5, 2, 2, 50, 0) And _
        AlignTimestamp(New DateTime(2014, 8, 16, 13, 11, 10), 300) =  _
        New DateTime(2014, 8, 16, 13, 10, 0) And _
        AlignTimestamp(New DateTime(2014, 10, 25, 8, 26, 4), 300) =  _
        New DateTime(2014, 10, 25, 8, 25, 0) And _
        AlignTimestamp(New DateTime(2015, 6, 2, 18, 36, 24), 3600) =  _
        New DateTime(2015, 6, 2, 18, 0, 0) And _
        AlignTimestamp(New DateTime(2016, 11, 21, 16, 24, 27), 3600) =  _
        New DateTime(2016, 11, 21, 16, 0, 0) And _
        AlignTimestamp(New DateTime(2014, 4, 4, 8, 42, 10), 3600) =  _
        New DateTime(2014, 4, 4, 8, 0, 0) And _
        AlignTimestamp(New DateTime(2016, 2, 22, 19, 8, 41), 3600) =  _
        New DateTime(2016, 2, 22, 19, 0, 0) And _
        AlignTimestamp(New DateTime(2015, 9, 13, 22, 48, 17), 3600) =  _
        New DateTime(2015, 9, 13, 22, 0, 0) And _
        AlignTimestamp(New DateTime(2016, 10, 20, 2, 47, 48), 86400) =  _
        New DateTime(2016, 10, 20, 0, 0, 0) And _
        AlignTimestamp(New DateTime(2014, 2, 8, 23, 12, 34), 86400) =  _
        New DateTime(2014, 2, 8, 0, 0, 0) And _
        AlignTimestamp(New DateTime(2016, 2, 27, 23, 40, 39), 86400) =  _
        New DateTime(2016, 2, 27, 0, 0, 0) And _
        AlignTimestamp(New DateTime(2015, 8, 26, 9, 35, 55), 86400) =  _
        New DateTime(2015, 8, 26, 0, 0, 0) And _
        AlignTimestamp(New DateTime(2016, 2, 11, 0, 44, 7), 86400) =  _
        New DateTime(2016, 2, 11, 0, 0, 0)

      For i = 0 To 19
        If i = 0 Then
          StatisticsData = Statistics({3411, 3067, 3159, 2579, 2604, 3549, _
            2028, 3521, 3629, 3418, 2091, 2828})
        ElseIf i = 1 Then
          StatisticsData = Statistics({3725, 3581, 2747, 3924, 3743, 2112, _
            3899, 2728, 3050, 3534, 2107, 3185})
        ElseIf i = 2 Then
          StatisticsData = Statistics({2937, 2596, 3245, 3296, 2528, 2559, _
            3660, 3649, 3178, 3972, 3822, 2454})
        ElseIf i = 3 Then
          StatisticsData = Statistics({3390, 3960, 2488, 3068, 2213, 3999, _
            3352, 2031, 3150, 2200, 2206, 3598})
        ElseIf i = 4 Then
          StatisticsData = Statistics({2569, 2091, 2592, 2764, 2602, 3897, _
            3960, 2803, 2557, 2321, 2326, 3293})
        ElseIf i = 5 Then
          StatisticsData = Statistics({2820, 2826, 3425, 2652, 3266, 2415, _
            2372, 3167, 2161, 2916, 3811, 2523})
        ElseIf i = 6 Then
          StatisticsData = Statistics({3570, 2758, 2579, 3839, 3263, 3255, _
            2857, 2196, 3122, 3389, 3827, 3670})
        ElseIf i = 7 Then
          StatisticsData = Statistics({2045, 3087, 3832, 2861, 3356, 3005, _
            3027, 2926, 2707, 2810, 2539, 2111})
        ElseIf i = 8 Then
          StatisticsData = Statistics({2488, 3958, 2122, 2781, 2730, 2980, _
            2311, 2949, 2515, 3258, 3084, 2313})
        ElseIf i = 9 Then
          StatisticsData = Statistics({3877, 3309, 3012, 2781, 2215, 3568, _
            2919, 3507, 3192, 3665, 2038, 2421})
        ElseIf i = 10 Then
          StatisticsData = Statistics({2148, 2211, 2663, 2256, 2000, 3074, _
            3314, 3088, 3655, 2164, 2384, 3358}, {3, 5, 10, 20, 28, 32, 41, _
            46, 57, 66, 74, 76})
        ElseIf i = 11 Then
          StatisticsData = Statistics({2908, 2714, 2300, 3409, 3858, 3060, _
            2179, 3515, 2804, 2924, 2984, 2415}, {10, 18, 28, 35, 44, 50, 51, _
            62, 63, 66, 74, 80})
        ElseIf i = 12 Then
          StatisticsData = Statistics({2659, 2191, 3180, 2340, 3855, 2196, _
            2888, 2546, 3745, 3501, 2546, 3347}, {9, 11, 13, 18, 21, 26, 28, _
            37, 42, 47, 56, 61})
        ElseIf i = 13 Then
          StatisticsData = Statistics({2513, 2180, 2062, 2645, 3580, 2595, _
            2471, 2961, 2509, 2681, 2090, 2965}, {6, 14, 23, 26, 32, 35, 37, _
            41, 45, 48, 58, 68})
        ElseIf i = 14 Then
          StatisticsData = Statistics({2412, 3729, 3177, 3510, 3856, 2662, _
            3086, 2161, 3269, 2820, 3921, 2229}, {7, 12, 23, 29, 30, 32, 36, _
            38, 41, 49, 55, 61})
        ElseIf i = 15 Then
          StatisticsData = Statistics({3847, 3240, 2695, 2298, 2960, 2439, _
            3987, 2261, 2058, 2691, 3095, 3846}, {4, 9, 15, 18, 20, 26, 36, _
            45, 49, 58, 64, 71})
        ElseIf i = 16 Then
          StatisticsData = Statistics({3076, 2813, 3694, 3652, 3345, 3444, _
            3994, 2680, 2990, 2826, 3391, 2358}, {7, 14, 24, 28, 37, 40, 49, _
            51, 55, 61, 69, 77})
        ElseIf i = 17 Then
          StatisticsData = Statistics({2846, 3086, 3629, 3082, 2855, 3018, _
            2456, 3238, 2980, 3362, 3773, 2741}, {6, 16, 23, 29, 35, 40, 49, _
            60, 64, 73, 75, 78})
        ElseIf i = 18 Then
          StatisticsData = Statistics({2605, 2586, 2301, 3060, 2447, 3169, _
            2727, 3752, 2956, 2381, 3368, 3495}, {6, 13, 24, 30, 38, 47, 57, _
            59, 69, 77, 86, 96})
        ElseIf i = 19 Then
          StatisticsData = Statistics({3228, 3564, 2323, 3616, 2405, 3914, _
            2132, 2123, 3586, 2759, 2927, 2239}, {10, 15, 21, 22, 24, 32, 34, _
            43, 46, 53, 55, 63})
        End If
        With StatisticsData
          UnitTestsPassed = UnitTestsPassed And _
            Round(.Slope, 4) = {-24.1399, -67.5699, 51.3427, -56.9825, _
            27.3182, -2.6573, 32.1923, -46.8462, -11.1224, -61.5455, 9.4424, _
            -0.1602, 10.7659, 4.8889, -4.6572, -0.4548, -6.6652, 2.2442, _
            8.6539, -12.0462}(i) And _
            Round(.OriginSlope, 4) = {383.2213, 397.5889, 426.4229, 371.4506, _
            374.8399, 372.6621, 425.6739, 359.6522, 360.8676, 379.3893, _
            52.1812, 50.6373, 75.2409, 60.0661, 73.5505, 61.2832, 59.1633, _
            53.9967, 46.4227, 66.0541}(i) And _
            Round(.Angle, 4) = {-87.6279, -89.1521, 88.8842, -88.9946, _
            87.9036, -69.3779, 88.2208, -88.7771, -84.8624, -89.0691, 83.9546, _
            -9.0998, 84.6933, 78.4399, -77.8813, -24.4582, -81.4673, 65.9827, _
            83.4085, -85.2545}(i) And _
            Round(.OriginAngle, 4) = {89.8505, 89.8559, 89.8656, 89.8458, _
            89.8471, 89.8463, 89.8654, 89.8407, 89.8412, 89.849, 88.9021, _
            88.8687, 89.2385, 89.0462, 89.221, 89.0651, 89.0317, 88.939, _
            88.766, 89.1327}(i) And _
            Round(.Intercept, 4) = {3123.1026, 3566.2179, 2875.6154, _
            3284.6538, 2664.3333, 2877.4487, 3016.6923, 3116.4872, 2851.9231, _
            3380.5, 2332.5305, 2930.2549, 2585.1154, 2427.9251, 3229.618, _
            2967.1468, 3472.9635, 2986.3475, 2469.7771, 3320.9411}(i) And _
            Round(.StandardError, 4) = {582.4218, 633.706, 535.0359, 720.9024, _
            619.3358, 506.8629, 525.9328, 483.3154, 527.9273, 573.7699, _
            544.1683, 523.5486, 590.2042, 436.8994, 644.5969, 698.2517, _
            478.875, 384.1275, 419.8051, 657.4656}(i) And _
            Round(.Correlation, 4) = {-0.1548, -0.374, 0.3411, -0.2864, _
            0.1645, -0.0198, 0.2255, -0.3441, -0.0794, -0.3759, 0.4296, _
            -0.0071, 0.3208, 0.2029, -0.1209, -0.0154, -0.3016, 0.1484, _
            0.5294, -0.3121}(i) And _
            Round(.ModifiedCorrelation, 4) = {0.819, 0.7932, 0.8652, 0.7909, _
            0.8474, 0.8345, 0.8554, 0.8061, 0.8274, 0.7962, 0.8665, 0.9037, _
            0.8892, 0.908, 0.887, 0.8274, 0.8714, 0.8915, 0.9047, _
            0.8561}(i) And _
            Round(.Determination, 4) = {0.024, 0.1398, 0.1164, 0.082, 0.0271, _
            0.0004, 0.0509, 0.1184, 0.0063, 0.1413, 0.1845, 0.0001, 0.1029, _
            0.0412, 0.0146, 0.0002, 0.091, 0.022, 0.2803, 0.0974}(i)
        End With
      Next i

      Return UnitTestsPassed

    End Function


    Public Class DBMPointDriverWaterUsageModel
      Inherits DBMPointDriverAbstract


      Public Sub New(Point As Object)

        MyBase.New(Point)

      End Sub


      Public Overrides Function GetData(Timestamp As DateTime) As Double

        ' Model based on hourly water usage in Leeuwarden 2016.
        ' Calculated using polynomial regressions based on hourly (quintic),
        ' daily (cubic) and monthly (quartic) periodicity.

        If TypeOf Point Is Integer Then ' Point contains offset in hours
          Timestamp = Timestamp.AddHours(DirectCast(Point, Integer))
        End If

        With Timestamp
          Return 790*(-0.00012*.Month^4+0.0035*.Month^3-0.032*.Month^2+0.1* _
            .Month+0.93)*(0.000917*.DayOfWeek^3-0.0155*.DayOfWeek^2+0.0628* _
            .DayOfWeek+0.956)*(-0.00001221*(.Hour+.Minute/60)^5+0.0007805* _
            (.Hour+.Minute/60)^4-0.01796*(.Hour+.Minute/60)^3+0.1709*(.Hour+ _
            .Minute/60)^2-0.5032*(.Hour+.Minute/60)+0.7023)
        End With

      End Function


    End Class


    Public Shared Function IntegrationTestsPassed As Boolean

      ' Integration tests, returns True if all tests pass.

      Dim InputPointDriver As DBMPointDriverWaterUsageModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Timestamp As DateTime
      Dim i As Integer
      Dim Result As DBMResult
      Dim _DBM As New DBM

      IntegrationTestsPassed = True

      InputPointDriver = New DBMPointDriverWaterUsageModel(0)
      CorrelationPoints.Add _
        (New DBMCorrelationPoint(New DBMPointDriverWaterUsageModel(490), False))
      Timestamp = New DateTime(2016, 1, 1, 0, 0, 0)

      For i = 0 to 19
        Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)
        With Result
          IntegrationTestsPassed = IntegrationTestsPassed And _
            Round(.Factor, 4) = {-1.7572, 0, 0, 0, 0, 0, -11.8493, -22.9119, _
            0, 0, 0, 0, 1.1375, 0, 0, 0, 0, 0, 0, 0}(i) And _
            Round(.PredictionData.MeasuredValue, 4) = {527.5796, 687.0052, _
            1097.1504, 950.9752, 496.1124, 673.6569, 1139.1957, 867.4313, _
            504.9407, 656.4434, 1065.7651, 898.9191, 471.2433, 668.1, _
            1103.9689, 897.7268, 525.3563, 676.7206, 1183.0887, _
            975.8324}(i) And _
            Round(.PredictionData.PredictedValue, 4) = {548.7285, 716.9982, _
            1059.0551, 919.4719, 488.6181, 683.6728, 1155.5986, 895.6872, _
            503.2566, 655.7115, 1061.2282, 893.3488, 464.4957, 666.2928, _
            1084.1527, 901.6546, 523.8671, 666.1729, 1190.3511, _
            975.4264}(i) And _
            Round(.PredictionData.LowerControlLimit, 4) = {536.6932, 651.1959, _
            971.9833, 841.23, 445.6438, 682.8191, 1154.2143, 894.454, _
            494.6574, 644.6202, 1044.558, 879.175, 458.564, 659.4903, _
            1073.6794, 887.693, 511.2267, 651.6276, 1159.4964, _
            951.2698}(i) And _
            Round(.PredictionData.UpperControlLimit, 4) = {560.7638, 782.8006, _
            1146.127, 997.7138, 531.5924, 684.5266, 1156.9829, 896.9205, _
            511.8557, 666.8028, 1077.8984, 907.5226, 470.4274, 673.0952, _
            1094.626, 915.6163, 536.5076, 680.7182, 1221.2057, 999.583}(i)
        End With
        Timestamp = Timestamp.AddSeconds(365*24*60*60/20)
      Next i

      Return IntegrationTestsPassed

    End Function


    Public Shared Function PerformanceIndex As Double

      ' Returns the performance of the DBM calculation as a performance index.
      ' The returned value indicates how many full days per second this system
      ' can calculate when performing real-time continuous calculations.

      Const DurationTicks As Double = 0.1*TicksPerSecond ' 0.1 seconds

      Dim InputPointDriver As DBMPointDriverWaterUsageModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Timestamp, Timer As DateTime
      Dim Result As DBMResult
      Dim _DBM As New DBM
      Dim i, Count As Integer

      InputPointDriver = New DBMPointDriverWaterUsageModel(0)
      CorrelationPoints.Add(New DBMCorrelationPoint( _
        New DBMPointDriverWaterUsageModel(5394), False))
      CorrelationPoints.Add(New DBMCorrelationPoint( _
        New DBMPointDriverWaterUsageModel(227), True))
      Timestamp = New DateTime(2016, 1, 1, 0, 0, 0)

      ' Pre-fill cache for the DBMPoint to calculate a more realistic value for
      ' the performance index as this then better simulates a real-time
      ' continuous calculation.
      Timestamp = Timestamp. _
        AddSeconds(PredictionsCacheSize*-CalculationInterval)
      For i = 1 To PredictionsCacheSize
        Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)
        Timestamp = Timestamp.AddSeconds(CalculationInterval)
      Next i

      Timer = Now
      Do While Now.Ticks-Timer.Ticks < DurationTicks
        Result = _DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)
        Count += 1
        Timestamp = Timestamp.AddSeconds(CalculationInterval)
      Loop

      Return Count/(DurationTicks/TicksPerSecond)/(24*60*60/CalculationInterval)

    End Function


  End Class


End Namespace
