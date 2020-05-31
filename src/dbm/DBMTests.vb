Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
Imports System.Globalization
Imports System.Globalization.CultureInfo
Imports System.Math
Imports System.String
Imports System.TimeSpan
Imports Vitens.DynamicBandwidthMonitor.DBM
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMInfo
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMMisc
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTests


    Private Shared Function Hash(Values() As Double,
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
      Dim StatisticsItem As New DBMStatisticsItem

      UnitTestsPassed = True

      UnitTestsPassed = UnitTestsPassed And
        Round(Hash({6, 4, 7, 1, 1, 4, 2, 4}), 4) = 2.2326 And
        Round(Hash({8, 4, 7, 3, 2, 6, 5, 7}), 4) = 3.6609 And
        Round(Hash({1, 5, 4, 7, 7, 8, 5, 1}), 4) = 1.8084

      UnitTestsPassed = UnitTestsPassed And
        Not IsNullOrEmpty(DBMInfo.Version) And
        Not IsNullOrEmpty(LicenseNotice)

      UnitTestsPassed = UnitTestsPassed And
        HasCorrelation(0.9, 45) = True And
        HasCorrelation(0.85, 45) = True And
        HasCorrelation(0.8, 45) = False And
        HasCorrelation(0.9, 0) = False And
        HasCorrelation(0.9, 35) = True And
        HasCorrelation(0.9, 25) = False And
        HasCorrelation(0.8, 35) = False And
        HasCorrelation(-0.9, 45) = False And
        HasCorrelation(0.9, -45) = False And
        HasCorrelation(0, 0) = False

      UnitTestsPassed = UnitTestsPassed And
        HasAnticorrelation(-0.9, -45, False) = True And
        HasAnticorrelation(-0.85, -45, True) = False And
        HasAnticorrelation(-0.8, -45, False) = False And
        HasAnticorrelation(-0.9, 0, False) = False And
        HasAnticorrelation(-0.9, -35, False) = True And
        HasAnticorrelation(-0.9, -25, False) = False And
        HasAnticorrelation(-0.8, -35, False) = False And
        HasAnticorrelation(0.9, -45, False) = False And
        HasAnticorrelation(-0.9, 45, False) = False And
        HasAnticorrelation(0, 0, False) = False

      UnitTestsPassed = UnitTestsPassed And
        Suppress(5, 0, 0, 0, 0, False) = 5 And
        Suppress(5, -0.8, 0, 0, 0, False) = 5 And
        Suppress(5, -0.9, -45, 0, 0, False) = 0 And
        Suppress(5, -0.9, 0, 0, 0, True) = 5 And
        Suppress(5, 0, 0, 0.8, 0, False) = 5 And
        Suppress(5, 0, 0, 0.9, 45, False) = 0 And
        Suppress(5, 0, 0, 0.9, 45, True) = 0 And
        Suppress(-0.9, -0.85, 0, 0, 0, False) = -0.9 And
        Suppress(-0.9, -0.95, -45, 0, 0, False) = 0 And
        Suppress(0.9, 0, 0, 0.85, 45, False) = 0 And
        Suppress(0.9, 0, 0, 0.95, 1, True) = 0.9 And
        Suppress(-0.9, -0.99, -45, 0, 0, False) = 0 And
        Suppress(-0.9, 0.99, 0, 0, 0, False) = -0.9 And
        Suppress(-0.9, 0.99, 0, 0, 0, True) = -0.9 And
        Suppress(0.99, -0.9, -45, 0, 0, False) = 0 And
        Suppress(0.99, -0.9, 0, 0, 0, True) = 0.99 And
        Suppress(-0.99, 0, 0, 0, 0, True) = -0.99 And
        Suppress(0.99, 0, 0, 0, 0, True) = 0.99 And
        Suppress(-0.98, -0.99, -1, 0, 0, False) = -0.98 And
        Suppress(-0.98, -0.99, -45, 0, 0, False) = 0

      UnitTestsPassed = UnitTestsPassed And
        Round(NormSInv(0.95), 4) = 1.6449 And
        Round(NormSInv(0.99), 4) = 2.3263 And
        Round(NormSInv(0.9999), 4) = 3.719 And
        Round(NormSInv(0.8974), 4) = 1.2669 And
        Round(NormSInv(0.7663), 4) = 0.7267 And
        Round(NormSInv(0.2248), 4) = -0.7561 And
        Round(NormSInv(0.9372), 4) = 1.5317 And
        Round(NormSInv(0.4135), 4) = -0.2186 And
        Round(NormSInv(0.2454), 4) = -0.689 And
        Round(NormSInv(0.2711), 4) = -0.6095 And
        Round(NormSInv(0.2287), 4) = -0.7431 And
        Round(NormSInv(0.6517), 4) = 0.3899 And
        Round(NormSInv(0.8663), 4) = 1.1091 And
        Round(NormSInv(0.9275), 4) = 1.4574 And
        Round(NormSInv(0.7089), 4) = 0.5502 And
        Round(NormSInv(0.1234), 4) = -1.1582 And
        Round(NormSInv(0.0837), 4) = -1.3806 And
        Round(NormSInv(0.6243), 4) = 0.3168 And
        Round(NormSInv(0.0353), 4) = -1.808 And
        Round(NormSInv(0.9767), 4) = 1.9899

      UnitTestsPassed = UnitTestsPassed And
        Round(TInv2T(0.3353, 16), 4) = 0.9934 And
        Round(TInv2T(0.4792, 12), 4) = 0.7303 And
        Round(TInv2T(0.4384, 9), 4) = 0.8108 And
        Round(TInv2T(0.0905, 6), 4) = 2.0152 And
        Round(TInv2T(0.63, 16), 4) = 0.4911 And
        Round(TInv2T(0.1533, 11), 4) = 1.5339 And
        Round(TInv2T(0.6297, 12), 4) = 0.4948 And
        Round(TInv2T(0.1512, 4), 4) = 1.7714 And
        Round(TInv2T(0.4407, 18), 4) = 0.7884 And
        Round(TInv2T(0.6169, 15), 4) = 0.5108 And
        Round(TInv2T(0.6077, 18), 4) = 0.5225 And
        Round(TInv2T(0.4076, 20), 4) = 0.8459 And
        Round(TInv2T(0.1462, 18), 4) = 1.5187 And
        Round(TInv2T(0.3421, 6), 4) = 1.0315 And
        Round(TInv2T(0.6566, 6), 4) = 0.4676 And
        Round(TInv2T(0.2986, 1), 4) = 1.9733 And
        Round(TInv2T(0.2047, 14), 4) = 1.3303 And
        Round(TInv2T(0.5546, 2), 4) = 0.7035 And
        Round(TInv2T(0.0862, 6), 4) = 2.0504 And
        Round(TInv2T(0.6041, 10), 4) = 0.5354

      UnitTestsPassed = UnitTestsPassed And
        Round(TInv(0.4097, 8), 4) = -0.2359 And
        Round(TInv(0.174, 19), 4) = -0.9623 And
        Round(TInv(0.6545, 15), 4) = 0.4053 And
        Round(TInv(0.7876, 5), 4) = 0.8686 And
        Round(TInv(0.2995, 3), 4) = -0.5861 And
        Round(TInv(0.0184, 2), 4) = -5.0679 And
        Round(TInv(0.892, 1), 4) = 2.8333 And
        Round(TInv(0.7058, 18), 4) = 0.551 And
        Round(TInv(0.3783, 2), 4) = -0.3549 And
        Round(TInv(0.2774, 15), 4) = -0.6041 And
        Round(TInv(0.0406, 8), 4) = -1.9945 And
        Round(TInv(0.1271, 4), 4) = -1.3303 And
        Round(TInv(0.241, 18), 4) = -0.718 And
        Round(TInv(0.0035, 1), 4) = -90.942 And
        Round(TInv(0.1646, 10), 4) = -1.0257 And
        Round(TInv(0.279, 11), 4) = -0.6041 And
        Round(TInv(0.8897, 4), 4) = 1.4502 And
        Round(TInv(0.5809, 13), 4) = 0.2083 And
        Round(TInv(0.3776, 11), 4) = -0.3197 And
        Round(TInv(0.5267, 15), 4) = 0.0681

      UnitTestsPassed = UnitTestsPassed And
        Round(MeanAbsoluteDeviationScaleFactor, 4) = 1.2533

      UnitTestsPassed = UnitTestsPassed And
        Round(MedianAbsoluteDeviationScaleFactor(1), 4) = 1 And
        Round(MedianAbsoluteDeviationScaleFactor(2), 4) = 1.2247 And
        Round(MedianAbsoluteDeviationScaleFactor(4), 4) = 1.3501 And
        Round(MedianAbsoluteDeviationScaleFactor(6), 4) = 1.3936 And
        Round(MedianAbsoluteDeviationScaleFactor(8), 4) = 1.4157 And
        Round(MedianAbsoluteDeviationScaleFactor(10), 4) = 1.429 And
        Round(MedianAbsoluteDeviationScaleFactor(12), 4) = 1.4378 And
        Round(MedianAbsoluteDeviationScaleFactor(14), 4) = 1.4442 And
        Round(MedianAbsoluteDeviationScaleFactor(16), 4) = 1.449 And
        Round(MedianAbsoluteDeviationScaleFactor(18), 4) = 1.4527 And
        Round(MedianAbsoluteDeviationScaleFactor(20), 4) = 1.4557 And
        Round(MedianAbsoluteDeviationScaleFactor(22), 4) = 1.4581 And
        Round(MedianAbsoluteDeviationScaleFactor(24), 4) = 1.4602 And
        Round(MedianAbsoluteDeviationScaleFactor(26), 4) = 1.4619 And
        Round(MedianAbsoluteDeviationScaleFactor(28), 4) = 1.4634 And
        Round(MedianAbsoluteDeviationScaleFactor(30), 4) = 1.4826 And
        Round(MedianAbsoluteDeviationScaleFactor(32), 4) = 1.4826 And
        Round(MedianAbsoluteDeviationScaleFactor(34), 4) = 1.4826 And
        Round(MedianAbsoluteDeviationScaleFactor(36), 4) = 1.4826 And
        Round(MedianAbsoluteDeviationScaleFactor(38), 4) = 1.4826

      UnitTestsPassed = UnitTestsPassed And
        Round(ControlLimitRejectionCriterion(0.99, 1), 4) = 63.6567 And
        Round(ControlLimitRejectionCriterion(0.99, 2), 4) = 9.9248 And
        Round(ControlLimitRejectionCriterion(0.99, 4), 4) = 4.6041 And
        Round(ControlLimitRejectionCriterion(0.99, 6), 4) = 3.7074 And
        Round(ControlLimitRejectionCriterion(0.99, 8), 4) = 3.3554 And
        Round(ControlLimitRejectionCriterion(0.99, 10), 4) = 3.1693 And
        Round(ControlLimitRejectionCriterion(0.99, 12), 4) = 3.0545 And
        Round(ControlLimitRejectionCriterion(0.99, 14), 4) = 2.9768 And
        Round(ControlLimitRejectionCriterion(0.99, 16), 4) = 2.9208 And
        Round(ControlLimitRejectionCriterion(0.99, 20), 4) = 2.8453 And
        Round(ControlLimitRejectionCriterion(0.99, 22), 4) = 2.8188 And
        Round(ControlLimitRejectionCriterion(0.99, 24), 4) = 2.7969 And
        Round(ControlLimitRejectionCriterion(0.99, 28), 4) = 2.7633 And
        Round(ControlLimitRejectionCriterion(0.99, 30), 4) = 2.5758 And
        Round(ControlLimitRejectionCriterion(0.99, 32), 4) = 2.5758 And
        Round(ControlLimitRejectionCriterion(0.99, 34), 4) = 2.5758 And
        Round(ControlLimitRejectionCriterion(0.95, 30), 4) = 1.96 And
        Round(ControlLimitRejectionCriterion(0.90, 30), 4) = 1.6449 And
        Round(ControlLimitRejectionCriterion(0.95, 25), 4) = 2.0595 And
        Round(ControlLimitRejectionCriterion(0.90, 20), 4) = 1.7247

      UnitTestsPassed = UnitTestsPassed And
        NonNaNCount({60}) = 1 And
        NonNaNCount({60, 70}) = 2 And
        NonNaNCount({60, 70, NaN}) = 2 And
        NonNaNCount({60, 70, NaN, 20}) = 3 And
        NonNaNCount({60, 70, NaN, 20, NaN}) = 3 And
        NonNaNCount({60, 70, NaN, 20, NaN, NaN}) = 3 And
        NonNaNCount({70, NaN, 20, NaN, NaN, 5}) = 3 And
        NonNaNCount({NaN, 20, NaN, NaN, 5, 10}) = 3 And
        NonNaNCount({20, NaN, NaN, 5, 10, 15}) = 4 And
        NonNaNCount({NaN, NaN, 5, 10, 15, 20}) = 4 And
        NonNaNCount({NaN, 5, 10, 15, 20, 25}) = 5 And
        NonNaNCount({5, 10, 15, 20, 25, 30}) = 6 And
        NonNaNCount({10, 15, 20, 25, 30}) = 5 And
        NonNaNCount({15, 20, 25, 30}) = 4 And
        NonNaNCount({20, 25, 30}) = 3 And
        NonNaNCount({25, 30}) = 2 And
        NonNaNCount({30}) = 1 And
        NonNaNCount({}) = 0 And
        NonNaNCount({NaN}) = 0 And
        NonNaNCount({NaN, NaN}) = 0

      UnitTestsPassed = UnitTestsPassed And
        Round(Mean({60}), 4) = 60 And
        Round(Mean({72}), 4) = 72 And
        Round(Mean({32, 95}), 4) = 63.5 And
        Round(Mean({81, 75}), 4) = 78 And
        Round(Mean({67, 76, 25}), 4) = 56 And
        Round(Mean({25, NaN, 27}), 4) = 26 And
        Round(Mean({31, 73, 83, 81}), 4) = 67 And
        Round(Mean({18, 58, 47, 47}), 4) = 42.5 And
        Round(Mean({10, NaN, 20, 30, NaN}), 4) = 20 And
        Round(Mean({67, 90, 74, 32, 62}), 4) = 65 And
        Round(Mean({78, 0, 98, 65, 69, 57}), 4) = 61.1667 And
        Round(Mean({49, 35, 74, 25, 28, 92}), 4) = 50.5 And
        Round(Mean({7, 64, 22, 7, 42, 34, 30}), 4) = 29.4286 And
        Round(Mean({13, 29, 39, 33, 96, 43, 17}), 4) = 38.5714 And
        Round(Mean({59, 78, 53, 7, 18, 44, 63, 40}), 4) = 45.25 And
        Round(Mean({77, 71, 99, 39, 50, 94, 67, 30}), 4) = 65.875 And
        Round(Mean({91, 69, 63, 5, 44, 93, 89, 45, 50}), 4) = 61 And
        Round(Mean({12, 84, 12, 94, 52, 17, 1, 13, 37}), 4) = 35.7778 And
        Round(Mean({18, 14, 54, 40, 73, 77, 4, 91, 53, 10}), 4) = 43.4 And
        Round(Mean({80, 30, 1, 92, 44, 61, 18, 72, 63, 41}), 4) = 50.2

      UnitTestsPassed = UnitTestsPassed And
        Median({57}) = 57 And
        Median({46}) = 46 And
        Median({79, 86}) = 82.5 And
        Median({46, 45}) = 45.5 And
        Median({10, NaN, 20}) = 15 And
        Median({58, 79, 68}) = 68 And
        Median({NaN, 10, 30, NaN}) = 20 And
        Median({30, 10, NaN, 15}) = 15 And
        IsNaN(Median({NaN, NaN, NaN, NaN, NaN})) And
        Median({19, 3, 9, 14, 31}) = 14 And
        Median({2, 85, 33, 10, 38, 56}) = 35.5 And
        Median({42, 65, 57, 92, 56, 59}) = 58 And
        Median({53, 17, 18, 73, 34, 96, 9}) = 34 And
        Median({23, 23, 43, 74, 8, 51, 88}) = 43 And
        Median({72, 46, 39, 7, 83, 96, 18, 50}) = 48 And
        Median({50, 25, 28, 99, 79, 97, 30, 16}) = 40 And
        Median({7, 22, 58, 32, 5, 90, 46, 91, 66}) = 46 And
        Median({81, 64, 5, 23, 48, 18, 19, 87, 15}) = 23 And
        Median({33, 82, 42, 33, 81, 56, 13, 13, 54, 6}) = 37.5 And
        Median({55, 40, 75, 23, 53, 85, 59, 9, 72, 44}) = 54

      UnitTestsPassed = UnitTestsPassed And
        Hash(AbsoluteDeviation({100, 44, 43, 45}, 44)) =
        Hash({56, 0, 1, 1}) And
        Hash(AbsoluteDeviation({76, 70, 84, 39}, 77)) =
        Hash({1, 7, 7, 38}) And
        Hash(AbsoluteDeviation({15, 74, 54, 23}, 93)) =
        Hash({78, 19, 39, 70}) And
        Hash(AbsoluteDeviation({48, 12, 85, 50}, 9)) =
        Hash({39, 3, 76, 41}) And
        Hash(AbsoluteDeviation({7, 24, 28, 41}, 69)) =
        Hash({62, 45, 41, 28}) And
        Hash(AbsoluteDeviation({55, 46, 81, 50}, 88)) =
        Hash({33, 42, 7, 38}) And
        Hash(AbsoluteDeviation({84, 14, 23, 54}, 52)) =
        Hash({32, 38, 29, 2}) And
        Hash(AbsoluteDeviation({92, 92, 45, 21}, 49)) =
        Hash({43, 43, 4, 28}) And
        Hash(AbsoluteDeviation({3, 9, 70, 51}, 36)) =
        Hash({33, 27, 34, 15}) And
        Hash(AbsoluteDeviation({14, 7, 8, 53}, 37)) =
        Hash({23, 30, 29, 16}) And
        Hash(AbsoluteDeviation({66, 37, 23, 80}, 3)) =
        Hash({63, 34, 20, 77}) And
        Hash(AbsoluteDeviation({93, 45, 18, 50}, 75)) =
        Hash({18, 30, 57, 25}) And
        Hash(AbsoluteDeviation({2, 73, 13, 0}, 37)) =
        Hash({35, 36, 24, 37}) And
        Hash(AbsoluteDeviation({76, 77, 56, 11}, 77)) =
        Hash({1, 0, 21, 66}) And
        Hash(AbsoluteDeviation({90, 36, 72, 54}, 81)) =
        Hash({9, 45, 9, 27}) And
        Hash(AbsoluteDeviation({5, 62, 50, 23}, 19)) =
        Hash({14, 43, 31, 4}) And
        Hash(AbsoluteDeviation({79, 51, 54, 24}, 60)) =
        Hash({19, 9, 6, 36}) And
        Hash(AbsoluteDeviation({26, 53, 38, 42}, 60)) =
        Hash({34, 7, 22, 18}) And
        Hash(AbsoluteDeviation({4, 100, 32, 23}, 32)) =
        Hash({28, 68, 0, 9}) And
        Hash(AbsoluteDeviation({15, 94, 58, 67}, 78)) =
        Hash({63, 16, 20, 11})

      UnitTestsPassed = UnitTestsPassed And
        Round(MeanAbsoluteDeviation(
        {19}), 4) = 0 And
        Round(MeanAbsoluteDeviation(
        {86}), 4) = 0 And
        Round(MeanAbsoluteDeviation(
        {7, 24}), 4) = 8.5 And
        Round(MeanAbsoluteDeviation(
        {12, 96}), 4) = 42 And
        Round(MeanAbsoluteDeviation(
        {19, 74, 70}), 4) = 23.5556 And
        Round(MeanAbsoluteDeviation(
        {96, 93, 65}), 4) = 13.1111 And
        Round(MeanAbsoluteDeviation(
        {47, 29, 24, 11}), 4) = 10.25 And
        Round(MeanAbsoluteDeviation(
        {3, 43, 53, 80}), 4) = 21.75 And
        Round(MeanAbsoluteDeviation(
        {17, 27, 98, 85, 51}), 4) = 28.72 And
        Round(MeanAbsoluteDeviation(
        {2, 82, 63, 1, 49}), 4) = 30.32 And
        Round(MeanAbsoluteDeviation(
        {9, 25, 41, 85, 82, 55}), 4) = 24.5 And
        Round(MeanAbsoluteDeviation(
        {5, 74, 53, 97, 81, 21}), 4) = 28.8333 And
        Round(MeanAbsoluteDeviation(
        {26, 81, 9, 18, 39, 97, 21}), 4) = 27.102 And
        Round(MeanAbsoluteDeviation(
        {5, 83, 31, 24, 55, 22, 87}), 4) = 26.6939 And
        Round(MeanAbsoluteDeviation(
        {22, 84, 6, 79, 89, 71, 34, 56}), 4) = 25.8438 And
        Round(MeanAbsoluteDeviation(
        {33, 39, 6, 88, 69, 11, 76, 65}), 4) = 26.125 And
        Round(MeanAbsoluteDeviation(
        {31, 52, 12, 60, 52, 44, 47, 81, 34}), 4) = 13.9012 And
        Round(MeanAbsoluteDeviation(
        {64, 63, 54, 94, 25, 80, 97, 45, 51}), 4) = 17.8519 And
        Round(MeanAbsoluteDeviation(
        {47, 22, 52, 22, 10, 38, 94, 85, 54, 41}), 4) = 19.9 And
        Round(MeanAbsoluteDeviation(
        {7, 12, 84, 29, 41, 8, 18, 15, 16, 84}), 4) = 22.96

      UnitTestsPassed = UnitTestsPassed And
        MedianAbsoluteDeviation(
        {2}) = 0 And
        MedianAbsoluteDeviation(
        {37}) = 0 And
        MedianAbsoluteDeviation(
        {87, 37}) = 25 And
        MedianAbsoluteDeviation(
        {13, 74}) = 30.5 And
        MedianAbsoluteDeviation(
        {39, 52, 93}) = 13 And
        MedianAbsoluteDeviation(
        {90, 24, 47}) = 23 And
        MedianAbsoluteDeviation(
        {11, 51, 20, 62}) = 20 And
        MedianAbsoluteDeviation(
        {74, 35, 9, 95}) = 30 And
        MedianAbsoluteDeviation(
        {32, 46, 15, 90, 66}) = 20 And
        MedianAbsoluteDeviation(
        {91, 19, 50, 55, 44}) = 6 And
        MedianAbsoluteDeviation(
        {2, 64, 87, 65, 61, 97}) = 13 And
        MedianAbsoluteDeviation(
        {35, 66, 73, 74, 71, 93}) = 4 And
        MedianAbsoluteDeviation(
        {54, 81, 80, 36, 11, 36, 45}) = 9 And
        MedianAbsoluteDeviation(
        {14, 69, 40, 68, 75, 10, 69}) = 7 And
        MedianAbsoluteDeviation(
        {40, 51, 28, 21, 91, 95, 66, 3}) = 22.5 And
        MedianAbsoluteDeviation(
        {57, 87, 94, 46, 51, 27, 10, 7}) = 30 And
        MedianAbsoluteDeviation(
        {3, 89, 62, 84, 86, 37, 14, 72, 33}) = 25 And
        MedianAbsoluteDeviation(
        {48, 6, 14, 2, 74, 89, 15, 8, 83}) = 13 And
        MedianAbsoluteDeviation(
        {2, 22, 91, 84, 43, 96, 55, 3, 9, 11}) = 26.5 And
        MedianAbsoluteDeviation(
        {6, 96, 82, 26, 47, 84, 34, 39, 60, 99}) = 28

      UnitTestsPassed = UnitTestsPassed And
        Not UseMeanAbsoluteDeviation({6, 5, 5, 9}) And
        Not UseMeanAbsoluteDeviation({2, 2, 10, 9}) And
        Not UseMeanAbsoluteDeviation({4, 0, 10, 6}) And
        Not UseMeanAbsoluteDeviation({6, 10, 1, 1}) And
        UseMeanAbsoluteDeviation({3, 0, 0, 0}) And
        Not UseMeanAbsoluteDeviation({10, 0, 8, 5}) And
        Not UseMeanAbsoluteDeviation({7, 4, 5, 3}) And
        UseMeanAbsoluteDeviation({5, 1, 5, 5}) And
        Not UseMeanAbsoluteDeviation({2, 7, 3, 8}) And
        Not UseMeanAbsoluteDeviation({2, 6, 10, 1}) And
        Not UseMeanAbsoluteDeviation({1, 6, 3, 5}) And
        Not UseMeanAbsoluteDeviation({3, 9, 7, 3}) And
        UseMeanAbsoluteDeviation({5, 5, 8, 5}) And
        Not UseMeanAbsoluteDeviation({5, 10, 5, 4}) And
        Not UseMeanAbsoluteDeviation({0, 2, 4, 1}) And
        Not UseMeanAbsoluteDeviation({7, 3, 0, 10}) And
        UseMeanAbsoluteDeviation({4, 4, 4, 0}) And
        Not UseMeanAbsoluteDeviation({5, 7, 4, 5}) And
        UseMeanAbsoluteDeviation({2, 2, 2, 2}) And
        UseMeanAbsoluteDeviation({9, 4, 4, 4})

      UnitTestsPassed = UnitTestsPassed And
        CentralTendency({3, 0, 0, 0}) = 0.75 And
        CentralTendency({2, 10, 0, 1}) = 1.5 And
        CentralTendency({7, 7, 7, 1}) = 5.5 And
        CentralTendency({5, 7, 2, 8}) = 6 And
        CentralTendency({9, 3, 4, 5}) = 4.5 And
        CentralTendency({3, 3, 3, 3}) = 3 And
        CentralTendency({8, 4, 2, 10}) = 6 And
        CentralTendency({2, 1, 10, 10}) = 6 And
        CentralTendency({3, 3, 6, 2}) = 3 And
        CentralTendency({9, 9, 6, 5}) = 7.5 And
        CentralTendency({2, 8, 8, 9}) = 8 And
        CentralTendency({7, 7, 4, 1}) = 5.5 And
        CentralTendency({5, 5, 5, 0}) = 3.75 And
        CentralTendency({4, 2, 3, 7}) = 3.5 And
        CentralTendency({2, 1, 5, 1}) = 1.5 And
        CentralTendency({9, 4, 5, 0}) = 4.5 And
        CentralTendency({1, 1, 7, 1}) = 2.5 And
        CentralTendency({1, 5, 9, 5}) = 5 And
        CentralTendency({3, 5, 1, 9}) = 4 And
        CentralTendency({0, 0, 0, 0}) = 0

      UnitTestsPassed = UnitTestsPassed And
        Round(ControlLimit({8, 2, 0, 10}, 0.99), 4) = 30.5449 And
        Round(ControlLimit({2, 4, 8, 7}, 0.99), 4) = 15.2724 And
        Round(ControlLimit({5, 8, 0, 2}, 0.99), 4) = 19.0906 And
        Round(ControlLimit({8, 1, 0, 3}, 0.99), 4) = 11.4543 And
        Round(ControlLimit({10, 7, 1, 3}, 0.99), 4) = 22.9087 And
        Round(ControlLimit({6, 2, 1, 9}, 0.99), 4) = 19.0906 And
        Round(ControlLimit({4, 9, 9, 3}, 0.99), 4) = 19.0906 And
        Round(ControlLimit({10, 7, 2, 8}, 0.99), 4) = 11.4543 And
        Round(ControlLimit({6, 0, 10, 1}, 0.99), 4) = 22.9087 And
        Round(ControlLimit({10, 3, 4, 2}, 0.99), 4) = 7.6362 And
        Round(ControlLimit({6, 4, 4, 4}, 0.99), 4) = 5.4904 And
        Round(ControlLimit({1, 0, 9, 9}, 0.99), 4) = 30.5449 And
        Round(ControlLimit({0, 3, 6, 2}, 0.99), 4) = 11.4543 And
        Round(ControlLimit({9, 7, 4, 6}, 0.99), 4) = 11.4543 And
        Round(ControlLimit({6, 6, 4, 1}, 0.99), 4) = 7.6362 And
        Round(ControlLimit({7, 3, 4, 1}, 0.99), 4) = 11.4543 And
        Round(ControlLimit({6, 4, 4, 10}, 0.99), 4) = 7.6362 And
        Round(ControlLimit({10, 5, 5, 5}, 0.99), 4) = 13.7259 And
        Round(ControlLimit({8, 5, 5, 5}, 0.98), 4) = 6.4023 And
        Round(ControlLimit({8, 4, 0, 0}, 0.95), 4) = 8.3213

      UnitTestsPassed = UnitTestsPassed And
        Hash(RemoveOutliers({100, 100, 100, 100, 100, 100, 100, 100, 999})) =
        Hash({100, 100, 100, 100, 100, 100, 100, 100, NaN}) And
        Hash(RemoveOutliers({100, 101, 102, 103, 104, 105, 106, 107, 999})) =
        Hash({100, 101, 102, 103, 104, 105, 106, 107, NaN}) And
        Hash(RemoveOutliers({2223.6946, 2770.1624, 2125.7544, 3948.9927,
        2184.2341, 2238.6421, 2170.0227, 2967.0674, 2177.3738, 3617.1328,
        2460.8193, 3315.8684})) = Hash({2223.6946, 2770.1624, 2125.7544, NaN,
        2184.2341, 2238.6421, 2170.0227, 2967.0674, 2177.3738, NaN,
        2460.8193, NaN}) And
        Hash(RemoveOutliers({3355.1553, 3624.3154, 3317.6895, 3610.0039,
        3990.751, 2950.4382, 2140.5908, 3237.4917, 3319.7139, 2829.2725,
        3406.9199, 3230.0078})) = Hash({3355.1553, 3624.3154, 3317.6895,
        3610.0039, 3990.751, 2950.4382, NaN, 3237.4917, 3319.7139, 2829.2725,
        3406.9199, 3230.0078}) And
        Hash(RemoveOutliers({2969.7808, 3899.0913, 2637.4045, 2718.73,
        2960.9597, 2650.6521, 2707.4294, 2034.5339, 2935.9111, 3458.7085,
        2584.53, 3999.4238})) = Hash({2969.7808, NaN, 2637.4045, 2718.73,
        2960.9597, 2650.6521, 2707.4294, 2034.5339, 2935.9111, 3458.7085,
        2584.53, NaN}) And
        Hash(RemoveOutliers({2774.8018, 2755.0251, 2756.6152, 3800.0625,
        2900.0671, 2784.0134, 3955.2947, 2847.0908, 2329.7837, 3282.4614,
        2597.1582, 3009.8796})) = Hash({2774.8018, 2755.0251, 2756.6152, NaN,
        2900.0671, 2784.0134, NaN, 2847.0908, 2329.7837, 3282.4614,
        2597.1582, 3009.8796}) And
        Hash(RemoveOutliers({3084.8821, 3394.1196, 3131.3245, 2799.9587,
        2528.3088, 3015.4998, 2912.2029, 2022.2645, 3666.5674, 3685.1973,
        3149.6931, 3070.0479})) = Hash({3084.8821, 3394.1196, 3131.3245,
        2799.9587, 2528.3088, 3015.4998, 2912.2029, NaN, 3666.5674,
        3685.1973, 3149.6931, 3070.0479}) And
        Hash(RemoveOutliers({3815.72, 3063.9106, 3535.0366, 2349.564,
        2597.2661, 3655.3076, 3452.7407, 2020.7682, 3810.7046, 3833.8396,
        3960.6016, 3866.8149})) = Hash({3815.72, 3063.9106, 3535.0366, NaN,
        2597.2661, 3655.3076, 3452.7407, NaN, 3810.7046, 3833.8396,
        3960.6016, 3866.8149}) And
        Hash(RemoveOutliers({2812.0613, 3726.7427, 2090.9749, 2548.4485,
        3900.5151, 3545.854, 3880.2229, 3940.9585, 3942.2234, 3263.0137,
        3701.8882, 2056.5291})) = Hash({2812.0613, 3726.7427, NaN, 2548.4485,
        3900.5151, 3545.854, 3880.2229, 3940.9585, 3942.2234, 3263.0137,
        3701.8882, NaN}) And
        Hash(RemoveOutliers({3798.4775, 2959.3879, 2317.3547, 2596.3599,
        2075.6292, 2563.9685, 2695.5081, 2386.2161, 2433.1106, 2810.3716,
        2499.7554, 3843.103})) = Hash({NaN, 2959.3879, 2317.3547, 2596.3599,
        2075.6292, 2563.9685, 2695.5081, 2386.2161, 2433.1106, 2810.3716,
        2499.7554, NaN}) And
        Hash(RemoveOutliers({2245.7856, 2012.4834, 2473.0103, 2684.5693,
        2645.4729, 2851.019, 2344.6099, 2408.1492, 3959.5967, 3954.0583,
        2399.2617, 2652.8855})) = Hash({2245.7856, 2012.4834, 2473.0103,
        2684.5693, 2645.4729, 2851.019, 2344.6099, 2408.1492, NaN, NaN,
        2399.2617, 2652.8855}) And
        Hash(RemoveOutliers({2004.5355, 2743.0693, 3260.7441, 2382.8906,
        2365.9385, 2243.333, 3506.5352, 3905.7717, 3516.5337, 2133.8328,
        2308.1809, 2581.4009})) = Hash({2004.5355, 2743.0693, 3260.7441,
        2382.8906, 2365.9385, 2243.333, 3506.5352, NaN, 3516.5337, 2133.8328,
        2308.1809, 2581.4009}) And
        Hash(RemoveOutliers({3250.5376, 3411.313, 2037.264, 3709.5815,
        3417.1167, 3996.0493, 3529.637, 3992.7163, 2786.95, 3728.834,
        3304.4272, 2248.9119})) = Hash({3250.5376, 3411.313, NaN, 3709.5815,
        3417.1167, 3996.0493, 3529.637, 3992.7163, 2786.95, 3728.834,
        3304.4272, 2248.9119}) And
        Hash(RemoveOutliers({2398.3125, 2742.4028, 2720.752, 2628.8442,
        2750.1482, 2724.4932, 2161.6875, 2644.4163, 2188.2952, 2455.4622,
        3332.5503, 2540.5198})) = Hash({2398.3125, 2742.4028, 2720.752,
        2628.8442, 2750.1482, 2724.4932, 2161.6875, 2644.4163, 2188.2952,
        2455.4622, NaN, 2540.5198}) And
        Hash(RemoveOutliers({3991.7854, 3607.98, 2686.032, 2546.969,
        3053.8796, 3138.9824, 2441.1689, 2737.1245, 2616.7139, 2550.5774,
        2406.0913, 2743.2361})) = Hash({NaN, 3607.98, 2686.032, 2546.969,
        3053.8796, 3138.9824, 2441.1689, 2737.1245, 2616.7139, 2550.5774,
        2406.0913, 2743.2361}) And
        Hash(RemoveOutliers({2361.5334, 3636.4312, 2187.593, 2281.5432,
        2132.3833, 2056.792, 2227.7795, 2757.1753, 3416.9126, 2568.927,
        2094.2065, 3449.3984})) = Hash({2361.5334, NaN, 2187.593, 2281.5432,
        2132.3833, 2056.792, 2227.7795, 2757.1753, NaN, 2568.927, 2094.2065,
        NaN}) And
        Hash(RemoveOutliers({2249.7119, 2411.8374, 3041.5498, 2679.1458,
        2561.1577, 2405.7229, 2775.2253, 2832.1233, 2540.2134, 3654.5903,
        3970.5173, 2920.5637})) = Hash({2249.7119, 2411.8374, 3041.5498,
        2679.1458, 2561.1577, 2405.7229, 2775.2253, 2832.1233, 2540.2134,
        3654.5903, NaN, 2920.5637}) And
        Hash(RemoveOutliers({2038.1091, 2248.3057, 2427.1646, 2337.2427,
        2642.043, 3497.5393, 3996.3579, 2178.979, 3968.8848, 3460.8613,
        2774.8486, 2338.1362})) = Hash({2038.1091, 2248.3057, 2427.1646,
        2337.2427, 2642.043, 3497.5393, NaN, 2178.979, NaN, 3460.8613,
        2774.8486, 2338.1362}) And
        Hash(RemoveOutliers({3010.9485, 2517.2876, 2057.7188, 2133.0801,
        3192.0308, 2035.0759, 3821.248, 2391.8086, 2267.896, 3751.3276,
        2340.9497, 2327.333})) = Hash({3010.9485, 2517.2876, 2057.7188,
        2133.0801, 3192.0308, 2035.0759, NaN, 2391.8086, 2267.896, NaN,
        2340.9497, 2327.333}) And
        Hash(RemoveOutliers({NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN,
        999})) = Hash({NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN, NaN})

      UnitTestsPassed = UnitTestsPassed And
        Round(ExponentialMovingAverage(
        {70.5547}), 4) = 70.5547 And
        Round(ExponentialMovingAverage(
        {53.3424, 57.9519}), 4) = 56.7995 And
        Round(ExponentialMovingAverage(
        {28.9562, 30.1948}), 4) = 29.8852 And
        Round(ExponentialMovingAverage(
        {77.474, 1.4018}), 4) = 20.4199 And
        Round(ExponentialMovingAverage(
        {76.0724, 81.449, 70.9038}), 4) = 74.6551 And
        Round(ExponentialMovingAverage(
        {4.5353, 41.4033, 86.2619}), 4) = 61.7699 And
        Round(ExponentialMovingAverage(
        {79.048, 37.3536, 96.1953}), 4) = 76.9338 And
        Round(ExponentialMovingAverage(
        {87.1446, 5.6237, 94.9557}), 4) = 68.3164 And
        Round(ExponentialMovingAverage(
        {36.4019, 52.4868, 76.7112}), 4) = 64.0315 And
        Round(ExponentialMovingAverage(
        {5.3505, 59.2458, 46.87, 29.8165}), 4) = 36.959 And
        Round(ExponentialMovingAverage(
        {62.2697, 64.7821, 26.3793, 27.9342}), 4) = 37.0099 And
        Round(ExponentialMovingAverage(
        {82.9802, 82.4602, 58.9163, 98.6093}), 4) = 83.4414 And
        Round(ExponentialMovingAverage(
        {91.0964, 22.6866, 69.5116, 98.0003, 24.3931}), 4) = 55.7929 And
        Round(ExponentialMovingAverage(
        {53.3873, 10.637, 99.9415, 67.6176, 1.5704}), 4) = 40.2177 And
        Round(ExponentialMovingAverage(
        {57.5184, 10.0052, 10.3023, 79.8884, 28.448}), 4) = 38.6235 And
        Round(ExponentialMovingAverage(
        {4.5649, 29.5773, 38.2011, 30.097, 94.8571}), 4) = 54.345 And
        Round(ExponentialMovingAverage(
        {97.9829, 40.1374, 27.828, 16.0442, 16.2822}), 4) = 27.0999 And
        Round(ExponentialMovingAverage({64.6587,
        41.0073, 41.2767, 71.273, 32.6206, 63.3179}), 4) = 52.9531 And
        Round(ExponentialMovingAverage({20.7561,
        18.6014, 58.3359, 8.0715, 45.7971, 90.573}), 4) = 51.847 And
        Round(ExponentialMovingAverage(
        {26.1368, 78.5212, 37.8903, 28.9665, 91.9377, 63.1742}), 4) = 60.2045

      UnitTestsPassed = UnitTestsPassed And
        Round(SlopeToAngle(-4.5806), 4) = -77.6849 And
        Round(SlopeToAngle(-4.2541), 4) = -76.7718 And
        Round(SlopeToAngle(1.7964), 4) = 60.8967 And
        Round(SlopeToAngle(-3.2474), 4) = -72.8844 And
        Round(SlopeToAngle(4.7917), 4) = 78.2119 And
        Round(SlopeToAngle(2.1792), 4) = 65.3504 And
        Round(SlopeToAngle(0.4736), 4) = 25.3422 And
        Round(SlopeToAngle(-2.0963), 4) = -64.4974 And
        Round(SlopeToAngle(-3.2077), 4) = -72.6851 And
        Round(SlopeToAngle(-1.5425), 4) = -57.0447 And
        Round(SlopeToAngle(-0.5587), 4) = -29.1921 And
        Round(SlopeToAngle(1.2829), 4) = 52.0642 And
        Round(SlopeToAngle(3.9501), 4) = 75.7936 And
        Round(SlopeToAngle(2.5841), 4) = 68.8445 And
        Round(SlopeToAngle(-3.4547), 4) = -73.8563 And
        Round(SlopeToAngle(3.2931), 4) = 73.1083 And
        Round(SlopeToAngle(3.2042), 4) = 72.6674 And
        Round(SlopeToAngle(3.1088), 4) = 72.1687 And
        Round(SlopeToAngle(-1.6831), 4) = -59.2837 And
        Round(SlopeToAngle(-2.0031), 4) = -63.4704

      For i = 0 To 19
        UnitTestsPassed = UnitTestsPassed And
          RandomNumber(0, i+1) >= 0 And
          RandomNumber(0, i+1) <= i+1
      Next i

      UnitTestsPassed = UnitTestsPassed And
        IsNaN(AlignPreviousInterval(0, 0)) And
        AlignPreviousInterval(0, 12) = 0 And
        AlignPreviousInterval(0, -12) = 0 And
        AlignPreviousInterval(353, 84) = 336 And
        AlignPreviousInterval(512, 71) = 497 And
        AlignPreviousInterval(-651, 34) = -680 And
        AlignPreviousInterval(-136, -20) = -120 And
        AlignPreviousInterval(800, -118) = 826 And
        AlignPreviousInterval(-671, 81) = -729 And
        AlignPreviousInterval(-769, -124) = -744 And
        AlignPreviousInterval(-676, -61) = -671 And
        AlignPreviousInterval(627, -14) = 630 And
        AlignPreviousInterval(337, -68) = 340 And
        AlignPreviousInterval(661, 37) = 629 And
        AlignPreviousInterval(228, 57) = 228 And
        AlignPreviousInterval(686, -22) = 704 And
        AlignPreviousInterval(846, -35) = 875 And
        AlignPreviousInterval(571, 108) = 540 And
        AlignPreviousInterval(-531, -56) = -504 And
        AlignPreviousInterval(-880, 105) = -945

      UnitTestsPassed = UnitTestsPassed And
        AlignTimestamp(New DateTime(2016, 4, 4, 16, 33, 2), 60) =
        New DateTime(2016, 4, 4, 16, 33, 0) And
        AlignTimestamp(New DateTime(2015, 7, 15, 2, 29, 58), 60) = 
        New DateTime(2015, 7, 15, 2, 29, 0) And
        AlignTimestamp(New DateTime(2016, 4, 1, 22, 5, 17), 60) = 
        New DateTime(2016, 4, 1, 22, 5, 0) And
        AlignTimestamp(New DateTime(2013, 12, 1, 21, 47, 35), 60) = 
        New DateTime(2013, 12, 1, 21, 47, 0) And
        AlignTimestamp(New DateTime(2016, 11, 22, 0, 22, 17), 60) = 
        New DateTime(2016, 11, 22, 0, 22, 0) And
        AlignTimestamp(New DateTime(2016, 10, 11, 19, 11, 41), 300) = 
        New DateTime(2016, 10, 11, 19, 10, 0) And
        AlignTimestamp(New DateTime(2013, 10, 26, 4, 24, 53), 300) = 
        New DateTime(2013, 10, 26, 4, 20, 0) And
        AlignTimestamp(New DateTime(2014, 5, 2, 2, 52, 41), 300) = 
        New DateTime(2014, 5, 2, 2, 50, 0) And
        AlignTimestamp(New DateTime(2014, 8, 16, 13, 11, 10), 300) = 
        New DateTime(2014, 8, 16, 13, 10, 0) And
        AlignTimestamp(New DateTime(2014, 10, 25, 8, 26, 4), 300) = 
        New DateTime(2014, 10, 25, 8, 25, 0) And
        AlignTimestamp(New DateTime(2015, 6, 2, 18, 36, 24), 3600) = 
        New DateTime(2015, 6, 2, 18, 0, 0) And
        AlignTimestamp(New DateTime(2016, 11, 21, 16, 24, 27), 3600) = 
        New DateTime(2016, 11, 21, 16, 0, 0) And
        AlignTimestamp(New DateTime(2014, 4, 4, 8, 42, 10), 3600) = 
        New DateTime(2014, 4, 4, 8, 0, 0) And
        AlignTimestamp(New DateTime(2016, 2, 22, 19, 8, 41), 3600) = 
        New DateTime(2016, 2, 22, 19, 0, 0) And
        AlignTimestamp(New DateTime(2015, 9, 13, 22, 48, 17), 3600) = 
        New DateTime(2015, 9, 13, 22, 0, 0) And
        AlignTimestamp(New DateTime(2016, 10, 20, 2, 47, 48), 86400) = 
        New DateTime(2016, 10, 20, 0, 0, 0) And
        AlignTimestamp(New DateTime(2014, 2, 8, 23, 12, 34), 86400) = 
        New DateTime(2014, 2, 8, 0, 0, 0) And
        AlignTimestamp(New DateTime(2016, 2, 27, 23, 40, 39), 86400) = 
        New DateTime(2016, 2, 27, 0, 0, 0) And
        AlignTimestamp(New DateTime(2015, 8, 26, 9, 35, 55), 86400) = 
        New DateTime(2015, 8, 26, 0, 0, 0) And
        AlignTimestamp(New DateTime(2016, 2, 11, 0, 44, 7), 86400) = 
        New DateTime(2016, 2, 11, 0, 0, 0)

      UnitTestsPassed = UnitTestsPassed And
        NextInterval(New DateTime(2016, 4, 4, 16, 33, 2)) =
        New DateTime(2016, 4, 4, 16, 35, 0) And
        NextInterval(New DateTime(2015, 7, 15, 2, 29, 58)) = 
        New DateTime(2015, 7, 15, 2, 30, 0) And
        NextInterval(New DateTime(2016, 4, 1, 22, 5, 17)) = 
        New DateTime(2016, 4, 1, 22, 10, 0) And
        NextInterval(New DateTime(2013, 12, 1, 21, 47, 35)) = 
        New DateTime(2013, 12, 1, 21, 50, 0) And
        NextInterval(New DateTime(2016, 11, 22, 0, 22, 17)) = 
        New DateTime(2016, 11, 22, 0, 25, 0) And
        NextInterval(New DateTime(2016, 10, 11, 19, 11, 41)) = 
        New DateTime(2016, 10, 11, 19, 15, 0) And
        NextInterval(New DateTime(2013, 10, 26, 4, 24, 53)) = 
        New DateTime(2013, 10, 26, 4, 25, 0) And
        NextInterval(New DateTime(2014, 5, 2, 2, 52, 41)) = 
        New DateTime(2014, 5, 2, 2, 55, 0) And
        NextInterval(New DateTime(2014, 8, 16, 13, 11, 10)) = 
        New DateTime(2014, 8, 16, 13, 15, 0) And
        NextInterval(New DateTime(2014, 10, 25, 8, 26, 4)) = 
        New DateTime(2014, 10, 25, 8, 30, 0) And
        NextInterval(New DateTime(2015, 6, 2, 18, 36, 24), 1) = 
        New DateTime(2015, 6, 2, 18, 40, 0) And
        NextInterval(New DateTime(2016, 11, 21, 16, 24, 27), 2) = 
        New DateTime(2016, 11, 21, 16, 30, 0) And
        NextInterval(New DateTime(2014, 4, 4, 8, 42, 10), 3) = 
        New DateTime(2014, 4, 4, 8, 55, 0) And
        NextInterval(New DateTime(2016, 2, 22, 19, 8, 41), 4) = 
        New DateTime(2016, 2, 22, 19, 25, 0) And
        NextInterval(New DateTime(2015, 9, 13, 22, 48, 17), 5) = 
        New DateTime(2015, 9, 13, 23, 10, 0) And
        NextInterval(New DateTime(2016, 10, 20, 2, 47, 48), 6) = 
        New DateTime(2016, 10, 20, 3, 15, 0) And
        NextInterval(New DateTime(2014, 2, 8, 23, 12, 34), 7) = 
        New DateTime(2014, 2, 8, 23, 45, 0) And
        NextInterval(New DateTime(2016, 2, 27, 23, 40, 39), 8) = 
        New DateTime(2016, 2, 28, 0, 20, 0) And
        NextInterval(New DateTime(2015, 8, 26, 9, 35, 55), -12) = 
        New DateTime(2015, 8, 26, 8, 35, 0) And
        NextInterval(New DateTime(2016, 2, 11, 0, 44, 7), 0) = 
        New DateTime(2016, 2, 11, 0, 40, 0)

      UnitTestsPassed = UnitTestsPassed And
        Computus(1864) = New DateTime(1864, 3, 27) And
        Computus(1900) = New DateTime(1900, 4, 15) And
        Computus(1933) = New DateTime(1933, 4, 16) And
        Computus(1999) = New DateTime(1999, 4, 4) And
        Computus(2001) = New DateTime(2001, 4, 15) And
        Computus(2003) = New DateTime(2003, 4, 20) And
        Computus(2005) = New DateTime(2005, 3, 27) And
        Computus(2007) = New DateTime(2007, 4, 8) And
        Computus(2013) = New DateTime(2013, 3, 31) And
        Computus(2017) = New DateTime(2017, 4, 16) And
        Computus(2019) = New DateTime(2019, 4, 21) And
        Computus(2021) = New DateTime(2021, 4, 4) And
        Computus(2023) = New DateTime(2023, 4, 9) And
        Computus(2027) = New DateTime(2027, 3, 28) And
        Computus(2031) = New DateTime(2031, 4, 13) And
        Computus(2033) = New DateTime(2033, 4, 17) And
        Computus(2037) = New DateTime(2037, 4, 5) And
        Computus(2099) = New DateTime(2099, 4, 12) And
        Computus(2172) = New DateTime(2172, 4, 12) And
        Computus(2292) = New DateTime(2292, 4, 10)

      UnitTestsPassed = UnitTestsPassed And
        IsHoliday(New DateTime(2021, 12, 25), InvariantCulture) And
        Not IsHoliday(New DateTime(2020, 10, 12), InvariantCulture) And
        IsHoliday(New DateTime(2012, 1, 1), New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2016, 3, 26), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2016, 3, 28), New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2016, 3, 29), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2012, 4, 30), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2018, 4, 27, 23, 59, 59),
        New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2014, 5, 29), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2020, 6, 1), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2009, 12, 25), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2011, 12, 26), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2010, 12, 31), New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2014, 2, 1, 0, 0, 1),
        New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2009, 2, 23), New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2014, 5, 5), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2015, 5, 5), New CultureInfo("nl-NL")) And
        Not IsHoliday(New DateTime(2011, 4, 11), New CultureInfo("nl-NL")) And
        IsHoliday(New DateTime(2022, 4, 17), New CultureInfo("en-US")) And
        Not IsHoliday(New DateTime(2020, 6, 1), New CultureInfo("en-US"))

      UnitTestsPassed = UnitTestsPassed And
        DaysSinceSunday(New DateTime(2016, 4, 4, 16, 33, 2)) = 1 And
        DaysSinceSunday(New DateTime(2015, 7, 15, 2, 29, 58)) = 3 And
        DaysSinceSunday(New DateTime(2016, 4, 1, 22, 5, 17)) = 5 And
        DaysSinceSunday(New DateTime(2013, 12, 1, 21, 47, 35)) = 0 And
        DaysSinceSunday(New DateTime(2016, 11, 22, 0, 22, 17)) = 2 And
        DaysSinceSunday(New DateTime(2016, 10, 11, 19, 11, 41)) = 2 And
        DaysSinceSunday(New DateTime(2013, 10, 26, 4, 24, 53)) = 6 And
        DaysSinceSunday(New DateTime(2014, 5, 2, 2, 52, 41)) = 5 And
        DaysSinceSunday(New DateTime(2014, 8, 16, 13, 11, 10)) = 6 And
        DaysSinceSunday(New DateTime(2014, 10, 25, 8, 26, 4)) = 6 And
        DaysSinceSunday(New DateTime(2015, 6, 2, 18, 36, 24)) = 2 And
        DaysSinceSunday(New DateTime(2016, 11, 21, 16, 24, 27)) = 1 And
        DaysSinceSunday(New DateTime(2014, 4, 4, 8, 42, 10)) = 5 And
        DaysSinceSunday(New DateTime(2016, 2, 22, 19, 8, 41)) = 1 And
        DaysSinceSunday(New DateTime(2015, 9, 13, 22, 48, 17)) = 0 And
        DaysSinceSunday(New DateTime(2016, 10, 20, 2, 47, 48)) = 4 And
        DaysSinceSunday(New DateTime(2014, 2, 8, 23, 12, 34)) = 6 And
        DaysSinceSunday(New DateTime(2016, 2, 27, 23, 40, 39)) = 6 And
        DaysSinceSunday(New DateTime(2015, 8, 26, 9, 35, 55)) = 3 And
        DaysSinceSunday(New DateTime(2016, 2, 11, 0, 44, 7)) = 4

      UnitTestsPassed = UnitTestsPassed And
        PreviousSunday(New DateTime(2016, 4, 4, 16, 33, 2)) =
        New DateTime(2016, 4, 3, 16, 33, 2) And
        PreviousSunday(New DateTime(2015, 7, 15, 2, 29, 58)) =
        New DateTime(2015, 7, 12, 2, 29, 58) And
        PreviousSunday(New DateTime(2016, 4, 1, 22, 5, 17)) =
        New DateTime(2016, 3, 27, 22, 5, 17) And
        PreviousSunday(New DateTime(2013, 12, 1, 21, 47, 35)) =
        New DateTime(2013, 12, 1, 21, 47, 35) And
        PreviousSunday(New DateTime(2016, 11, 22, 0, 22, 17)) =
        New DateTime(2016, 11, 20, 0, 22, 17) And
        PreviousSunday(New DateTime(2016, 10, 11, 19, 11, 41)) =
        New DateTime(2016, 10, 9, 19, 11, 41) And
        PreviousSunday(New DateTime(2013, 10, 26, 4, 24, 53)) =
        New DateTime(2013, 10, 20, 4, 24, 53) And
        PreviousSunday(New DateTime(2014, 5, 2, 2, 52, 41)) =
        New DateTime(2014, 4, 27, 2, 52, 41) And
        PreviousSunday(New DateTime(2014, 8, 16, 13, 11, 10)) =
        New DateTime(2014, 8, 10, 13, 11, 10) And
        PreviousSunday(New DateTime(2014, 10, 25, 8, 26, 4)) =
        New DateTime(2014, 10, 19, 8, 26, 4) And
        PreviousSunday(New DateTime(2015, 6, 2, 18, 36, 24)) =
        New DateTime(2015, 5, 31, 18, 36, 24) And
        PreviousSunday(New DateTime(2016, 11, 21, 16, 24, 27)) =
        New DateTime(2016, 11, 20, 16, 24, 27) And
        PreviousSunday(New DateTime(2014, 4, 4, 8, 42, 10)) =
        New DateTime(2014, 3, 30, 8, 42, 10) And
        PreviousSunday(New DateTime(2016, 2, 22, 19, 8, 41)) =
        New DateTime(2016, 2, 21, 19, 8, 41) And
        PreviousSunday(New DateTime(2015, 9, 13, 22, 48, 17)) =
        New DateTime(2015, 9, 13, 22, 48, 17) And
        PreviousSunday(New DateTime(2016, 10, 20, 2, 47, 48)) =
        New DateTime(2016, 10, 16, 2, 47, 48) And
        PreviousSunday(New DateTime(2014, 2, 8, 23, 12, 34)) =
        New DateTime(2014, 2, 2, 23, 12, 34) And
        PreviousSunday(New DateTime(2016, 2, 27, 23, 40, 39)) =
        New DateTime(2016, 2, 21, 23, 40, 39) And
        PreviousSunday(New DateTime(2018, 3, 26, 2, 44, 7)) =
        New DateTime(2018, 3, 25, 2, 44, 7) And
        PreviousSunday(New DateTime(2018, 10, 29, 2, 12, 19)) =
        New DateTime(2018, 10, 28, 2, 12, 19)

      UnitTestsPassed = UnitTestsPassed And
        OffsetHoliday(New DateTime(2021, 12, 25), InvariantCulture) =
        New DateTime(2021, 12, 19) And
        OffsetHoliday(New DateTime(2020, 10, 12), InvariantCulture) =
        New DateTime(2020, 10, 12) And
        OffsetHoliday(New DateTime(2012, 1, 1), New CultureInfo("nl-NL")) =
        New DateTime(2012, 1, 1) And
        OffsetHoliday(New DateTime(2016, 3, 26), New CultureInfo("nl-NL")) =
        New DateTime(2016, 3, 26) And
        OffsetHoliday(New DateTime(2016, 3, 28), New CultureInfo("nl-NL")) =
        New DateTime(2016, 3, 27) And
        OffsetHoliday(New DateTime(2016, 3, 29), New CultureInfo("nl-NL")) =
        New DateTime(2016, 3, 29) And
        OffsetHoliday(New DateTime(2012, 4, 30), New CultureInfo("nl-NL")) =
        New DateTime(2012, 4, 29) And
        OffsetHoliday(New DateTime(2018, 4, 27, 23, 59, 59),
        New CultureInfo("nl-NL")) = New DateTime(2018, 4, 22, 23, 59, 59) And
        OffsetHoliday(New DateTime(2014, 5, 29), New CultureInfo("nl-NL")) =
        New DateTime(2014, 5, 25) And
        OffsetHoliday(New DateTime(2020, 6, 1), New CultureInfo("nl-NL")) =
        New DateTime(2020, 5, 31) And
        OffsetHoliday(New DateTime(2009, 12, 25), New CultureInfo("nl-NL")) =
        New DateTime(2009, 12, 20) And
        OffsetHoliday(New DateTime(2011, 12, 26), New CultureInfo("nl-NL")) =
        New DateTime(2011, 12, 25) And
        OffsetHoliday(New DateTime(2010, 12, 31), New CultureInfo("nl-NL")) =
        New DateTime(2010, 12, 26) And
        OffsetHoliday(New DateTime(2014, 2, 1, 0, 0, 1),
        New CultureInfo("nl-NL")) = New DateTime(2014, 2, 1, 0, 0, 1) And
        OffsetHoliday(New DateTime(2009, 2, 23), New CultureInfo("nl-NL")) =
        New DateTime(2009, 2, 23) And
        OffsetHoliday(New DateTime(2014, 5, 5), New CultureInfo("nl-NL")) =
        New DateTime(2014, 5, 5) And
        OffsetHoliday(New DateTime(2015, 5, 5), New CultureInfo("nl-NL")) =
        New DateTime(2015, 5, 3) And
        OffsetHoliday(New DateTime(2011, 4, 11), New CultureInfo("nl-NL")) =
        New DateTime(2011, 4, 11) And
        OffsetHoliday(New DateTime(2022, 4, 17), New CultureInfo("en-US")) =
        New DateTime(2022, 4, 17) And
        OffsetHoliday(New DateTime(2020, 6, 1), New CultureInfo("en-US")) =
        New DateTime(2020, 6, 1)

      For i = 0 To 19
        If i = 0 Then
          StatisticsItem = Statistics({3411, 3067, 3159, 2579, 2604, 3549,
            2028, 3521, 3629, 3418, 2091, 2828})
        ElseIf i = 1 Then
          StatisticsItem = Statistics({3725, 3581, 2747, 3924, 3743, 2112,
            3899, 2728, 3050, 3534, 2107, 3185})
        ElseIf i = 2 Then
          StatisticsItem = Statistics({2937, 2596, 3245, 3296, 2528, 2559,
            3660, 3649, 3178, 3972, 3822, 2454})
        ElseIf i = 3 Then
          StatisticsItem = Statistics({3390, 3960, 2488, 3068, 2213, 3999,
            3352, 2031, 3150, 2200, 2206, 3598})
        ElseIf i = 4 Then
          StatisticsItem = Statistics({2569, 2091, 2592, 2764, 2602, 3897,
            3960, 2803, 2557, 2321, 2326, 3293})
        ElseIf i = 5 Then
          StatisticsItem = Statistics({2820, 2826, 3425, 2652, 3266, 2415,
            2372, 3167, 2161, 2916, 3811, 2523})
        ElseIf i = 6 Then
          StatisticsItem = Statistics({3570, 2758, 2579, 3839, 3263, 3255,
            2857, 2196, 3122, 3389, 3827, 3670})
        ElseIf i = 7 Then
          StatisticsItem = Statistics({2045, 3087, 3832, 2861, 3356, 3005,
            3027, 2926, 2707, 2810, 2539, 2111})
        ElseIf i = 8 Then
          StatisticsItem = Statistics({2488, 3958, 2122, 2781, 2730, 2980,
            2311, 2949, 2515, 3258, 3084, 2313})
        ElseIf i = 9 Then
          StatisticsItem = Statistics({3877, 3309, 3012, 2781, 2215, 3568,
            2919, 3507, 3192, 3665, 2038, 2421})
        ElseIf i = 10 Then
          StatisticsItem = Statistics({2148, 2211, 2663, 2256, 2000, 3074,
            3314, 3088, 3655, 2164, 2384, 3358}, {3, 5, 10, 20, 28, 32, 41,
            46, 57, 66, 74, 76})
        ElseIf i = 11 Then
          StatisticsItem = Statistics({2908, 2714, 2300, 3409, 3858, 3060,
            2179, 3515, 2804, 2924, 2984, 2415}, {10, 18, 28, 35, 44, 50, 51,
            62, 63, 66, 74, 80})
        ElseIf i = 12 Then
          StatisticsItem = Statistics({2659, 2191, 3180, 2340, 3855, 2196,
            2888, 2546, 3745, 3501, 2546, 3347}, {9, 11, 13, 18, 21, 26, 28,
            37, 42, 47, 56, 61})
        ElseIf i = 13 Then
          StatisticsItem = Statistics({2513, 2180, 2062, 2645, 3580, 2595,
            2471, 2961, 2509, 2681, 2090, 2965}, {6, 14, 23, 26, 32, 35, 37,
            41, 45, 48, 58, 68})
        ElseIf i = 14 Then
          StatisticsItem = Statistics({2412, 3729, 3177, 3510, 3856, 2662,
            3086, 2161, 3269, 2820, 3921, 2229}, {7, 12, 23, 29, 30, 32, 36,
            38, 41, 49, 55, 61})
        ElseIf i = 15 Then
          StatisticsItem = Statistics({3847, 3240, 2695, 2298, 2960, 2439,
            3987, 2261, 2058, 2691, 3095, 3846}, {4, 9, 15, 18, 20, 26, 36,
            45, 49, 58, 64, 71})
        ElseIf i = 16 Then
          StatisticsItem = Statistics({3076, 2813, 3694, 3652, 3345, 3444,
            3994, 2680, 2990, 2826, 3391, 2358}, {7, 14, 24, 28, 37, 40, 49,
            51, 55, 61, 69, 77})
        ElseIf i = 17 Then
          StatisticsItem = Statistics({2846, 3086, 3629, 3082, 2855, 3018,
            2456, 3238, 2980, 3362, 3773, 2741}, {6, 16, 23, 29, 35, 40, 49,
            60, 64, 73, 75, 78})
        ElseIf i = 18 Then
          StatisticsItem = Statistics({2605, 2586, 2301, 3060, 2447, 3169,
            2727, 3752, 2956, 2381, 3368, 3495}, {6, 13, 24, 30, 38, 47, 57,
            59, 69, 77, 86, 96})
        ElseIf i = 19 Then
          StatisticsItem = Statistics({3228, 3564, 2323, 3616, 2405, 3914,
            2132, 2123, 3586, 2759, 2927, 2239}, {10, 15, 21, 22, 24, 32, 34,
            43, 46, 53, 55, 63})
        End If
        With StatisticsItem
          UnitTestsPassed = UnitTestsPassed And
            Round(.Slope, 4) = {-24.1399, -67.5699, 51.3427, -56.9825,
            27.3182, -2.6573, 32.1923, -46.8462, -11.1224, -61.5455, 9.4424,
            -0.1602, 10.7659, 4.8889, -4.6572, -0.4548, -6.6652, 2.2442,
            8.6539, -12.0462}(i) And
            Round(.OriginSlope, 4) = {383.2213, 397.5889, 426.4229, 371.4506,
            374.8399, 372.6621, 425.6739, 359.6522, 360.8676, 379.3893,
            52.1812, 50.6373, 75.2409, 60.0661, 73.5505, 61.2832, 59.1633,
            53.9967, 46.4227, 66.0541}(i) And
            Round(.Angle, 4) = {-87.6279, -89.1521, 88.8842, -88.9946,
            87.9036, -69.3779, 88.2208, -88.7771, -84.8624, -89.0691, 83.9546,
            -9.0998, 84.6933, 78.4399, -77.8813, -24.4582, -81.4673, 65.9827,
            83.4085, -85.2545}(i) And
            Round(.OriginAngle, 4) = {89.8505, 89.8559, 89.8656, 89.8458,
            89.8471, 89.8463, 89.8654, 89.8407, 89.8412, 89.849, 88.9021,
            88.8687, 89.2385, 89.0462, 89.221, 89.0651, 89.0317, 88.939,
            88.766, 89.1327}(i) And
            Round(.Intercept, 4) = {3123.1026, 3566.2179, 2875.6154,
            3284.6538, 2664.3333, 2877.4487, 3016.6923, 3116.4872, 2851.9231,
            3380.5, 2332.5305, 2930.2549, 2585.1154, 2427.9251, 3229.618,
            2967.1468, 3472.9635, 2986.3475, 2469.7771, 3320.9411}(i) And
            Round(.StandardError, 4) = {582.4218, 633.706, 535.0359, 720.9024,
            619.3358, 506.8629, 525.9328, 483.3154, 527.9273, 573.7699,
            544.1683, 523.5486, 590.2042, 436.8994, 644.5969, 698.2517,
            478.875, 384.1275, 419.8051, 657.4656}(i) And
            Round(.Correlation, 4) = {-0.1548, -0.374, 0.3411, -0.2864,
            0.1645, -0.0198, 0.2255, -0.3441, -0.0794, -0.3759, 0.4296,
            -0.0071, 0.3208, 0.2029, -0.1209, -0.0154, -0.3016, 0.1484,
            0.5294, -0.3121}(i) And
            Round(.ModifiedCorrelation, 4) = {0.819, 0.7932, 0.8652, 0.7909,
            0.8474, 0.8345, 0.8554, 0.8061, 0.8274, 0.7962, 0.8665, 0.9037,
            0.8892, 0.908, 0.887, 0.8274, 0.8714, 0.8915, 0.9047,
            0.8561}(i) And
            Round(.Determination, 4) = {0.024, 0.1398, 0.1164, 0.082, 0.0271,
            0.0004, 0.0509, 0.1184, 0.0063, 0.1413, 0.1845, 0.0001, 0.1029,
            0.0412, 0.0146, 0.0002, 0.091, 0.022, 0.2803, 0.0974}(i)
        End With
      Next i

      UnitTestsPassed = UnitTestsPassed And
        PIAFIntervalSeconds(-14, 3600) = 300 And
        PIAFIntervalSeconds(-13, 3600) = 300 And
        PIAFIntervalSeconds(-12, 3600) = 330 And
        PIAFIntervalSeconds(-10, 3600) = 412.5 And
        PIAFIntervalSeconds(-7, 3600) = 660 And
        PIAFIntervalSeconds(-5, 3600) = 1100 And
        PIAFIntervalSeconds(-3, 3600) = 3300 And
        PIAFIntervalSeconds(-2, 3600) = 3600 And
        PIAFIntervalSeconds(-1, 3600) = 300 And
        PIAFIntervalSeconds(0, 3600) = 300 And
        PIAFIntervalSeconds(1, 3600) = 3600 And
        PIAFIntervalSeconds(2, 3600) = 3300 And
        PIAFIntervalSeconds(3, 3600) = 1650 And
        PIAFIntervalSeconds(5, 3600) = 825 And
        PIAFIntervalSeconds(7, 3600) = 550 And
        PIAFIntervalSeconds(9, 3600) = 412.5 And
        PIAFIntervalSeconds(11, 3600) = 330 And
        PIAFIntervalSeconds(12, 3600) = 300 And
        PIAFIntervalSeconds(13, 3600) = 300 And
        PIAFIntervalSeconds(14, 3600) = 300

      UnitTestsPassed = UnitTestsPassed And
        Not PIAFShouldPrepareData(-600) And
        Not PIAFShouldPrepareData(-599) And
        Not PIAFShouldPrepareData(-1) And
        Not PIAFShouldPrepareData(0) And
        Not PIAFShouldPrepareData(1) And
        Not PIAFShouldPrepareData(299) And
        Not PIAFShouldPrepareData(300) And
        Not PIAFShouldPrepareData(301) And
        Not PIAFShouldPrepareData(599) And
        PIAFShouldPrepareData(600) And
        PIAFShouldPrepareData(601) And
        PIAFShouldPrepareData(899) And
        PIAFShouldPrepareData(900) And
        PIAFShouldPrepareData(901) And
        PIAFShouldPrepareData(1199) And
        PIAFShouldPrepareData(1200) And
        PIAFShouldPrepareData(1201) And
        PIAFShouldPrepareData(1499) And
        PIAFShouldPrepareData(1500) And
        PIAFShouldPrepareData(1501)

      Return UnitTestsPassed

    End Function


    Public Shared Function IntegrationTestsPassed As Boolean

      ' Integration tests, returns True if all tests pass.

      Dim InputPointDriver As DBMPointDriverTestModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Timestamp As DateTime
      Dim i As Integer
      Dim Result As DBMResult
      Dim DBM As New DBM

      IntegrationTestsPassed = True

      InputPointDriver = New DBMPointDriverTestModel(0)
      CorrelationPoints.Add(
        New DBMCorrelationPoint(New DBMPointDriverTestModel(490), False))
      Timestamp = New DateTime(2016, 1, 1, 0, 0, 0)

      For i = 0 To 19
        Result = DBM.Result(InputPointDriver, CorrelationPoints, Timestamp,
          New CultureInfo("nl-NL")) ' Use Dutch locale for New Year's Day test
        With Result
          IntegrationTestsPassed = IntegrationTestsPassed And
            Round(.Factor, 4) = {0, 0, 0, 0, 0, 0, -11.8493, -22.9119, 0, 0, 0,
            0, 1.1375, 0, 0, 0, 0, 0, 0, 0}(i) And
            .HasEvent = {False, False, False, False, False, False, True,
            True, False, False, False, False, True, False, False, False,
            False, False, False, False}(i) And
            .HasSuppressedEvent = {False, False, False, False, False,
            True, False, False, False, False, False, False, False, False,
            True, False, False, False, False, False}(i) And
            Round(.ForecastItem.Measurement, 4) = {527.5796, 687.0052,
            1097.1504, 950.9752, 496.1124, 673.6569, 1139.1957, 867.4313,
            504.9407, 656.4434, 1065.7651, 898.9191, 471.2433, 668.1,
            1103.9689, 897.7268, 525.3563, 676.7206, 1183.0887,
            975.8324}(i) And
            Round(.ForecastItem.ForecastValue, 4) = {517.2028, 716.9982,
            1059.0551, 919.4719, 488.6181, 683.6728, 1155.5986, 895.6872,
            503.2566, 655.7115, 1061.2282, 893.3488, 464.4957, 666.2928,
            1084.1527, 901.6546, 523.8671, 666.1729, 1190.3511,
            975.4264}(i) And
            Round(.ForecastItem.Range(0.95), 4) = {7.3871, 43.4768, 57.5299,
            51.6959, 28.3939, 0.5641, 0.9146, 0.8148, 5.6816, 7.3282, 11.0143,
            9.3649, 3.9192, 4.4945, 6.9199, 9.2247, 8.3518, 9.6103, 20.3862,
            15.9607}(i) And
            Round(.ForecastItem.Range(BandwidthCI), 4) = {11.1804, 65.8024,
            87.0718, 78.2419, 42.9743, 0.8537, 1.3843, 1.2332, 8.5991, 11.0913,
            16.6702, 14.1738, 5.9317, 6.8025, 10.4733, 13.9616, 12.6405,
            14.5453, 30.8546, 24.1566}(i) And
            Round(.ForecastItem.LowerControlLimit, 4) = {506.0224, 651.1959,
            971.9833, 841.23, 445.6438, 682.8191, 1154.2143, 894.454,
            494.6574, 644.6202, 1044.558, 879.175, 458.564, 659.4903,
            1073.6794, 887.693, 511.2267, 651.6276, 1159.4964,
            951.2698}(i) And
            Round(.ForecastItem.UpperControlLimit, 4) = {528.3832, 782.8006,
            1146.127, 997.7138, 531.5924, 684.5266, 1156.9829, 896.9205,
            511.8557, 666.8028, 1077.8984, 907.5226, 470.4274, 673.0952,
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

      Dim InputPointDriver As DBMPointDriverTestModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Timestamp, Timer As DateTime
      Dim Result As DBMResult
      Dim DBM As New DBM
      Dim Count As Integer

      InputPointDriver = New DBMPointDriverTestModel(0)
      CorrelationPoints.Add(New DBMCorrelationPoint(
        New DBMPointDriverTestModel(5394), False))
      CorrelationPoints.Add(New DBMCorrelationPoint(
        New DBMPointDriverTestModel(227), True))
      Timestamp = New DateTime(2016, 1, 1, 0, 0, 0)

      Timer = Now
      Do While Now.Ticks-Timer.Ticks < DurationTicks
        Result = DBM.Result(InputPointDriver, CorrelationPoints, Timestamp,
          New CultureInfo("nl-NL")) ' Use Dutch locale for holidays
        Count += 1
        Timestamp = Timestamp.AddSeconds(CalculationInterval)
      Loop

      Return Count/(DurationTicks/TicksPerSecond)/(24*60*60/CalculationInterval)

    End Function


  End Class


End Namespace
