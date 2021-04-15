Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBM
Imports Vitens.DynamicBandwidthMonitor.DBMAssert
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics
Imports Vitens.DynamicBandwidthMonitor.DBMTimeSeries


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTests


    Public Shared Sub RunUnitTests

      Dim i As Integer

      ' DBMAssert
      AssertEqual(True, True)
      AssertEqual(Today, Today)
      AssertEqual(NaN, NaN)
      AssertEqual(1, 1)
      AssertEqual(-2.13, -2.13)

      AssertArrayEqual({1}, {1})
      AssertArrayEqual({-1.23, 4.56}, {-1.23, 4.56})

      AssertAlmostEqual(1.231, 1.232, 2)
      AssertAlmostEqual(1.23123, 1.23124)
      AssertAlmostEqual(-2, -2)

      AssertTrue(True)

      AssertFalse(False)

      AssertNaN(NaN)

      AssertAlmostEqual(Hash({6, 4, 7, 1, 1, 4, 2, 4}), 2.2326)
      AssertAlmostEqual(Hash({8, 4, 7, 3, 2, 6, 5, 7}), 3.6609)
      AssertAlmostEqual(Hash({1, 5, 4, 7, 7, 8, 5, 1}), 1.8084)

      ' DBM
      AssertTrue(HasCorrelation(0.9, 45))
      AssertTrue(HasCorrelation(0.85, 45))
      AssertFalse(HasCorrelation(0.8, 45))
      AssertFalse(HasCorrelation(0.9, 0))
      AssertTrue(HasCorrelation(0.9, 35))
      AssertFalse(HasCorrelation(0.9, 25))
      AssertFalse(HasCorrelation(0.8, 35))
      AssertFalse(HasCorrelation(-0.9, 45))
      AssertFalse(HasCorrelation(0.9, -45))
      AssertFalse(HasCorrelation(0, 0))

      AssertTrue(HasAnticorrelation(-0.9, -45, False))
      AssertFalse(HasAnticorrelation(-0.85, -45, True))
      AssertFalse(HasAnticorrelation(-0.8, -45, False))
      AssertFalse(HasAnticorrelation(-0.9, 0, False))
      AssertTrue(HasAnticorrelation(-0.9, -35, False))
      AssertFalse(HasAnticorrelation(-0.9, -25, False))
      AssertFalse(HasAnticorrelation(-0.8, -35, False))
      AssertFalse(HasAnticorrelation(0.9, -45, False))
      AssertFalse(HasAnticorrelation(-0.9, 45, False))
      AssertFalse(HasAnticorrelation(0, 0, False))

      AssertEqual(Suppress(5, 0, 0, 0, 0, False), 5)
      AssertEqual(Suppress(5, -0.8, 0, 0, 0, False), 5)
      AssertEqual(Suppress(5, -0.9, -45, 0, 0, False), 0)
      AssertEqual(Suppress(5, -0.9, 0, 0, 0, True), 5)
      AssertEqual(Suppress(5, 0, 0, 0.8, 0, False), 5)
      AssertEqual(Suppress(5, 0, 0, 0.9, 45, False), 0)
      AssertEqual(Suppress(5, 0, 0, 0.9, 45, True), 0)
      AssertEqual(Suppress(-0.9, -0.85, 0, 0, 0, False), -0.9)
      AssertEqual(Suppress(-0.9, -0.95, -45, 0, 0, False), 0)
      AssertEqual(Suppress(0.9, 0, 0, 0.85, 45, False), 0)
      AssertEqual(Suppress(0.9, 0, 0, 0.95, 1, True), 0.9)
      AssertEqual(Suppress(-0.9, -0.99, -45, 0, 0, False), 0)
      AssertEqual(Suppress(-0.9, 0.99, 0, 0, 0, False), -0.9)
      AssertEqual(Suppress(-0.9, 0.99, 0, 0, 0, True), -0.9)
      AssertEqual(Suppress(0.99, -0.9, -45, 0, 0, False), 0)
      AssertEqual(Suppress(0.99, -0.9, 0, 0, 0, True), 0.99)
      AssertEqual(Suppress(-0.99, 0, 0, 0, 0, True), -0.99)
      AssertEqual(Suppress(0.99, 0, 0, 0, 0, True), 0.99)
      AssertEqual(Suppress(-0.98, -0.99, -1, 0, 0, False), -0.98)
      AssertEqual(Suppress(-0.98, -0.99, -45, 0, 0, False), 0)

      ' DBMMath
      AssertAlmostEqual(NormSInv(0.95), 1.6449)
      AssertAlmostEqual(NormSInv(0.99), 2.3263)
      AssertAlmostEqual(NormSInv(0.9999), 3.719)
      AssertAlmostEqual(NormSInv(0.8974), 1.2669)
      AssertAlmostEqual(NormSInv(0.7663), 0.7267)
      AssertAlmostEqual(NormSInv(0.2248), -0.7561)
      AssertAlmostEqual(NormSInv(0.9372), 1.5317)
      AssertAlmostEqual(NormSInv(0.4135), -0.2186)
      AssertAlmostEqual(NormSInv(0.2454), -0.689)
      AssertAlmostEqual(NormSInv(0.2711), -0.6095)
      AssertAlmostEqual(NormSInv(0.2287), -0.7431)
      AssertAlmostEqual(NormSInv(0.6517), 0.3899)
      AssertAlmostEqual(NormSInv(0.8663), 1.1091)
      AssertAlmostEqual(NormSInv(0.9275), 1.4574)
      AssertAlmostEqual(NormSInv(0.7089), 0.5502)
      AssertAlmostEqual(NormSInv(0.1234), -1.1582)
      AssertAlmostEqual(NormSInv(0.0837), -1.3806)
      AssertAlmostEqual(NormSInv(0.6243), 0.3168)
      AssertAlmostEqual(NormSInv(0.0353), -1.808)
      AssertAlmostEqual(NormSInv(0.9767), 1.9899)

      AssertAlmostEqual(TInv2T(0.3353, 16), 0.9934)
      AssertAlmostEqual(TInv2T(0.4792, 12), 0.7303)
      AssertAlmostEqual(TInv2T(0.4384, 9), 0.8108)
      AssertAlmostEqual(TInv2T(0.0905, 6), 2.0152)
      AssertAlmostEqual(TInv2T(0.63, 16), 0.4911)
      AssertAlmostEqual(TInv2T(0.1533, 11), 1.5339)
      AssertAlmostEqual(TInv2T(0.6297, 12), 0.4948)
      AssertAlmostEqual(TInv2T(0.1512, 4), 1.7714)
      AssertAlmostEqual(TInv2T(0.4407, 18), 0.7884)
      AssertAlmostEqual(TInv2T(0.6169, 15), 0.5108)
      AssertAlmostEqual(TInv2T(0.6077, 18), 0.5225)
      AssertAlmostEqual(TInv2T(0.4076, 20), 0.8459)
      AssertAlmostEqual(TInv2T(0.1462, 18), 1.5187)
      AssertAlmostEqual(TInv2T(0.3421, 6), 1.0315)
      AssertAlmostEqual(TInv2T(0.6566, 6), 0.4676)
      AssertAlmostEqual(TInv2T(0.2986, 1), 1.9733)
      AssertAlmostEqual(TInv2T(0.2047, 14), 1.3303)
      AssertAlmostEqual(TInv2T(0.5546, 2), 0.7035)
      AssertAlmostEqual(TInv2T(0.0862, 6), 2.0504)
      AssertAlmostEqual(TInv2T(0.6041, 10), 0.5354)

      AssertAlmostEqual(TInv(0.4097, 8), -0.2359)
      AssertAlmostEqual(TInv(0.174, 19), -0.9623)
      AssertAlmostEqual(TInv(0.6545, 15), 0.4053)
      AssertAlmostEqual(TInv(0.7876, 5), 0.8686)
      AssertAlmostEqual(TInv(0.2995, 3), -0.5861)
      AssertAlmostEqual(TInv(0.0184, 2), -5.0679)
      AssertAlmostEqual(TInv(0.892, 1), 2.8333)
      AssertAlmostEqual(TInv(0.7058, 18), 0.551)
      AssertAlmostEqual(TInv(0.3783, 2), -0.3549)
      AssertAlmostEqual(TInv(0.2774, 15), -0.6041)
      AssertAlmostEqual(TInv(0.0406, 8), -1.9945)
      AssertAlmostEqual(TInv(0.1271, 4), -1.3303)
      AssertAlmostEqual(TInv(0.241, 18), -0.718)
      AssertAlmostEqual(TInv(0.0035, 1), -90.942)
      AssertAlmostEqual(TInv(0.1646, 10), -1.0257)
      AssertAlmostEqual(TInv(0.279, 11), -0.6041)
      AssertAlmostEqual(TInv(0.8897, 4), 1.4502)
      AssertAlmostEqual(TInv(0.5809, 13), 0.2083)
      AssertAlmostEqual(TInv(0.3776, 11), -0.3197)
      AssertAlmostEqual(TInv(0.5267, 15), 0.0681)

      AssertAlmostEqual(MeanAbsoluteDeviationScaleFactor, 1.2533)

      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(1), 1)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(2), 1.2247)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(4), 1.3501)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(6), 1.3936)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(8), 1.4157)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(10), 1.429)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(12), 1.4378)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(14), 1.4442)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(16), 1.449)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(18), 1.4527)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(20), 1.4557)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(22), 1.4581)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(24), 1.4602)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(26), 1.4619)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(28), 1.4634)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(30), 1.4826)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(32), 1.4826)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(34), 1.4826)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(36), 1.4826)
      AssertAlmostEqual(MedianAbsoluteDeviationScaleFactor(38), 1.4826)

      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 1), 63.6567)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 2), 9.9248)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 4), 4.6041)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 6), 3.7074)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 8), 3.3554)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 10), 3.1693)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 12), 3.0545)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 14), 2.9768)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 16), 2.9208)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 20), 2.8453)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 22), 2.8188)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 24), 2.7969)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 28), 2.7633)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 30), 2.5758)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 32), 2.5758)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.99, 34), 2.5758)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.95, 30), 1.96)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.90, 30), 1.6449)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.95, 25), 2.0595)
      AssertAlmostEqual(ControlLimitRejectionCriterion(0.90, 20), 1.7247)

      AssertEqual(NonNaNCount({60}), 1)
      AssertEqual(NonNaNCount({60, 70}), 2)
      AssertEqual(NonNaNCount({60, 70, NaN}), 2)
      AssertEqual(NonNaNCount({60, 70, NaN, 20}), 3)
      AssertEqual(NonNaNCount({60, 70, NaN, 20, NaN}), 3)
      AssertEqual(NonNaNCount({60, 70, NaN, 20, NaN, NaN}), 3)
      AssertEqual(NonNaNCount({70, NaN, 20, NaN, NaN, 5}), 3)
      AssertEqual(NonNaNCount({NaN, 20, NaN, NaN, 5, 10}), 3)
      AssertEqual(NonNaNCount({20, NaN, NaN, 5, 10, 15}), 4)
      AssertEqual(NonNaNCount({NaN, NaN, 5, 10, 15, 20}), 4)
      AssertEqual(NonNaNCount({NaN, 5, 10, 15, 20, 25}), 5)
      AssertEqual(NonNaNCount({5, 10, 15, 20, 25, 30}), 6)
      AssertEqual(NonNaNCount({10, 15, 20, 25, 30}), 5)
      AssertEqual(NonNaNCount({15, 20, 25, 30}), 4)
      AssertEqual(NonNaNCount({20, 25, 30}), 3)
      AssertEqual(NonNaNCount({25, 30}), 2)
      AssertEqual(NonNaNCount({30}), 1)
      AssertEqual(NonNaNCount({}), 0)
      AssertEqual(NonNaNCount({NaN}), 0)
      AssertEqual(NonNaNCount({NaN, NaN}), 0)

      AssertEqual(Mean({60}), 60)
      AssertEqual(Mean({72}), 72)
      AssertEqual(Mean({32, 95}), 63.5)
      AssertEqual(Mean({81, 75}), 78)
      AssertEqual(Mean({67, 76, 25}), 56)
      AssertEqual(Mean({25, NaN, 27}), 26)
      AssertEqual(Mean({31, 73, 83, 81}), 67)
      AssertEqual(Mean({18, 58, 47, 47}), 42.5)
      AssertEqual(Mean({10, NaN, 20, 30, NaN}), 20)
      AssertEqual(Mean({67, 90, 74, 32, 62}), 65)
      AssertAlmostEqual(Mean({78, 0, 98, 65, 69, 57}), 61.1667)
      AssertEqual(Mean({49, 35, 74, 25, 28, 92}), 50.5)
      AssertAlmostEqual(Mean({7, 64, 22, 7, 42, 34, 30}), 29.4286)
      AssertAlmostEqual(Mean({13, 29, 39, 33, 96, 43, 17}), 38.5714)
      AssertEqual(Mean({59, 78, 53, 7, 18, 44, 63, 40}), 45.25)
      AssertEqual(Mean({77, 71, 99, 39, 50, 94, 67, 30}), 65.875)
      AssertEqual(Mean({91, 69, 63, 5, 44, 93, 89, 45, 50}), 61)
      AssertAlmostEqual(Mean({12, 84, 12, 94, 52, 17, 1, 13, 37}), 35.7778)
      AssertEqual(Mean({18, 14, 54, 40, 73, 77, 4, 91, 53, 10}), 43.4)
      AssertEqual(Mean({80, 30, 1, 92, 44, 61, 18, 72, 63, 41}), 50.2)

      AssertEqual(Median({57}), 57)
      AssertEqual(Median({46}), 46)
      AssertEqual(Median({79, 86}), 82.5)
      AssertEqual(Median({46, 45}), 45.5)
      AssertEqual(Median({10, NaN, 20}), 15)
      AssertEqual(Median({58, 79, 68}), 68)
      AssertEqual(Median({NaN, 10, 30, NaN}), 20)
      AssertEqual(Median({30, 10, NaN, 15}), 15)
      AssertNaN(Median({NaN, NaN, NaN, NaN, NaN}))
      AssertEqual(Median({19, 3, 9, 14, 31}), 14)
      AssertEqual(Median({2, 85, 33, 10, 38, 56}), 35.5)
      AssertEqual(Median({42, 65, 57, 92, 56, 59}), 58)
      AssertEqual(Median({53, 17, 18, 73, 34, 96, 9}), 34)
      AssertEqual(Median({23, 23, 43, 74, 8, 51, 88}), 43)
      AssertEqual(Median({72, 46, 39, 7, 83, 96, 18, 50}), 48)
      AssertEqual(Median({50, 25, 28, 99, 79, 97, 30, 16}), 40)
      AssertEqual(Median({7, 22, 58, 32, 5, 90, 46, 91, 66}), 46)
      AssertEqual(Median({81, 64, 5, 23, 48, 18, 19, 87, 15}), 23)
      AssertEqual(Median({33, 82, 42, 33, 81, 56, 13, 13, 54, 6}), 37.5)
      AssertEqual(Median({55, 40, 75, 23, 53, 85, 59, 9, 72, 44}), 54)

      AssertArrayEqual(AbsoluteDeviation({100, 44, 43, 45}, 44), {56, 0, 1, 1})
      AssertArrayEqual(AbsoluteDeviation({76, 70, 84, 39}, 77), {1, 7, 7, 38})
      AssertArrayEqual(AbsoluteDeviation(
        {15, 74, 54, 23}, 93), {78, 19, 39, 70})
      AssertArrayEqual(AbsoluteDeviation({48, 12, 85, 50}, 9), {39, 3, 76, 41})
      AssertArrayEqual(AbsoluteDeviation({7, 24, 28, 41}, 69), {62, 45, 41, 28})
      AssertArrayEqual(AbsoluteDeviation({55, 46, 81, 50}, 88), {33, 42, 7, 38})
      AssertArrayEqual(AbsoluteDeviation({84, 14, 23, 54}, 52), {32, 38, 29, 2})
      AssertArrayEqual(AbsoluteDeviation({92, 92, 45, 21}, 49), {43, 43, 4, 28})
      AssertArrayEqual(AbsoluteDeviation({3, 9, 70, 51}, 36), {33, 27, 34, 15})
      AssertArrayEqual(AbsoluteDeviation({14, 7, 8, 53}, 37), {23, 30, 29, 16})
      AssertArrayEqual(AbsoluteDeviation({66, 37, 23, 80}, 3), {63, 34, 20, 77})
      AssertArrayEqual(AbsoluteDeviation(
        {93, 45, 18, 50}, 75), {18, 30, 57, 25})
      AssertArrayEqual(AbsoluteDeviation({2, 73, 13, 0}, 37), {35, 36, 24, 37})
      AssertArrayEqual(AbsoluteDeviation({76, 77, 56, 11}, 77), {1, 0, 21, 66})
      AssertArrayEqual(AbsoluteDeviation({90, 36, 72, 54}, 81), {9, 45, 9, 27})
      AssertArrayEqual(AbsoluteDeviation({5, 62, 50, 23}, 19), {14, 43, 31, 4})
      AssertArrayEqual(AbsoluteDeviation({79, 51, 54, 24}, 60), {19, 9, 6, 36})
      AssertArrayEqual(AbsoluteDeviation({26, 53, 38, 42}, 60), {34, 7, 22, 18})
      AssertArrayEqual(AbsoluteDeviation({4, 100, 32, 23}, 32), {28, 68, 0, 9})
      AssertArrayEqual(AbsoluteDeviation(
        {15, 94, 58, 67}, 78), {63, 16, 20, 11})

      AssertEqual(MeanAbsoluteDeviation({19}), 0)
      AssertEqual(MeanAbsoluteDeviation({86}), 0)
      AssertEqual(MeanAbsoluteDeviation({7, 24}), 8.5)
      AssertEqual(MeanAbsoluteDeviation({12, 96}), 42)
      AssertAlmostEqual(MeanAbsoluteDeviation({19, 74, 70}), 23.5556)
      AssertAlmostEqual(MeanAbsoluteDeviation({96, 93, 65}), 13.1111)
      AssertEqual(MeanAbsoluteDeviation({47, 29, 24, 11}), 10.25)
      AssertEqual(MeanAbsoluteDeviation({3, 43, 53, 80}), 21.75)
      AssertEqual(MeanAbsoluteDeviation({17, 27, 98, 85, 51}), 28.72)
      AssertEqual(MeanAbsoluteDeviation({2, 82, 63, 1, 49}), 30.32)
      AssertEqual(MeanAbsoluteDeviation({9, 25, 41, 85, 82, 55}),24.5)
      AssertAlmostEqual(MeanAbsoluteDeviation({5, 74, 53, 97, 81, 21}),28.8333)
      AssertAlmostEqual(MeanAbsoluteDeviation(
        {26, 81, 9, 18, 39, 97, 21}), 27.102)
      AssertAlmostEqual(MeanAbsoluteDeviation(
        {5, 83, 31, 24, 55, 22, 87}), 26.6939)
      AssertAlmostEqual(MeanAbsoluteDeviation(
        {22, 84, 6, 79, 89, 71, 34, 56}), 25.8438)
      AssertEqual(MeanAbsoluteDeviation(
        {33, 39, 6, 88, 69, 11, 76, 65}), 26.125)
      AssertAlmostEqual(MeanAbsoluteDeviation(
        {31, 52, 12, 60, 52, 44, 47, 81, 34}), 13.9012)
      AssertAlmostEqual(MeanAbsoluteDeviation(
        {64, 63, 54, 94, 25, 80, 97, 45, 51}), 17.8519)
      AssertEqual(MeanAbsoluteDeviation(
        {47, 22, 52, 22, 10, 38, 94, 85, 54, 41}), 19.9)
      AssertEqual(MeanAbsoluteDeviation(
        {7, 12, 84, 29, 41, 8, 18, 15, 16, 84}), 22.96)

      AssertEqual(MedianAbsoluteDeviation({2}), 0)
      AssertEqual(MedianAbsoluteDeviation({37}), 0)
      AssertEqual(MedianAbsoluteDeviation({87, 37}), 25)
      AssertEqual(MedianAbsoluteDeviation({13, 74}), 30.5)
      AssertEqual(MedianAbsoluteDeviation({39, 52, 93}), 13)
      AssertEqual(MedianAbsoluteDeviation({90, 24, 47}), 23)
      AssertEqual(MedianAbsoluteDeviation({11, 51, 20, 62}), 20)
      AssertEqual(MedianAbsoluteDeviation({74, 35, 9, 95}), 30)
      AssertEqual(MedianAbsoluteDeviation({32, 46, 15, 90, 66}), 20)
      AssertEqual(MedianAbsoluteDeviation({91, 19, 50, 55, 44}), 6)
      AssertEqual(MedianAbsoluteDeviation({2, 64, 87, 65, 61, 97}), 13)
      AssertEqual(MedianAbsoluteDeviation({35, 66, 73, 74, 71, 93}), 4)
      AssertEqual(MedianAbsoluteDeviation({54, 81, 80, 36, 11, 36, 45}), 9)
      AssertEqual(MedianAbsoluteDeviation({14, 69, 40, 68, 75, 10, 69}), 7)
      AssertEqual(MedianAbsoluteDeviation(
        {40, 51, 28, 21, 91, 95, 66, 3}), 22.5)
      AssertEqual(MedianAbsoluteDeviation({57, 87, 94, 46, 51, 27, 10, 7}), 30)
      AssertEqual(MedianAbsoluteDeviation(
        {3, 89, 62, 84, 86, 37, 14, 72, 33}), 25)
      AssertEqual(MedianAbsoluteDeviation(
        {48, 6, 14, 2, 74, 89, 15, 8, 83}), 13)
      AssertEqual(MedianAbsoluteDeviation(
        {2, 22, 91, 84, 43, 96, 55, 3, 9, 11}), 26.5)
      AssertEqual(MedianAbsoluteDeviation(
        {6, 96, 82, 26, 47, 84, 34, 39, 60, 99}), 28)

      AssertFalse(UseMeanAbsoluteDeviation({6, 5, 5, 9}))
      AssertFalse(UseMeanAbsoluteDeviation({2, 2, 10, 9}))
      AssertFalse(UseMeanAbsoluteDeviation({4, 0, 10, 6}))
      AssertFalse(UseMeanAbsoluteDeviation({6, 10, 1, 1}))
      AssertTrue(UseMeanAbsoluteDeviation({3, 0, 0, 0}))
      AssertFalse(UseMeanAbsoluteDeviation({10, 0, 8, 5}))
      AssertFalse(UseMeanAbsoluteDeviation({7, 4, 5, 3}))
      AssertTrue(UseMeanAbsoluteDeviation({5, 1, 5, 5}))
      AssertFalse(UseMeanAbsoluteDeviation({2, 7, 3, 8}))
      AssertFalse(UseMeanAbsoluteDeviation({2, 6, 10, 1}))
      AssertFalse(UseMeanAbsoluteDeviation({1, 6, 3, 5}))
      AssertFalse(UseMeanAbsoluteDeviation({3, 9, 7, 3}))
      AssertTrue(UseMeanAbsoluteDeviation({5, 5, 8, 5}))
      AssertFalse(UseMeanAbsoluteDeviation({5, 10, 5, 4}))
      AssertFalse(UseMeanAbsoluteDeviation({0, 2, 4, 1}))
      AssertFalse(UseMeanAbsoluteDeviation({7, 3, 0, 10}))
      AssertTrue(UseMeanAbsoluteDeviation({4, 4, 4, 0}))
      AssertFalse(UseMeanAbsoluteDeviation({5, 7, 4, 5}))
      AssertTrue(UseMeanAbsoluteDeviation({2, 2, 2, 2}))
      AssertTrue(UseMeanAbsoluteDeviation({9, 4, 4, 4}))

      AssertEqual(CentralTendency({3, 0, 0, 0}), 0.75)
      AssertEqual(CentralTendency({2, 10, 0, 1}), 1.5)
      AssertEqual(CentralTendency({7, 7, 7, 1}), 5.5)
      AssertEqual(CentralTendency({5, 7, 2, 8}), 6)
      AssertEqual(CentralTendency({9, 3, 4, 5}), 4.5)
      AssertEqual(CentralTendency({3, 3, 3, 3}), 3)
      AssertEqual(CentralTendency({8, 4, 2, 10}), 6)
      AssertEqual(CentralTendency({2, 1, 10, 10}), 6)
      AssertEqual(CentralTendency({3, 3, 6, 2}), 3)
      AssertEqual(CentralTendency({9, 9, 6, 5}), 7.5)
      AssertEqual(CentralTendency({2, 8, 8, 9}), 8)
      AssertEqual(CentralTendency({7, 7, 4, 1}), 5.5)
      AssertEqual(CentralTendency({5, 5, 5, 0}), 3.75)
      AssertEqual(CentralTendency({4, 2, 3, 7}), 3.5)
      AssertEqual(CentralTendency({2, 1, 5, 1}), 1.5)
      AssertEqual(CentralTendency({9, 4, 5, 0}), 4.5)
      AssertEqual(CentralTendency({1, 1, 7, 1}), 2.5)
      AssertEqual(CentralTendency({1, 5, 9, 5}), 5)
      AssertEqual(CentralTendency({3, 5, 1, 9}), 4)
      AssertEqual(CentralTendency({0, 0, 0, 0}), 0)

      AssertAlmostEqual(ControlLimit({8, 2, 0, 10}, 0.99), 30.5449)
      AssertAlmostEqual(ControlLimit({2, 4, 8, 7}, 0.99), 15.2724)
      AssertAlmostEqual(ControlLimit({5, 8, 0, 2}, 0.99), 19.0906)
      AssertAlmostEqual(ControlLimit({8, 1, 0, 3}, 0.99), 11.4543)
      AssertAlmostEqual(ControlLimit({10, 7, 1, 3}, 0.99), 22.9087)
      AssertAlmostEqual(ControlLimit({6, 2, 1, 9}, 0.99), 19.0906)
      AssertAlmostEqual(ControlLimit({4, 9, 9, 3}, 0.99), 19.0906)
      AssertAlmostEqual(ControlLimit({10, 7, 2, 8}, 0.99), 11.4543)
      AssertAlmostEqual(ControlLimit({6, 0, 10, 1}, 0.99), 22.9087)
      AssertAlmostEqual(ControlLimit({10, 3, 4, 2}, 0.99), 7.6362)
      AssertAlmostEqual(ControlLimit({6, 4, 4, 4}, 0.99), 5.4904)
      AssertAlmostEqual(ControlLimit({1, 0, 9, 9}, 0.99), 30.5449)
      AssertAlmostEqual(ControlLimit({0, 3, 6, 2}, 0.99), 11.4543)
      AssertAlmostEqual(ControlLimit({9, 7, 4, 6}, 0.99), 11.4543)
      AssertAlmostEqual(ControlLimit({6, 6, 4, 1}, 0.99), 7.6362)
      AssertAlmostEqual(ControlLimit({7, 3, 4, 1}, 0.99), 11.4543)
      AssertAlmostEqual(ControlLimit({6, 4, 4, 10}, 0.99), 7.6362)
      AssertAlmostEqual(ControlLimit({10, 5, 5, 5}, 0.99), 13.7259)
      AssertAlmostEqual(ControlLimit({8, 5, 5, 5}, 0.98), 6.4023)
      AssertAlmostEqual(ControlLimit({8, 4, 0, 0}, 0.95), 8.3213)

      AssertArrayEqual(RemoveOutliers(
        {100, 100, 100, 100, 100, 100, 100, 100, 999}),
        {100, 100, 100, 100, 100, 100, 100, 100, NaN})
      AssertArrayEqual(RemoveOutliers(
        {100, 101, 102, 103, 104, 105, 106, 107, 999}),
        {100, 101, 102, 103, 104, 105, 106, 107, NaN})
      AssertArrayEqual(RemoveOutliers(
        {2223.6946, 2770.1624, 2125.7544, 3948.9927, 2184.2341, 2238.6421,
        2170.0227, 2967.0674, 2177.3738, 3617.1328, 2460.8193, 3315.8684}),
        {2223.6946, 2770.1624, 2125.7544, NaN, 2184.2341, 2238.6421,
        2170.0227, 2967.0674, 2177.3738, NaN, 2460.8193, NaN})
      AssertArrayEqual(RemoveOutliers(
        {3355.1553, 3624.3154, 3317.6895, 3610.0039, 3990.751, 2950.4382,
        2140.5908, 3237.4917, 3319.7139, 2829.2725, 3406.9199, 3230.0078}),
        {3355.1553, 3624.3154, 3317.6895, 3610.0039, 3990.751, 2950.4382,
        NaN, 3237.4917, 3319.7139, 2829.2725, 3406.9199, 3230.0078})
      AssertArrayEqual(RemoveOutliers(
        {2969.7808, 3899.0913, 2637.4045, 2718.73, 2960.9597, 2650.6521,
        2707.4294, 2034.5339, 2935.9111, 3458.7085, 2584.53, 3999.4238}),
        {2969.7808, NaN, 2637.4045, 2718.73, 2960.9597, 2650.6521,
        2707.4294, 2034.5339, 2935.9111, 3458.7085, 2584.53, NaN})
      AssertArrayEqual(RemoveOutliers(
        {2774.8018, 2755.0251, 2756.6152, 3800.0625, 2900.0671, 2784.0134,
        3955.2947, 2847.0908, 2329.7837, 3282.4614, 2597.1582, 3009.8796}),
        {2774.8018, 2755.0251, 2756.6152, NaN, 2900.0671, 2784.0134,
        NaN, 2847.0908, 2329.7837, 3282.4614, 2597.1582, 3009.8796})
      AssertArrayEqual(RemoveOutliers(
        {3084.8821, 3394.1196, 3131.3245, 2799.9587, 2528.3088, 3015.4998,
        2912.2029, 2022.2645, 3666.5674, 3685.1973, 3149.6931, 3070.0479}),
        {3084.8821, 3394.1196, 3131.3245, 2799.9587, 2528.3088, 3015.4998,
        2912.2029, NaN, 3666.5674, 3685.1973, 3149.6931, 3070.0479})
      AssertArrayEqual(RemoveOutliers(
        {3815.72, 3063.9106, 3535.0366, 2349.564, 2597.2661, 3655.3076,
        3452.7407, 2020.7682, 3810.7046, 3833.8396, 3960.6016, 3866.8149}),
        {3815.72, 3063.9106, 3535.0366, NaN, 2597.2661, 3655.3076,
        3452.7407, NaN, 3810.7046, 3833.8396, 3960.6016, 3866.8149})
      AssertArrayEqual(RemoveOutliers(
        {2812.0613, 3726.7427, 2090.9749, 2548.4485, 3900.5151, 3545.854,
        3880.2229, 3940.9585, 3942.2234, 3263.0137, 3701.8882, 2056.5291}),
        {2812.0613, 3726.7427, NaN, 2548.4485, 3900.5151, 3545.854,
        3880.2229, 3940.9585, 3942.2234, 3263.0137, 3701.8882, NaN})
      AssertArrayEqual(RemoveOutliers(
        {3798.4775, 2959.3879, 2317.3547, 2596.3599, 2075.6292, 2563.9685,
        2695.5081, 2386.2161, 2433.1106, 2810.3716, 2499.7554, 3843.103}),
        {NaN, 2959.3879, 2317.3547, 2596.3599, 2075.6292, 2563.9685,
        2695.5081, 2386.2161, 2433.1106, 2810.3716, 2499.7554, NaN})
      AssertArrayEqual(RemoveOutliers(
        {2245.7856, 2012.4834, 2473.0103, 2684.5693, 2645.4729, 2851.019,
        2344.6099, 2408.1492, 3959.5967, 3954.0583, 2399.2617, 2652.8855}),
        {2245.7856, 2012.4834, 2473.0103, 2684.5693, 2645.4729, 2851.019,
        2344.6099, 2408.1492, NaN, NaN, 2399.2617, 2652.8855})
      AssertArrayEqual(RemoveOutliers(
        {2004.5355, 2743.0693, 3260.7441, 2382.8906, 2365.9385, 2243.333,
        3506.5352, 3905.7717, 3516.5337, 2133.8328, 2308.1809, 2581.4009}),
        {2004.5355, 2743.0693, 3260.7441, 2382.8906, 2365.9385, 2243.333,
        3506.5352, NaN, 3516.5337, 2133.8328, 2308.1809, 2581.4009})
      AssertArrayEqual(RemoveOutliers(
        {3250.5376, 3411.313, 2037.264, 3709.5815, 3417.1167, 3996.0493,
        3529.637, 3992.7163, 2786.95, 3728.834, 3304.4272, 2248.9119}),
        {3250.5376, 3411.313, NaN, 3709.5815, 3417.1167, 3996.0493,
        3529.637, 3992.7163, 2786.95, 3728.834, 3304.4272, 2248.9119})
      AssertArrayEqual(RemoveOutliers(
        {2398.3125, 2742.4028, 2720.752, 2628.8442, 2750.1482, 2724.4932,
        2161.6875, 2644.4163, 2188.2952, 2455.4622, 3332.5503, 2540.5198}),
        {2398.3125, 2742.4028, 2720.752, 2628.8442, 2750.1482, 2724.4932,
        2161.6875, 2644.4163, 2188.2952, 2455.4622, NaN, 2540.5198})
      AssertArrayEqual(RemoveOutliers(
        {3991.7854, 3607.98, 2686.032, 2546.969, 3053.8796, 3138.9824,
        2441.1689, 2737.1245, 2616.7139, 2550.5774, 2406.0913, 2743.2361}),
        {NaN, 3607.98, 2686.032, 2546.969, 3053.8796, 3138.9824,
        2441.1689, 2737.1245, 2616.7139, 2550.5774, 2406.0913, 2743.2361})
      AssertArrayEqual(RemoveOutliers(
        {2361.5334, 3636.4312, 2187.593, 2281.5432, 2132.3833, 2056.792,
        2227.7795, 2757.1753, 3416.9126, 2568.927, 2094.2065, 3449.3984}),
        {2361.5334, NaN, 2187.593, 2281.5432, 2132.3833, 2056.792,
        2227.7795, 2757.1753, NaN, 2568.927, 2094.2065, NaN})
      AssertArrayEqual(RemoveOutliers(
        {2249.7119, 2411.8374, 3041.5498, 2679.1458, 2561.1577, 2405.7229,
        2775.2253, 2832.1233, 2540.2134, 3654.5903, 3970.5173, 2920.5637}),
        {2249.7119, 2411.8374, 3041.5498, 2679.1458, 2561.1577, 2405.7229,
        2775.2253, 2832.1233, 2540.2134, 3654.5903, NaN, 2920.5637})
      AssertArrayEqual(RemoveOutliers(
        {2038.1091, 2248.3057, 2427.1646, 2337.2427, 2642.043, 3497.5393,
        3996.3579, 2178.979, 3968.8848, 3460.8613, 2774.8486, 2338.1362}),
        {2038.1091, 2248.3057, 2427.1646, 2337.2427, 2642.043, 3497.5393,
        NaN, 2178.979, NaN, 3460.8613, 2774.8486, 2338.1362})
      AssertArrayEqual(RemoveOutliers(
        {3010.9485, 2517.2876, 2057.7188, 2133.0801, 3192.0308, 2035.0759,
        3821.248, 2391.8086, 2267.896, 3751.3276, 2340.9497, 2327.333}),
        {3010.9485, 2517.2876, 2057.7188, 2133.0801, 3192.0308, 2035.0759,
        NaN, 2391.8086, 2267.896, NaN, 2340.9497, 2327.333})
      AssertArrayEqual(RemoveOutliers(
        {NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN, 999}),
        {NaN, 10, NaN, 10, NaN, 10, NaN, 30, 20, NaN, NaN})

      AssertEqual(ExponentialWeights(1)(0), 1)
      AssertAlmostEqual(ExponentialWeights(2)(0), 0.25)
      AssertAlmostEqual(ExponentialWeights(3)(1), 0.2857)
      AssertAlmostEqual(ExponentialWeights(4)(2), 0.2757)
      AssertAlmostEqual(ExponentialWeights(5)(3), 0.2559)
      AssertAlmostEqual(ExponentialWeights(6)(4), 0.2353)
      AssertAlmostEqual(ExponentialWeights(7)(1), 0.0685)
      AssertAlmostEqual(ExponentialWeights(8)(3), 0.0939)
      AssertAlmostEqual(ExponentialWeights(9)(1), 0.0484)
      AssertAlmostEqual(ExponentialWeights(10)(3), 0.063)
      AssertAlmostEqual(ExponentialWeights(11)(0), 0.0311)
      AssertAlmostEqual(ExponentialWeights(12)(1), 0.0335)
      AssertAlmostEqual(ExponentialWeights(13)(3), 0.0412)
      AssertAlmostEqual(ExponentialWeights(14)(7), 0.0653)
      AssertAlmostEqual(ExponentialWeights(15)(0), 0.0223)
      AssertAlmostEqual(ExponentialWeights(16)(9), 0.0642)
      AssertAlmostEqual(ExponentialWeights(17)(4), 0.0313)
      AssertAlmostEqual(ExponentialWeights(18)(15), 0.0974)
      AssertAlmostEqual(ExponentialWeights(19)(11), 0.0553)
      AssertAlmostEqual(ExponentialWeights(20)(16), 0.0816)

      AssertEqual(ExponentialMovingAverage({70.5547}), 70.5547)
      AssertAlmostEqual(ExponentialMovingAverage({53.3424, 57.9519}), 56.7995)
      AssertAlmostEqual(ExponentialMovingAverage({28.9562, 30.1948}), 29.8852)
      AssertAlmostEqual(ExponentialMovingAverage({77.474, 1.4018}), 20.4199)
      AssertAlmostEqual(ExponentialMovingAverage(
        {76.0724, 81.449, 70.9038}), 74.6551)
      AssertAlmostEqual(ExponentialMovingAverage(
        {4.5353, 41.4033, 86.2619}), 61.7699)
      AssertAlmostEqual(ExponentialMovingAverage(
        {79.048, 37.3536, 96.1953}), 76.9338)
      AssertAlmostEqual(ExponentialMovingAverage(
        {87.1446, 5.6237, 94.9557}), 68.3164)
      AssertAlmostEqual(ExponentialMovingAverage(
        {36.4019, 52.4868, 76.7112}), 64.0315)
      AssertAlmostEqual(ExponentialMovingAverage(
        {5.3505, 59.2458, 46.87, 29.8165}), 36.959)
      AssertAlmostEqual(ExponentialMovingAverage(
        {82.9802, 82.4602, 58.9163, 98.6093}), 83.4414)
      AssertAlmostEqual(ExponentialMovingAverage(
        {53.3873, 10.637, 99.9415, 67.6176, 1.5704}), 40.2177)
      AssertAlmostEqual(ExponentialMovingAverage(
        {97.9829, 40.1374, 27.828, 16.0442, 16.2822}), 27.0999)
      AssertAlmostEqual(ExponentialMovingAverage(
        {64.6587, 41.0073, 41.2767, 71.273, 32.6206, 63.3179}), 52.9531)
      AssertAlmostEqual(ExponentialMovingAverage(
        {20.7561, 18.6014, 58.3359, 8.0715, 45.7971, 90.573}), 51.847)
      AssertAlmostEqual(ExponentialMovingAverage(
        {26.1368, 78.5212, 37.8903, 28.9665, 91.9377, 63.1742}), 60.2045)
      AssertAlmostEqual(ExponentialMovingAverage(
        {26.1368, NaN, 37.8903, 28.9665, 91.9377, 63.1742}), 58.4863)
      AssertAlmostEqual(ExponentialMovingAverage(
        {NaN, 78.5212, 37.8903, NaN, 91.9377, 63.1742}), 69.7265)
      AssertAlmostEqual(ExponentialMovingAverage(
        {26.1368, 78.5212, NaN, 28.9665, NaN, NaN}), NaN)
      AssertNaN(ExponentialMovingAverage({NaN, NaN, NaN, NaN, NaN, NaN}))

      AssertAlmostEqual(SlopeToAngle(-4.5806), -77.6849)
      AssertAlmostEqual(SlopeToAngle(-4.2541), -76.7718)
      AssertAlmostEqual(SlopeToAngle(1.7964), 60.8967)
      AssertAlmostEqual(SlopeToAngle(-3.2474), -72.8844)
      AssertAlmostEqual(SlopeToAngle(4.7917), 78.2119)
      AssertAlmostEqual(SlopeToAngle(2.1792), 65.3504)
      AssertAlmostEqual(SlopeToAngle(0.4736), 25.3422)
      AssertAlmostEqual(SlopeToAngle(-2.0963), -64.4974)
      AssertAlmostEqual(SlopeToAngle(-3.2077), -72.6851)
      AssertAlmostEqual(SlopeToAngle(-1.5425), -57.0447)
      AssertAlmostEqual(SlopeToAngle(-0.5587), -29.1921)
      AssertAlmostEqual(SlopeToAngle(1.2829), 52.0642)
      AssertAlmostEqual(SlopeToAngle(3.9501), 75.7936)
      AssertAlmostEqual(SlopeToAngle(2.5841), 68.8445)
      AssertAlmostEqual(SlopeToAngle(-3.4547), -73.8563)
      AssertAlmostEqual(SlopeToAngle(3.2931), 73.1083)
      AssertAlmostEqual(SlopeToAngle(3.2042), 72.6674)
      AssertAlmostEqual(SlopeToAngle(3.1088), 72.1687)
      AssertAlmostEqual(SlopeToAngle(-1.6831), -59.2837)
      AssertAlmostEqual(SlopeToAngle(-2.0031), -63.4704)

      ' DBMDate
      AssertEqual(PreviousInterval(
        New DateTime(2016, 4, 4, 16, 33, 2)),
        New DateTime(2016, 4, 4, 16, 30, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2015, 7, 15, 2, 29, 58)),
        New DateTime(2015, 7, 15, 2, 25, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 4, 1, 22, 5, 17)),
        New DateTime(2016, 4, 1, 22, 5, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2013, 12, 1, 21, 47, 35)),
        New DateTime(2013, 12, 1, 21, 45, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 11, 22, 0, 22, 17)),
        New DateTime(2016, 11, 22, 0, 20, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 10, 11, 19, 11, 41)),
        New DateTime(2016, 10, 11, 19, 10, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2013, 10, 26, 4, 24, 53)),
        New DateTime(2013, 10, 26, 4, 20, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2014, 5, 2, 2, 52, 41)),
        New DateTime(2014, 5, 2, 2, 50, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2014, 8, 16, 13, 11, 10)),
        New DateTime(2014, 8, 16, 13, 10, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2014, 10, 25, 8, 26, 4)),
        New DateTime(2014, 10, 25, 8, 25, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2015, 6, 2, 18, 36, 24)),
        New DateTime(2015, 6, 2, 18, 35, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 11, 21, 16, 24, 27)),
        New DateTime(2016, 11, 21, 16, 20, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2014, 4, 4, 8, 42, 10)),
        New DateTime(2014, 4, 4, 8, 40, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 2, 22, 19, 8, 41)),
        New DateTime(2016, 2, 22, 19, 5, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2015, 9, 13, 22, 48, 17)),
        New DateTime(2015, 9, 13, 22, 45, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 10, 20, 2, 47, 48)),
        New DateTime(2016, 10, 20, 2, 45, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2014, 2, 8, 23, 12, 34)),
        New DateTime(2014, 2, 8, 23, 10, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 2, 27, 23, 40, 39)),
        New DateTime(2016, 2, 27, 23, 40, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2015, 8, 26, 9, 35, 55)),
        New DateTime(2015, 8, 26, 9, 35, 0))
      AssertEqual(PreviousInterval(
        New DateTime(2016, 2, 11, 0, 44, 7)),
        New DateTime(2016, 2, 11, 0, 40, 0))

      AssertEqual(NextInterval(
        New DateTime(2016, 4, 4, 16, 33, 2)),
        New DateTime(2016, 4, 4, 16, 35, 0))
      AssertEqual(NextInterval(
        New DateTime(2015, 7, 15, 2, 29, 58)),
        New DateTime(2015, 7, 15, 2, 30, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 4, 1, 22, 5, 17)),
        New DateTime(2016, 4, 1, 22, 10, 0))
      AssertEqual(NextInterval(
        New DateTime(2013, 12, 1, 21, 47, 35)),
        New DateTime(2013, 12, 1, 21, 50, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 11, 22, 0, 22, 17)),
        New DateTime(2016, 11, 22, 0, 25, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 10, 11, 19, 11, 41)),
        New DateTime(2016, 10, 11, 19, 15, 0))
      AssertEqual(NextInterval(
        New DateTime(2013, 10, 26, 4, 24, 53)),
        New DateTime(2013, 10, 26, 4, 25, 0))
      AssertEqual(NextInterval(
        New DateTime(2014, 5, 2, 2, 52, 41)),
        New DateTime(2014, 5, 2, 2, 55, 0))
      AssertEqual(NextInterval(
        New DateTime(2014, 8, 16, 13, 11, 10)),
        New DateTime(2014, 8, 16, 13, 15, 0))
      AssertEqual(NextInterval(
        New DateTime(2014, 10, 25, 8, 26, 4)),
        New DateTime(2014, 10, 25, 8, 30, 0))
      AssertEqual(NextInterval(
        New DateTime(2015, 6, 2, 18, 36, 24)),
        New DateTime(2015, 6, 2, 18, 40, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 11, 21, 16, 24, 27)),
        New DateTime(2016, 11, 21, 16, 25, 0))
      AssertEqual(NextInterval(
        New DateTime(2014, 4, 4, 8, 42, 10)),
        New DateTime(2014, 4, 4, 8, 45, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 2, 22, 19, 8, 41)),
        New DateTime(2016, 2, 22, 19, 10, 0))
      AssertEqual(NextInterval(
        New DateTime(2015, 9, 13, 22, 48, 17)),
        New DateTime(2015, 9, 13, 22, 50, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 10, 20, 2, 47, 48)),
        New DateTime(2016, 10, 20, 2, 50, 0))
      AssertEqual(NextInterval(
        New DateTime(2014, 2, 8, 23, 12, 34)),
        New DateTime(2014, 2, 8, 23, 15, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 2, 27, 23, 40, 39)),
        New DateTime(2016, 2, 27, 23, 45, 0))
      AssertEqual(NextInterval(
        New DateTime(2015, 8, 26, 9, 35, 55)),
        New DateTime(2015, 8, 26, 9, 40, 0))
      AssertEqual(NextInterval(
        New DateTime(2016, 2, 11, 0, 44, 7)),
        New DateTime(2016, 2, 11, 0, 45, 0))

      AssertEqual(IntervalSeconds(-25, 86400)*23, 86100)
      AssertEqual(IntervalSeconds(6, 3600)*5, 3300)
      AssertEqual(IntervalSeconds(-14, 3600), 300)
      AssertEqual(IntervalSeconds(-13, 3600), 300)
      AssertEqual(IntervalSeconds(-12, 3600), 330)
      AssertEqual(IntervalSeconds(-10, 3600), 412.5)
      AssertEqual(IntervalSeconds(-7, 3600), 660)
      AssertEqual(IntervalSeconds(-3, 3600), 3300)
      AssertEqual(IntervalSeconds(-2, 3600), 3600)
      AssertEqual(IntervalSeconds(-1, 3600), 300)
      AssertEqual(IntervalSeconds(0, 3600), 300)
      AssertEqual(IntervalSeconds(1, 3600), 3600)
      AssertEqual(IntervalSeconds(2, 3600), 3300)
      AssertEqual(IntervalSeconds(3, 3600), 1650)
      AssertEqual(IntervalSeconds(5, 3600), 825)
      AssertEqual(IntervalSeconds(9, 3600), 412.5)
      AssertEqual(IntervalSeconds(11, 3600), 330)
      AssertEqual(IntervalSeconds(12, 3600), 300)
      AssertEqual(IntervalSeconds(13, 3600), 300)
      AssertEqual(IntervalSeconds(14, 3600), 300)

      AssertEqual(Computus(1864), New DateTime(1864, 3, 27))
      AssertEqual(Computus(1900), New DateTime(1900, 4, 15))
      AssertEqual(Computus(1933), New DateTime(1933, 4, 16))
      AssertEqual(Computus(1999), New DateTime(1999, 4, 4))
      AssertEqual(Computus(2001), New DateTime(2001, 4, 15))
      AssertEqual(Computus(2003), New DateTime(2003, 4, 20))
      AssertEqual(Computus(2005), New DateTime(2005, 3, 27))
      AssertEqual(Computus(2007), New DateTime(2007, 4, 8))
      AssertEqual(Computus(2013), New DateTime(2013, 3, 31))
      AssertEqual(Computus(2017), New DateTime(2017, 4, 16))
      AssertEqual(Computus(2019), New DateTime(2019, 4, 21))
      AssertEqual(Computus(2021), New DateTime(2021, 4, 4))
      AssertEqual(Computus(2023), New DateTime(2023, 4, 9))
      AssertEqual(Computus(2027), New DateTime(2027, 3, 28))
      AssertEqual(Computus(2031), New DateTime(2031, 4, 13))
      AssertEqual(Computus(2033), New DateTime(2033, 4, 17))
      AssertEqual(Computus(2037), New DateTime(2037, 4, 5))
      AssertEqual(Computus(2099), New DateTime(2099, 4, 12))
      AssertEqual(Computus(2172), New DateTime(2172, 4, 12))
      AssertEqual(Computus(2292), New DateTime(2292, 4, 10))

      AssertTrue(IsHoliday(New DateTime(2021, 12, 25), InvariantCulture))
      AssertFalse(IsHoliday(New DateTime(2020, 10, 12), InvariantCulture))
      AssertTrue(IsHoliday(New DateTime(2012, 1, 1), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(
        New DateTime(2016, 3, 26), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2016, 3, 28), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(
        New DateTime(2016, 3, 29), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2012, 4, 30), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(
        New DateTime(2018, 4, 27, 23, 59, 59), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2014, 5, 29), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2020, 6, 1), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(
        New DateTime(2009, 12, 25), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(
        New DateTime(2011, 12, 26), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(
        New DateTime(2010, 12, 31), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(
        New DateTime(2014, 2, 1, 0, 0, 1), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(
        New DateTime(2009, 2, 23), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(New DateTime(2014, 5, 5), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2015, 5, 5), New CultureInfo("nl-NL")))
      AssertFalse(IsHoliday(
        New DateTime(2011, 4, 11), New CultureInfo("nl-NL")))
      AssertTrue(IsHoliday(New DateTime(2022, 4, 17), New CultureInfo("en-US")))
      AssertFalse(IsHoliday(New DateTime(2020, 6, 1), New CultureInfo("en-US")))

      AssertEqual(DaysSinceSunday(New DateTime(2016, 4, 4, 16, 33, 2)), 1)
      AssertEqual(DaysSinceSunday(New DateTime(2015, 7, 15, 2, 29, 58)), 3)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 4, 1, 22, 5, 17)), 5)
      AssertEqual(DaysSinceSunday(New DateTime(2013, 12, 1, 21, 47, 35)), 0)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 11, 22, 0, 22, 17)), 2)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 10, 11, 19, 11, 41)), 2)
      AssertEqual(DaysSinceSunday(New DateTime(2013, 10, 26, 4, 24, 53)), 6)
      AssertEqual(DaysSinceSunday(New DateTime(2014, 5, 2, 2, 52, 41)), 5)
      AssertEqual(DaysSinceSunday(New DateTime(2014, 8, 16, 13, 11, 10)), 6)
      AssertEqual(DaysSinceSunday(New DateTime(2014, 10, 25, 8, 26, 4)), 6)
      AssertEqual(DaysSinceSunday(New DateTime(2015, 6, 2, 18, 36, 24)), 2)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 11, 21, 16, 24, 27)), 1)
      AssertEqual(DaysSinceSunday(New DateTime(2014, 4, 4, 8, 42, 10)), 5)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 2, 22, 19, 8, 41)), 1)
      AssertEqual(DaysSinceSunday(New DateTime(2015, 9, 13, 22, 48, 17)), 0)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 10, 20, 2, 47, 48)), 4)
      AssertEqual(DaysSinceSunday(New DateTime(2014, 2, 8, 23, 12, 34)), 6)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 2, 27, 23, 40, 39)), 6)
      AssertEqual(DaysSinceSunday(New DateTime(2015, 8, 26, 9, 35, 55)), 3)
      AssertEqual(DaysSinceSunday(New DateTime(2016, 2, 11, 0, 44, 7)), 4)

      AssertEqual(PreviousSunday(
        New DateTime(2016, 4, 4, 16, 33, 2)),
        New DateTime(2016, 4, 3, 16, 33, 2))
      AssertEqual(PreviousSunday(
        New DateTime(2015, 7, 15, 2, 29, 58)),
        New DateTime(2015, 7, 12, 2, 29, 58))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 4, 1, 22, 5, 17)),
        New DateTime(2016, 3, 27, 22, 5, 17))
      AssertEqual(PreviousSunday(
        New DateTime(2013, 12, 1, 21, 47, 35)),
        New DateTime(2013, 12, 1, 21, 47, 35))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 11, 22, 0, 22, 17)),
        New DateTime(2016, 11, 20, 0, 22, 17))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 10, 11, 19, 11, 41)),
        New DateTime(2016, 10, 9, 19, 11, 41))
      AssertEqual(PreviousSunday(
        New DateTime(2013, 10, 26, 4, 24, 53)),
        New DateTime(2013, 10, 20, 4, 24, 53))
      AssertEqual(PreviousSunday(
        New DateTime(2014, 5, 2, 2, 52, 41)),
        New DateTime(2014, 4, 27, 2, 52, 41))
      AssertEqual(PreviousSunday(
        New DateTime(2014, 8, 16, 13, 11, 10)),
        New DateTime(2014, 8, 10, 13, 11, 10))
      AssertEqual(PreviousSunday(
        New DateTime(2014, 10, 25, 8, 26, 4)),
        New DateTime(2014, 10, 19, 8, 26, 4))
      AssertEqual(PreviousSunday(
        New DateTime(2015, 6, 2, 18, 36, 24)),
        New DateTime(2015, 5, 31, 18, 36, 24))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 11, 21, 16, 24, 27)),
        New DateTime(2016, 11, 20, 16, 24, 27))
      AssertEqual(PreviousSunday(
        New DateTime(2014, 4, 4, 8, 42, 10)),
        New DateTime(2014, 3, 30, 8, 42, 10))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 2, 22, 19, 8, 41)),
        New DateTime(2016, 2, 21, 19, 8, 41))
      AssertEqual(PreviousSunday(
        New DateTime(2015, 9, 13, 22, 48, 17)),
        New DateTime(2015, 9, 13, 22, 48, 17))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 10, 20, 2, 47, 48)),
        New DateTime(2016, 10, 16, 2, 47, 48))
      AssertEqual(PreviousSunday(
        New DateTime(2014, 2, 8, 23, 12, 34)),
        New DateTime(2014, 2, 2, 23, 12, 34))
      AssertEqual(PreviousSunday(
        New DateTime(2016, 2, 27, 23, 40, 39)),
        New DateTime(2016, 2, 21, 23, 40, 39))
      AssertEqual(PreviousSunday(
        New DateTime(2018, 3, 26, 2, 44, 7)),
        New DateTime(2018, 3, 25, 2, 44, 7))
      AssertEqual(PreviousSunday(
        New DateTime(2018, 10, 29, 2, 12, 19)),
        New DateTime(2018, 10, 28, 2, 12, 19))

      AssertEqual(OffsetHoliday(
        New DateTime(2021, 12, 25), InvariantCulture),
        New DateTime(2021, 12, 19))
      AssertEqual(OffsetHoliday(
        New DateTime(2020, 10, 12), InvariantCulture),
        New DateTime(2020, 10, 12))
      AssertEqual(OffsetHoliday(
        New DateTime(2012, 1, 1), New CultureInfo("nl-NL")),
        New DateTime(2012, 1, 1))
      AssertEqual(OffsetHoliday(
        New DateTime(2016, 3, 26), New CultureInfo("nl-NL")),
        New DateTime(2016, 3, 26))
      AssertEqual(OffsetHoliday(
        New DateTime(2016, 3, 28), New CultureInfo("nl-NL")),
        New DateTime(2016, 3, 27))
      AssertEqual(OffsetHoliday(
        New DateTime(2016, 3, 29), New CultureInfo("nl-NL")),
        New DateTime(2016, 3, 29))
      AssertEqual(OffsetHoliday(
        New DateTime(2012, 4, 30), New CultureInfo("nl-NL")),
        New DateTime(2012, 4, 29))
      AssertEqual(OffsetHoliday(
        New DateTime(2018, 4, 27, 23, 59, 59), New CultureInfo("nl-NL")),
        New DateTime(2018, 4, 22, 23, 59, 59))
      AssertEqual(OffsetHoliday(
        New DateTime(2014, 5, 29), New CultureInfo("nl-NL")),
        New DateTime(2014, 5, 25))
      AssertEqual(OffsetHoliday(
        New DateTime(2020, 6, 1), New CultureInfo("nl-NL")),
        New DateTime(2020, 5, 31))
      AssertEqual(OffsetHoliday(
        New DateTime(2009, 12, 25), New CultureInfo("nl-NL")),
        New DateTime(2009, 12, 20))
      AssertEqual(OffsetHoliday(
        New DateTime(2011, 12, 26), New CultureInfo("nl-NL")),
        New DateTime(2011, 12, 25))
      AssertEqual(OffsetHoliday(
        New DateTime(2010, 12, 31), New CultureInfo("nl-NL")),
        New DateTime(2010, 12, 26))
      AssertEqual(OffsetHoliday(
        New DateTime(2014, 2, 1, 0, 0, 1), New CultureInfo("nl-NL")),
        New DateTime(2014, 2, 1, 0, 0, 1))
      AssertEqual(OffsetHoliday(
        New DateTime(2009, 2, 23), New CultureInfo("nl-NL")),
        New DateTime(2009, 2, 23))
      AssertEqual(OffsetHoliday(
        New DateTime(2014, 5, 5), New CultureInfo("nl-NL")),
        New DateTime(2014, 5, 5))
      AssertEqual(OffsetHoliday(
        New DateTime(2015, 5, 5), New CultureInfo("nl-NL")),
        New DateTime(2015, 5, 3))
      AssertEqual(OffsetHoliday(
        New DateTime(2011, 4, 11), New CultureInfo("nl-NL")),
        New DateTime(2011, 4, 11))
      AssertEqual(OffsetHoliday(
        New DateTime(2022, 4, 17), New CultureInfo("en-US")),
        New DateTime(2022, 4, 17))
      AssertEqual(OffsetHoliday(
        New DateTime(2020, 6, 1), New CultureInfo("en-US")),
        New DateTime(2020, 6, 1))

      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 3, 2, 11, 25, 1)),
        New DateTime(2016, 12, 4, 9, 5, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 12, 7, 6, 35, 59)),
        New DateTime(2017, 9, 10, 4, 15, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2018, 3, 24, 14, 13, 10)),
        New DateTime(2017, 12, 24, 11, 50, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 8, 12, 23, 20, 36)),
        New DateTime(2016, 5, 15, 21, 0, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 9, 2, 11, 44, 11)),
        New DateTime(2017, 6, 4, 9, 20, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2018, 6, 4, 16, 59, 55)),
        New DateTime(2018, 3, 11, 14, 35, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 9, 5, 4, 38, 50)),
        New DateTime(2016, 6, 12, 2, 15, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 8, 15, 4, 11, 4)),
        New DateTime(2016, 5, 22, 1, 50, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2019, 2, 19, 9, 41, 5)),
        New DateTime(2018, 11, 25, 7, 20, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 8, 16, 16, 53, 5)),
        New DateTime(2017, 5, 21, 14, 30, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2019, 12, 21, 7, 34, 23)),
        New DateTime(2019, 9, 22, 5, 10, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2018, 7, 4, 19, 9, 44)),
        New DateTime(2018, 4, 8, 16, 45, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 1, 11, 8, 24, 29)),
        New DateTime(2016, 10, 16, 6, 0, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2018, 11, 17, 13, 41, 35)),
        New DateTime(2018, 8, 19, 11, 20, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2019, 3, 8, 15, 56, 31)),
        New DateTime(2018, 12, 9, 13, 35, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 8, 7, 17, 22, 45)),
        New DateTime(2016, 5, 15, 15, 0, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 8, 20, 17, 49, 15)),
        New DateTime(2017, 5, 28, 15, 25, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 2, 7, 21, 55, 52)),
        New DateTime(2015, 11, 15, 19, 35, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2017, 11, 22, 17, 4, 58)),
        New DateTime(2017, 8, 27, 14, 40, 0))
      AssertEqual(DataPreparationTimestamp(
        New DateTime(2016, 1, 6, 22, 49, 9)),
        New DateTime(2015, 10, 11, 20, 25, 0))

      AssertEqual(EMATimeOffset(1), 0)
      AssertEqual(EMATimeOffset(2), 0)
      AssertEqual(EMATimeOffset(3), -300)
      AssertEqual(EMATimeOffset(5), -300)
      AssertEqual(EMATimeOffset(6), -600)
      AssertEqual(EMATimeOffset(EMAPreviousPeriods+1), -600)
      AssertEqual(EMATimeOffset(8), -600)
      AssertEqual(EMATimeOffset(9), -900)
      AssertEqual(EMATimeOffset(11), -900)
      AssertEqual(EMATimeOffset(12), -1200)
      AssertEqual(EMATimeOffset(14), -1200)
      AssertEqual(EMATimeOffset(15), -1500)
      AssertEqual(EMATimeOffset(17), -1500)
      AssertEqual(EMATimeOffset(18), -1800)
      AssertEqual(EMATimeOffset(20), -1800)
      AssertEqual(EMATimeOffset(152), -15600)
      AssertEqual(EMATimeOffset(747), -76800)
      AssertEqual(EMATimeOffset(826), -84900)
      AssertEqual(EMATimeOffset(38000), -3915600)
      AssertEqual(EMATimeOffset(56000), -5770200)

      ' DBMTimeSeries
      AssertAlmostEqual(TimeWeightedValue(37.8202, 35.3707,
        New DateTime(2021, 4, 2, 21, 52, 48),
        New DateTime(2021, 4, 2, 21, 51, 16), False), -0.039)
      AssertAlmostEqual(TimeWeightedValue(89.065, 88.0274,
        New DateTime(2021, 4, 2, 19, 35, 35),
        New DateTime(2021, 4, 2, 19, 25, 36), True), -0.6175)
      AssertEqual(TimeWeightedValue(56.6827, 51.6966,
        New DateTime(2021, 4, 2, 15, 46, 26),
        New DateTime(2021, 4, 2, 15, 46, 26), False), 0)
      AssertEqual(TimeWeightedValue(1.4532, 3.5828,
        New DateTime(2021, 4, 2, 13, 37, 33),
        New DateTime(2021, 4, 2, 13, 37, 33)), 0)
      AssertAlmostEqual(TimeWeightedValue(54.5572, 59.5128,
        New DateTime(2021, 4, 2, 16, 24, 10),
        New DateTime(2021, 4, 2, 16, 30, 29), False), 0.2502)
      AssertAlmostEqual(TimeWeightedValue(61.101, 60.7455,
        New DateTime(2021, 4, 2, 21, 52, 40),
        New DateTime(2021, 4, 2, 22, 1, 50), True), 0.389)
      AssertAlmostEqual(TimeWeightedValue(92.1337, 91.216,
        New DateTime(2021, 4, 2, 11, 21, 43),
        New DateTime(2021, 4, 2, 11, 21, 47), True), 0.0043)
      AssertAlmostEqual(TimeWeightedValue(8.4389, 11.2189,
        New DateTime(2021, 4, 2, 10, 7, 1),
        New DateTime(2021, 4, 2, 10, 7, 42), False), 0.0047)
      AssertAlmostEqual(TimeWeightedValue(31.9645, 31.3657,
        New DateTime(2021, 4, 2, 23, 20, 9),
        New DateTime(2021, 4, 2, 23, 34, 2)), 0.3082)
      AssertAlmostEqual(TimeWeightedValue(79.3012, 79.5086,
        New DateTime(2021, 4, 2, 20, 5, 30),
        New DateTime(2021, 4, 2, 20, 15, 1), False), 0.5248)
      AssertAlmostEqual(TimeWeightedValue(12.7809, 16.1845,
        New DateTime(2021, 4, 2, 4, 38, 10),
        New DateTime(2021, 4, 2, 4, 49, 19), True), 0.099)
      AssertAlmostEqual(TimeWeightedValue(84.7342, 83.7452,
        New DateTime(2021, 4, 2, 20, 1, 2),
        New DateTime(2021, 4, 2, 20, 8, 48), False), 0.4543)
      AssertAlmostEqual(TimeWeightedValue(61.349, 62.735,
        New DateTime(2021, 4, 2, 18, 46, 33),
        New DateTime(2021, 4, 2, 18, 49, 9), True), 0.1108)
      AssertAlmostEqual(TimeWeightedValue(24.2132, 23.4243,
        New DateTime(2021, 4, 2, 12, 4, 8),
        New DateTime(2021, 4, 2, 12, 8, 45), False), 0.0764)
      AssertAlmostEqual(TimeWeightedValue(65.4336, 65.7781,
        New DateTime(2021, 4, 2, 17, 28, 20),
        New DateTime(2021, 4, 2, 17, 29, 28)), 0.0515)
      AssertAlmostEqual(TimeWeightedValue(35.2388, 37.7602,
        New DateTime(2021, 4, 2, 21, 42, 57),
        New DateTime(2021, 4, 2, 21, 56, 19), False), 0.3388)
      AssertAlmostEqual(TimeWeightedValue(58.8315, 55.2946,
        New DateTime(2021, 4, 2, 22, 51, 18),
        New DateTime(2021, 4, 2, 22, 56, 14), True), 0.2016)
      AssertAlmostEqual(TimeWeightedValue(34.7707, 27.6712,
        New DateTime(2021, 4, 2, 11, 25, 40),
        New DateTime(2021, 4, 2, 11, 34, 54)), 0.223)
      AssertAlmostEqual(TimeWeightedValue(37.8202, 35.3707,
        New DateTime(2021, 4, 2, 21, 51, 16),
        New DateTime(2021, 4, 2, 21, 52, 48), False), 0.039)
      AssertAlmostEqual(TimeWeightedValue(89.065, 88.0274,
        New DateTime(2021, 4, 2, 19, 25, 36),
        New DateTime(2021, 4, 2, 19, 35, 35)), 0.6175)

      ' DBMStatistics
      For i = 0 To 19
        With {
          Statistics({3411, 3067, 3159, 2579, 2604, 3549, 2028, 3521, 3629,
            3418, 2091, 2828}),
          Statistics({3725, 3581, 2747, 3924, 3743, 2112, 3899, 2728, 3050,
            3534, 2107, 3185}),
          Statistics({2937, 2596, 3245, 3296, 2528, 2559, 3660, 3649, 3178,
            3972, 3822, 2454}),
          Statistics({3390, 3960, 2488, 3068, 2213, 3999, 3352, 2031, 3150,
            2200, 2206, 3598}),
          Statistics({2569, 2091, 2592, 2764, 2602, 3897, 3960, 2803, 2557,
            2321, 2326, 3293}),
          Statistics({2820, 2826, 3425, 2652, 3266, 2415, 2372, 3167, 2161,
            2916, 3811, 2523}),
          Statistics({3570, 2758, 2579, 3839, 3263, 3255, 2857, 2196, 3122,
            3389, 3827, 3670}),
          Statistics({2045, 3087, 3832, 2861, 3356, 3005, 3027, 2926, 2707,
            2810, 2539, 2111}),
          Statistics({2488, 3958, 2122, 2781, 2730, 2980, 2311, 2949, 2515,
            3258, 3084, 2313}),
          Statistics({3877, 3309, 3012, 2781, 2215, 3568, 2919, 3507, 3192,
            3665, 2038, 2421}),
          Statistics({2148, 2211, 2663, 2256, 2000, 3074, 3314, 3088, 3655,
            2164, 2384, 3358}, {3, 5, 10, 20, 28, 32, 41, 46, 57, 66, 74, 76}),
          Statistics({2908, 2714, 2300, 3409, 3858, 3060, 2179, 3515, 2804,
            2924, 2984, 2415}, {9, 18, 28, 35, 44, 50, 51, 62, 63, 66, 74, 80}),
          Statistics({2659, 2191, 3180, 2340, 3855, 2196, 2888, 2546, 3745,
            3501, 2546, 3347}, {9, 11, 13, 18, 21, 26, 28, 37, 42, 47, 56, 61}),
          Statistics({2513, 2180, 2062, 2645, 3580, 2595, 2471, 2961, 2509,
            2681, 2090, 2965}, {6, 14, 23, 26, 32, 35, 37, 41, 45, 48, 58, 68}),
          Statistics({2412, 3729, 3177, 3510, 3856, 2662, 3086, 2161, 3269,
            2820, 3921, 2229}, {7, 12, 23, 29, 30, 32, 36, 38, 41, 49, 55, 61}),
          Statistics({3847, 3240, 2695, 2298, 2960, 2439, 3987, 2261, 2058,
            2691, 3095, 3846}, {4, 9, 15, 18, 20, 26, 36, 45, 49, 58, 64, 71}),
          Statistics({3076, 2813, 3694, 3652, 3345, 3444, 3994, 2680, 2990,
            2826, 3391, 2358}, {7, 14, 24, 28, 37, 40, 49, 51, 55, 61, 69, 77}),
          Statistics({2846, 3086, 3629, 3082, 2855, 3018, 2456, 3238, 2980,
            3362, 3773, 2741}, {6, 16, 23, 29, 35, 40, 49, 60, 64, 73, 75, 78}),
          Statistics({2605, 2586, 2301, 3060, 2447, 3169, 2727, 3752, 2956,
            2381, 3368, 3495}, {6, 13, 24, 30, 38, 47, 57, 59, 69, 77, 86, 96}),
          Statistics({3228, 3564, 2323, 3616, 2405, 3914, 2132, 2123, 3586,
            2759, 2927, 2239}, {10, 15, 21, 22, 24, 32, 34, 43, 46, 53, 55, 63})
          }(i)
          AssertAlmostEqual(.Mean, {5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5, 5.5,
            5.5, 38.1667, 48.3333, 30.75, 36.0833, 34.4167, 34.5833, 42.6667,
            45.6667, 50.1667, 34.8333}(i))
          AssertAlmostEqual(.NMBE, {542.697, 579.8333, 573.1818, 539.2273,
            510.7424, 519.5152, 579.6818, 518.7879, 506.4091, 552.0909, 69.5568,
            59.4655, 93.8347, 71.1755, 88.1816, 84.3422, 73.7324, 66.6387,
            56.8854, 82.2919}(i))
          AssertAlmostEqual(.RMSD, {3033.058, 3249.7593, 3194.8412, 3044.4737,
            2866.8538, 2894.582, 3225.9939, 2891.9635, 2826.9462, 3088.8983,
            2709.0779, 2913.7367, 2939.9523, 2599.8648, 3092.6607, 2985.8331,
            3180.1242, 3063.4419, 2887.1264, 2936.4299}(i))
          AssertAlmostEqual(.CVRMSD, {551.4651, 590.8653, 580.8802, 553.5407,
            521.2461, 526.2876, 586.5444, 525.8115, 513.9902, 561.6179, 70.9802,
            60.2842, 95.6082, 72.0517, 89.8594, 86.3373, 74.5342, 67.0827,
            57.5507, 84.2994}(i))
          AssertAlmostEqual(.Slope, {-24.1399, -67.5699, 51.3427, -56.9825,
            27.3182, -2.6573, 32.1923, -46.8462, -11.1224, -61.5455, 9.4424,
            -0.1552, 10.7659, 4.8889, -4.6572, -0.4548, -6.6652, 2.2442,
            8.6539, -12.0462}(i))
          AssertAlmostEqual(.OriginSlope, {383.2213, 397.5889, 426.4229,
            371.4506, 374.8399, 372.6621, 425.6739, 359.6522, 360.8676,
            379.3893, 52.1812, 50.5792, 75.2409, 60.0661, 73.5505, 61.2832,
            59.1633, 53.9967, 46.4227, 66.0541}(i))
          AssertAlmostEqual(.Angle, {-87.6279, -89.1521, 88.8842, -88.9946,
            87.9036, -69.3779, 88.2208, -88.7771, -84.8624, -89.0691, 83.9546,
            -8.8239, 84.6933, 78.4399, -77.8813, -24.4582, -81.4673, 65.9827,
            83.4085, -85.2545}(i))
          AssertAlmostEqual(.OriginAngle, {89.8505, 89.8559, 89.8656, 89.8458,
            89.8471, 89.8463, 89.8654, 89.8407, 89.8412, 89.849, 88.9021,
            88.8674, 89.2385, 89.0462, 89.221, 89.0651, 89.0317, 88.939,
            88.766, 89.1327}(i))
          AssertAlmostEqual(.Intercept, {3123.1026, 3566.2179, 2875.6154,
            3284.6538, 2664.3333, 2877.4487, 3016.6923, 3116.4872, 2851.9231,
            3380.5, 2332.5305, 2930.0031, 2585.1154, 2427.9251, 3229.618,
            2967.1468, 3472.9635, 2986.3475, 2469.7771, 3320.9411}(i))
          AssertAlmostEqual(.StandardError, {582.4218, 633.706, 535.0359,
            720.9024, 619.3358, 506.8629, 525.9328, 483.3154, 527.9273,
            573.7699, 544.1683, 523.5492, 590.2042, 436.8994, 644.5969,
            698.2517, 478.875, 384.1275, 419.8051, 657.4656}(i))
          AssertAlmostEqual(.Correlation, {-0.1548, -0.374, 0.3411, -0.2864,
            0.1645, -0.0198, 0.2255, -0.3441, -0.0794, -0.3759, 0.4296,
            -0.0069, 0.3208, 0.2029, -0.1209, -0.0154, -0.3016, 0.1484,
            0.5294, -0.3121}(i))
          AssertAlmostEqual(.ModifiedCorrelation, {0.819, 0.7932, 0.8652,
            0.7909, 0.8474, 0.8345, 0.8554, 0.8061, 0.8274, 0.7962, 0.8665,
            0.9024, 0.8892, 0.908, 0.887, 0.8274, 0.8714, 0.8915, 0.9047,
            0.8561}(i))
          AssertAlmostEqual(.Determination, {0.024, 0.1398, 0.1164, 0.082,
            0.0271, 0.0004, 0.0509, 0.1184, 0.0063, 0.1413, 0.1845, 0, 0.1029,
            0.0412, 0.0146, 0.0002, 0.091, 0.022, 0.2803, 0.0974}(i))
          AssertEqual(.Calibrated, {False, False, False, False, False, False,
            False, False, False, False, False, False, False, False, False,
            False, False, False, False, False}(i))
          AssertAlmostEqual(.SystematicError, {542.697, 579.8333, 573.1818,
            539.2273, 510.7424, 519.5152, 579.6818, 518.7879, 506.4091,
            552.0909, 69.5568, 59.4655, 93.8347, 71.1755, 88.1816, 84.3422,
            73.7324, 66.6387, 56.8854, 82.2919}(i))
          AssertAlmostEqual(.RandomError, {8.7681, 11.032, 7.6984, 14.3134,
            10.5037, 6.7725, 6.8625, 7.0237, 7.5811, 9.527, 1.4234, 0.8187,
            1.7735, 0.8762, 1.6778, 1.9952, 0.8017, 0.444, 0.6653, 2.0076}(i))
          AssertAlmostEqual(.Fit, {0.024, 0.1398, 0.1164, 0.082, 0.0271, 0.0004,
            0.0509, 0.1184, 0.0063, 0.1413, 0.1845, 0, 0.1029, 0.0412, 0.0146,
            0.0002, 0.091, 0.022, 0.2803, 0.0974}(i))
        End With
      Next i

    End Sub


    Public Shared Sub RunIntegrationTests

      Dim InputPointDriver As DBMPointDriverTestModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim DBM As New DBM
      Dim i As Integer
      Dim Timestamp As DateTime
      Dim Result As DBMResult

      InputPointDriver = New DBMPointDriverTestModel(0)
      CorrelationPoints.Add(
        New DBMCorrelationPoint(New DBMPointDriverTestModel(490), False))

      ' GetResult Timestamp - test alignment
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0)).Timestamp,
        New DateTime(2016, 1, 1, 0, 0, 0))
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 1)).Timestamp,
        New DateTime(2016, 1, 1, 0, 0, 0))
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 4, 59)).Timestamp,
        New DateTime(2016, 1, 1, 0, 0, 0))
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 5, 0)).Timestamp,
        New DateTime(2016, 1, 1, 0, 5, 0))
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 7, 12)).Timestamp,
        New DateTime(2016, 1, 1, 0, 5, 0))

      ' GetResult IsFutureData - test future data
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.SnapshotTimestamp.AddSeconds(
        -2*CalculationInterval)).IsFutureData, False)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.SnapshotTimestamp.AddSeconds(
        -CalculationInterval)).IsFutureData, True)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.SnapshotTimestamp).IsFutureData, True)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.SnapshotTimestamp.AddSeconds(
        CalculationInterval)).IsFutureData, True)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.SnapshotTimestamp.AddSeconds(
        2*CalculationInterval)).IsFutureData, True)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.CalculationTimestamp.AddSeconds(
        -2*CalculationInterval)).IsFutureData, False)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.CalculationTimestamp.AddSeconds(
        -CalculationInterval)).IsFutureData, False)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.CalculationTimestamp).IsFutureData, False)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.CalculationTimestamp.AddSeconds(
        CalculationInterval)).IsFutureData, True)
      AssertEqual(DBM.GetResult(InputPointDriver, CorrelationPoints,
        InputPointDriver.CalculationTimestamp.AddSeconds(
        2*CalculationInterval)).IsFutureData, True)

      ' GetResult - test calculation results
      For i = -19 To 19
        Timestamp = New DateTime(2016, 1, 1, 0, 0, 0).
          AddSeconds(Abs(i)*365*24*60*60/20) ' 19 backwards to 0, forwards to 19
        Result = DBM.GetResult(InputPointDriver, CorrelationPoints, Timestamp,
          New CultureInfo("nl-NL")) ' Use Dutch locale for New Year's Day test
        With Result
          AssertAlmostEqual(.Factor, {0, 0, 0, 0, 0, 0, -11.8493, -22.9119, 0,
            0, 0, 0, 1.1375, 0, 0, 0, 0, 0, 0, 0}(Abs(i)))
          If {6, 7, 12}.Contains(Abs(i)) Then
            AssertTrue(.HasEvent)
          Else
            AssertFalse(.HasEvent)
          End If
          If {5, 14}.Contains(Abs(i)) Then
            AssertTrue(.HasSuppressedEvent)
          Else
            AssertFalse(.HasSuppressedEvent)
          End If
          AssertAlmostEqual(.ForecastItem.Measurement, {511.0947, 714.6129,
            1093.1833, 949.5972, 508.9843, 700.7281, 1135.0766, 866.1744,
            505.9721, 682.8229, 1061.9115, 897.6165, 471.1027, 694.9479,
            1099.9772, 896.426, 529.3383, 703.9149, 1178.8108,
            974.4184}(Abs(i)))
          AssertAlmostEqual(.ForecastItem.Forecast, {520.9651, 745.8112,
            1055.2258, 918.1396, 501.8384, 711.1466, 1151.4202, 894.3894,
            504.2845, 682.0616, 1057.391, 892.0543, 464.3571, 693.0681,
            1080.2326, 900.3481, 527.8378, 692.9434, 1186.047, 974.013}(Abs(i)))
          AssertAlmostEqual(.ForecastItem.Range(0.95), {7.7131, 45.224,
            57.3219, 51.6209, 28.7259, 0.5868, 0.9113, 0.8136, 5.6932, 7.6227,
            10.9745, 9.3513, 3.918, 4.6751, 6.8949, 9.2113, 8.4151, 9.9965,
            20.3125, 15.9376}(Abs(i)))
          AssertAlmostEqual(.ForecastItem.Range(BandwidthCI), {11.6737, 68.4467,
            86.757, 78.1285, 43.4768, 0.8881, 1.3793, 1.2315, 8.6167, 11.537,
            16.6099, 14.1533, 5.93, 7.0758, 10.4354, 13.9414, 12.7363, 15.1298,
            30.743, 24.1216}(Abs(i)))
          AssertAlmostEqual(.ForecastItem.LowerControlLimit, {509.2914,
            677.3645, 968.4688, 840.0111, 458.3617, 710.2585, 1150.0409,
            893.1579, 495.6678, 670.5246, 1040.7811, 877.9011, 458.4272,
            685.9923, 1069.7972, 886.4067, 515.1015, 677.8135, 1155.3039,
            949.8915}(Abs(i)))
          AssertAlmostEqual(.ForecastItem.UpperControlLimit, {532.6389,
            814.2578, 1141.9828, 996.2681, 545.3152, 712.0346, 1152.7995,
            895.6208, 512.9012, 693.5986, 1074.0009, 906.2076, 470.2871,
            700.1439, 1090.668, 914.2895, 540.5741, 708.0732, 1216.79,
            998.1346}(Abs(i)))
        End With
      Next i

      ' GetResults Count - test number of results
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 0, 0, 0)).Count, 0)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 0, 5, 0)).Count, 1)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 0, 6, 12)).Count, 1)
      AssertEqual(DBM.GetResults(InputPointDriver, Nothing,
        New DateTime(2016, 1, 1, 0, 3, 55),
        New DateTime(2016, 1, 1, 0, 8, 55)).Count, 1)
      AssertEqual(DBM.GetResults(InputPointDriver, Nothing,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 0, 10, 0)).Count, 2)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 2, 41),
        New DateTime(2016, 1, 1, 0, 10, 0)).Count, 2)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 1, 9),
        New DateTime(2016, 1, 1, 0, 14, 57)).Count, 2)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 1, 0, 0)).Count, 12)
      AssertEqual(DBM.GetResults(InputPointDriver, Nothing,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 12, 0, 0)).Count, 144)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0)).Count, 288)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 0).Count, 288)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 1).Count, 1)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 13).Count, 13)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 24).Count, 24)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), -2).Count, 1)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), -1).Count, 288)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 287).Count, 287)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 288).Count, 288)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 2, 0, 0, 0), 289).Count, 288)
      AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1, 0, 0, 0),
        New DateTime(2016, 1, 1, 0, 0, 0), 1000).Count, 0)

      ' GetResults Timestamp - test intervals
      For i = 0 To 4
        AssertEqual(DBM.GetResults(InputPointDriver, CorrelationPoints,
          New DateTime(2016, 1, 1, 0, 0, 0),
          New DateTime(2016, 1, 1, 1, 0, 0), 5)(i).Timestamp,
          New DateTime(2016, 1, 1, 0, {0, 10, 25, 40, 55}(i), 0))
      Next i

    End Sub


    Public Shared Function RunQualityTests As String

      Dim InputPointDriver As DBMPointDriverTestModel
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim DBM As New DBM

      InputPointDriver = New DBMPointDriverTestModel(0)
      CorrelationPoints.Add(
        New DBMCorrelationPoint(New DBMPointDriverTestModel(490), False))

      Return Statistics(DBM.GetResults(InputPointDriver, CorrelationPoints,
        New DateTime(2016, 1, 1), New DateTime(2017, 1, 1))).Brief

    End Function


  End Class


End Namespace
