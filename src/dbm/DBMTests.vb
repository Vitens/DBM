Option Explicit
Option Strict


' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
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
Imports System.Environment
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBM
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMStatistics


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMTests


    ' This class contains unit and integration tests. For integration tests, an
    ' internal data array in the DBMPointDriverTest driver is used by
    ' DBMDataManager.


    Private Shared DBM As New DBM


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


    Private Shared Function TestCase(InputPoint As String,
      CorrelationPoint As String, Timestamp As DateTime, _
      SubtractSelf As Boolean, ExpResults() As Double, _
      ExpAbsErrStats() As Double, ExpRelErrStats() As Double, _
      ExpAbsErrs() As Double, ExpRelErrs() As Double, _
      ExpCorrAbsErrs() As Double, ExpCorrRelErrs() As Double) As Boolean

      ' Runs a complete test case for integration testing. The internal data
      ' array is used as input for tags, so order of tests is important and
      ' can not be changed.

      Dim InputPointDriver, CorrelationPointDriver As DBMPointDriverTest
      Dim CorrelationPoints As New List(Of DBMCorrelationPoint)
      Dim Result As DBMResult

      InputPointDriver = New DBMPointDriverTest(InputPoint)
      If CorrelationPoint IsNot Nothing Then
        CorrelationPointDriver = New DBMPointDriverTest(CorrelationPoint)
        CorrelationPoints.Add _
          (New DBMCorrelationPoint(CorrelationPointDriver, SubtractSelf))
      Else
        CorrelationPointDriver = Nothing
      End If
      Result = DBM.Result(InputPointDriver, CorrelationPoints, Timestamp)

      With Result
        Return Hash({Round(.Factor, 3), Round(.OriginalFactor, 3), _
          Round(.PredictionData.MeasuredValue), _
          Round(.PredictionData.PredictedValue), _
          Round(.PredictionData.LowerControlLimit), _
          Round(.PredictionData.UpperControlLimit)}) = Hash(ExpResults) And _
          Hash({.AbsoluteErrorStatsData.Count, _
          Round(.AbsoluteErrorStatsData.Slope, 4), _
          Round(.AbsoluteErrorStatsData.OriginSlope, 4), _
          Round(.AbsoluteErrorStatsData.Angle, 4), _
          Round(.AbsoluteErrorStatsData.OriginAngle, 4), _
          Round(.AbsoluteErrorStatsData.Intercept, 4), _
          Round(.AbsoluteErrorStatsData.StandardError, 4), _
          Round(.AbsoluteErrorStatsData.Correlation, 4), _
          Round(.AbsoluteErrorStatsData.ModifiedCorrelation, 4), _
          Round(.AbsoluteErrorStatsData.Determination, 4)}) = _
          Hash(ExpAbsErrStats) And _
          Hash({.RelativeErrorStatsData.Count, _
          Round(.RelativeErrorStatsData.Slope, 4), _
          Round(.RelativeErrorStatsData.OriginSlope, 4), _
          Round(.RelativeErrorStatsData.Angle, 4), _
          Round(.RelativeErrorStatsData.OriginAngle, 4), _
          Round(.RelativeErrorStatsData.Intercept, 4), _
          Round(.RelativeErrorStatsData.StandardError, 4), _
          Round(.RelativeErrorStatsData.Correlation, 4), _
          Round(.RelativeErrorStatsData.ModifiedCorrelation, 4), _
          Round(.RelativeErrorStatsData.Determination, 4)}) = _
          Hash(ExpRelErrStats) And _
          Hash(.AbsoluteErrors, 4) = Hash(ExpAbsErrs) And _
          Hash(.RelativeErrors, 4) = Hash(ExpRelErrs) And _
          Hash(.CorrelationAbsoluteErrors, 4) = Hash(ExpCorrAbsErrs) And _
          Hash(.CorrelationRelativeErrors, 4) = Hash(ExpCorrRelErrs) And _
          (.SuppressedBy Is Nothing Or (.SuppressedBy IsNot Nothing And _
          CorrelationPointDriver Is .SuppressedBy))
      End With

    End Function


    Private Shared Function UnitTestResults As Boolean

      ' Unit tests, returns True if all tests pass.

      Dim StatisticsData As New DBMStatisticsData
      Dim i As Integer

      UnitTestResults = True

      UnitTestResults = UnitTestResults And _
        Round(Hash({6, 4, 7, 1, 1, 4, 2, 4}), 4) = 2.2326 And _
        Round(Hash({8, 4, 7, 3, 2, 6, 5, 7}), 4) = 3.6609 And _
        Round(Hash({1, 5, 4, 7, 7, 8, 5, 1}), 4) = 1.8084

      UnitTestResults = UnitTestResults And _
        Suppress(5, 0, 0, 0, 0, False) = 5 And _
        Suppress(5, -0.8, 0, 0, 0, False) = 5 And _
        Suppress(5, -0.9, -45, 0, 0, False) = -0.9 And _
        Suppress(5, -0.9, 0, 0, 0, True) = 5 And _
        Suppress(5, 0, 0, 0.8, 0, False) = 5 And _
        Suppress(5, 0, 0, 0.9, 45, False) = 0.9 And _
        Suppress(5, 0, 0, 0.9, 45, True) = 0.9 And _
        Suppress(-0.9, -0.85, 0, 0, 0, False) = -0.9 And _
        Suppress(-0.9, -0.95, -45, 0, 0, False) = -0.95 And _
        Suppress(0.9, 0, 0, 0.85, 45, False) = 0.9 And _
        Suppress(0.9, 0, 0, 0.95, 1, True) = 0.9 And _
        Suppress(-0.9, -0.99, -45, 0, 0, False) = -0.99 And _
        Suppress(-0.9, 0.99, 0, 0, 0, False) = -0.9 And _
        Suppress(-0.9, 0.99, 0, 0, 0, True) = -0.9 And _
        Suppress(0.99, -0.9, -45, 0, 0, False) = -0.9 And _
        Suppress(0.99, -0.9, 0, 0, 0, True) = 0.99 And _
        Suppress(-0.99, 0, 0, 0, 0, True) = -0.99 And _
        Suppress(0.99, 0, 0, 0, 0, True) = 0.99 And _
        Suppress(-0.98, -0.99, -1, 0, 0, False) = -0.98 And _
        Suppress(-0.98, -0.99, -45, 0, 0, False) = -0.99

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
        Round(MeanAbsoluteDeviationScaleFactor, 4) = 1.2533

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
        Median({57}) = 57 And _
        Median({46}) = 46 And _
        Median({79, 86}) = 82.5 And _
        Median({46, 45}) = 45.5 And _
        Median({10, NaN, 20}) = 15 And _
        Median({58, 79, 68}) = 68 And _
        Median({NaN, 10, 30, NaN}) = 20 And _
        Median({30, 10, NaN, 15}) = 15 And _
        Median({99, 10, 41, 40, 55}) = 41 And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      UnitTestResults = UnitTestResults And _
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

      For i = 0 To 19
        If i = 0 Then
          StatisticsData = Calculate({3411, 3067, 3159, 2579, 2604, 3549, _
            2028, 3521, 3629, 3418, 2091, 2828})
        ElseIf i = 1 Then
          StatisticsData = Calculate({3725, 3581, 2747, 3924, 3743, 2112, _
            3899, 2728, 3050, 3534, 2107, 3185})
        ElseIf i = 2 Then
          StatisticsData = Calculate({2937, 2596, 3245, 3296, 2528, 2559, _
            3660, 3649, 3178, 3972, 3822, 2454})
        ElseIf i = 3 Then
          StatisticsData = Calculate({3390, 3960, 2488, 3068, 2213, 3999, _
            3352, 2031, 3150, 2200, 2206, 3598})
        ElseIf i = 4 Then
          StatisticsData = Calculate({2569, 2091, 2592, 2764, 2602, 3897, _
            3960, 2803, 2557, 2321, 2326, 3293})
        ElseIf i = 5 Then
          StatisticsData = Calculate({2820, 2826, 3425, 2652, 3266, 2415, _
            2372, 3167, 2161, 2916, 3811, 2523})
        ElseIf i = 6 Then
          StatisticsData = Calculate({3570, 2758, 2579, 3839, 3263, 3255, _
            2857, 2196, 3122, 3389, 3827, 3670})
        ElseIf i = 7 Then
          StatisticsData = Calculate({2045, 3087, 3832, 2861, 3356, 3005, _
            3027, 2926, 2707, 2810, 2539, 2111})
        ElseIf i = 8 Then
          StatisticsData = Calculate({2488, 3958, 2122, 2781, 2730, 2980, _
            2311, 2949, 2515, 3258, 3084, 2313})
        ElseIf i = 9 Then
          StatisticsData = Calculate({3877, 3309, 3012, 2781, 2215, 3568, _
            2919, 3507, 3192, 3665, 2038, 2421})
        ElseIf i = 10 Then
          StatisticsData = Calculate({2148, 2211, 2663, 2256, 2000, 3074, _
            3314, 3088, 3655, 2164, 2384, 3358}, {3, 5, 10, 20, 28, 32, 41, _
            46, 57, 66, 74, 76})
        ElseIf i = 11 Then
          StatisticsData = Calculate({2908, 2714, 2300, 3409, 3858, 3060, _
            2179, 3515, 2804, 2924, 2984, 2415}, {10, 18, 28, 35, 44, 50, 51, _
            62, 63, 66, 74, 80})
        ElseIf i = 12 Then
          StatisticsData = Calculate({2659, 2191, 3180, 2340, 3855, 2196, _
            2888, 2546, 3745, 3501, 2546, 3347}, {9, 11, 13, 18, 21, 26, 28, _
            37, 42, 47, 56, 61})
        ElseIf i = 13 Then
          StatisticsData = Calculate({2513, 2180, 2062, 2645, 3580, 2595, _
            2471, 2961, 2509, 2681, 2090, 2965}, {6, 14, 23, 26, 32, 35, 37, _
            41, 45, 48, 58, 68})
        ElseIf i = 14 Then
          StatisticsData = Calculate({2412, 3729, 3177, 3510, 3856, 2662, _
            3086, 2161, 3269, 2820, 3921, 2229}, {7, 12, 23, 29, 30, 32, 36, _
            38, 41, 49, 55, 61})
        ElseIf i = 15 Then
          StatisticsData = Calculate({3847, 3240, 2695, 2298, 2960, 2439, _
            3987, 2261, 2058, 2691, 3095, 3846}, {4, 9, 15, 18, 20, 26, 36, _
            45, 49, 58, 64, 71})
        ElseIf i = 16 Then
          StatisticsData = Calculate({3076, 2813, 3694, 3652, 3345, 3444, _
            3994, 2680, 2990, 2826, 3391, 2358}, {7, 14, 24, 28, 37, 40, 49, _
            51, 55, 61, 69, 77})
        ElseIf i = 17 Then
          StatisticsData = Calculate({2846, 3086, 3629, 3082, 2855, 3018, _
            2456, 3238, 2980, 3362, 3773, 2741}, {6, 16, 23, 29, 35, 40, 49, _
            60, 64, 73, 75, 78})
        ElseIf i = 18 Then
          StatisticsData = Calculate({2605, 2586, 2301, 3060, 2447, 3169, _
            2727, 3752, 2956, 2381, 3368, 3495}, {6, 13, 24, 30, 38, 47, 57, _
            59, 69, 77, 86, 96})
        ElseIf i = 19 Then
          StatisticsData = Calculate({3228, 3564, 2323, 3616, 2405, 3914, _
            2132, 2123, 3586, 2759, 2927, 2239}, {10, 15, 21, 22, 24, 32, 34, _
            43, 46, 53, 55, 63})
        End If
        UnitTestResults = UnitTestResults And _
          Round(StatisticsData.Slope, 4) = {-24.1399, -67.5699, 51.3427, _
          -56.9825, 27.3182, -2.6573, 32.1923, -46.8462, -11.1224, -61.5455, _
          9.4424, -0.1602, 10.7659, 4.8889, -4.6572, -0.4548, -6.6652, 2.2442, _
          8.6539, -12.0462}(i) And _
          Round(StatisticsData.OriginSlope, 4) = {383.2213, 397.5889, _
          426.4229, 371.4506, 374.8399, 372.6621, 425.6739, 359.6522, _
          360.8676, 379.3893, 52.1812, 50.6373, 75.2409, 60.0661, 73.5505, _
          61.2832, 59.1633, 53.9967, 46.4227, 66.0541}(i) And _
          Round(StatisticsData.Angle, 4) = {-87.6279, -89.1521, 88.8842, _
          -88.9946, 87.9036, -69.3779, 88.2208, -88.7771, -84.8624, -89.0691, _
          83.9546, -9.0998, 84.6933, 78.4399, -77.8813, -24.4582, -81.4673, _
          65.9827, 83.4085, -85.2545}(i) And _
          Round(StatisticsData.OriginAngle, 4) = {89.8505, 89.8559, 89.8656, _
          89.8458, 89.8471, 89.8463, 89.8654, 89.8407, 89.8412, 89.849, _
          88.9021, 88.8687, 89.2385, 89.0462, 89.221, 89.0651, 89.0317, _
          88.939, 88.766, 89.1327}(i) And _
          Round(StatisticsData.Intercept, 4) = {3123.1026, 3566.2179, _
          2875.6154, 3284.6538, 2664.3333, 2877.4487, 3016.6923, 3116.4872, _
          2851.9231, 3380.5, 2332.5305, 2930.2549, 2585.1154, 2427.9251, _
          3229.618, 2967.1468, 3472.9635, 2986.3475, 2469.7771, _
          3320.9411}(i) And _
          Round(StatisticsData.StandardError, 4) = {582.4218, 633.706, _
          535.0359, 720.9024, 619.3358, 506.8629, 525.9328, 483.3154, _
          527.9273, 573.7699, 544.1683, 523.5486, 590.2042, 436.8994, _
          644.5969, 698.2517, 478.875, 384.1275, 419.8051, 657.4656}(i) And _
          Round(StatisticsData.Correlation, 4) = {-0.1548, -0.374, 0.3411, _
          -0.2864, 0.1645, -0.0198, 0.2255, -0.3441, -0.0794, -0.3759, _
          0.4296, -0.0071, 0.3208, 0.2029, -0.1209, -0.0154, -0.3016, _
          0.1484, 0.5294, -0.3121}(i) And _
          Round(StatisticsData.ModifiedCorrelation, 4) = {0.819, 0.7932, _
          0.8652, 0.7909, 0.8474, 0.8345, 0.8554, 0.8061, 0.8274, 0.7962, _
          0.8665, 0.9037, 0.8892, 0.908, 0.887, 0.8274, 0.8714, 0.8915, _
          0.9047, 0.8561}(i) And _
          Round(StatisticsData.Determination, 4) = {0.024, 0.1398, 0.1164, _
          0.082, 0.0271, 0.0004, 0.0509, 0.1184, 0.0063, 0.1413, 0.1845, _
          0.0001, 0.1029, 0.0412, 0.0146, 0.0002, 0.091, 0.022, 0.2803, _
          0.0974}(i)
      Next i

      Return UnitTestResults

    End Function


    Private Shared Function IntegrationTestResults As Boolean

      ' Integration tests, returns True if all tests pass.

      Dim i As Integer

      IntegrationTestResults = True
      For i = 0 To 1 ' Run all cases twice to test cache

        ' Normal situation; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2016, 2, 5, 10, 0, 0), False, {0, 0, 1154, _
          1192, 1023, 1361}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 38.5067}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.0334}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Cache test 1/2, Normal situation; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2016, 2, 5, 10, 0, 0), False, {0, 0, 1154, _
          1192, 1023, 1361}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 38.5067}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.0334}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Cache test 2/2, Normal situation; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2016, 2, 5, 10, 5, 0), False, {0, 0, 1141, _
          1194, 1021, 1367}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 53.2608}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.0467}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Normal situation; Vitens
        IntegrationTestResults = IntegrationTestResults And TestCase("B", _
          Nothing, New DateTime(2015, 5, 23, 2, 0, 0), False, {0, 0, 27087, _
          27408, 25605, 29212}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 321.5827}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.0119}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Cache test, Normal situation; Vitens
        IntegrationTestResults = IntegrationTestResults And TestCase("B", _
          Nothing, New DateTime(2015, 5, 23, 2, 0, 0), False, {0, 0, 27087, _
          27408, 25605, 29212}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 321.5827}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.0119}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' New year's day; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2016, 1, 1, 1, 50, 0), False, {1.206, 1.206, _
          570, 378, 219, 537}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -192.0268}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.3368}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' New year's day (correlation relative error); Leeuwarden/Franeker
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "C", _
          New DateTime(2016, 1, 1, 1, 50, 0), False, {0.849, 1.206, 570, 378, _
          219, 537}, {24, 0.5621, 0.4322, 29.3416, 23.372, 28.3683, 28.2221, _
          0.911, 0.858, 0.8299}, {24, 0.9813, 0.7481, 44.459, 36.8017, 0.0877, _
          0.0829, 0.9129, 0.8487, 0.8333}, {-40.5195, -30.2586, 7.956, _
          -12.169, -0.8141, 13.2748, 86.7014, 60.1825, 33.7038, 1.2135, _
          -4.2378, -8.9689, -69.7832, -70.1429, -126.9404, -125.648, _
          -187.9416, -172.3536, -225.4554, -189.0826, -223.691, -255.7146, _
          -279.9642, -192.0268}, {-0.062, -0.0474, 0.0132, -0.0212, -0.0015, _
          0.0254, 0.1913, 0.1281, 0.0709, 0.0025, -0.0086, -0.0182, -0.1288, _
          -0.1302, -0.2172, -0.2232, -0.3131, -0.2984, -0.3699, -0.327, _
          -0.361, -0.4108, -0.4425, -0.3368}, {34.4414, 51.6882, 46.1136, _
          71.0202, 77.3139, 69.9561, 65.194, 57.2358, 26.6788, 8.4616, 9.8408, _
          -10.6877, -41.394, -63.8991, -67.9372, -56.5465, -69.2763, -75.7582, _
          -86.5879, -91.6995, -97.3343, -98.9255, -105.1413, -103.3049}, _
          {0.0889, 0.1454, 0.1306, 0.2212, 0.2535, 0.2316, 0.2225, 0.194, _
          0.0867, 0.0263, 0.0313, -0.0338, -0.1233, -0.1893, -0.1989, -0.1613, _
          -0.1963, -0.2119, -0.2429, -0.2548, -0.2697, -0.2793, -0.296, _
          -0.2956})

        ' New year's day (self-correlation); Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "A", _
          New DateTime(2016, 1, 1, 1, 50, 0), False, {1, 1.206, 570, 378, 219, _
          537}, {24, 1, 1, 45, 45, 0, 0, 1, 1, 1}, {24, 1, 1, 45, 45, 0, 0, 1, _
          1, 1}, {-40.5195, -30.2586, 7.956, -12.169, -0.8141, 13.2748, _
          86.7014, 60.1825, 33.7038, 1.2135, -4.2378, -8.9689, -69.7832, _
          -70.1429, -126.9404, -125.648, -187.9416, -172.3536, -225.4554, _
          -189.0826, -223.691, -255.7146, -279.9642, -192.0268}, {-0.062, _
          -0.0474, 0.0132, -0.0212, -0.0015, 0.0254, 0.1913, 0.1281, 0.0709, _
          0.0025, -0.0086, -0.0182, -0.1288, -0.1302, -0.2172, -0.2232, _
          -0.3131, -0.2984, -0.3699, -0.327, -0.361, -0.4108, -0.4425, _
          -0.3368}, {-40.5195, -30.2586, 7.956, -12.169, -0.8141, 13.2748, _
          86.7014, 60.1825, 33.7038, 1.2135, -4.2378, -8.9689, -69.7832, _
          -70.1429, -126.9404, -125.648, -187.9416, -172.3536, -225.4554, _
          -189.0826, -223.691, -255.7146, -279.9642, -192.0268}, {-0.062, _
          -0.0474, 0.0132, -0.0212, -0.0015, 0.0254, 0.1913, 0.1281, 0.0709, _
          0.0025, -0.0086, -0.0182, -0.1288, -0.1302, -0.2172, -0.2232, _
          -0.3131, -0.2984, -0.3699, -0.327, -0.361, -0.4108, -0.4425, -0.3368})

        ' Subtract self test, New year's day; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "A", _
          New DateTime(2016, 1, 1, 1, 50, 0), True, {1.206, 1.206, 570, 378, _
          219, 537}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0}, {-40.5195, -30.2586, 7.956, -12.169, -0.8141, 13.2748, _
          86.7014, 60.1825, 33.7038, 1.2135, -4.2378, -8.9689, -69.7832, _
          -70.1429, -126.9404, -125.648, -187.9416, -172.3536, -225.4554, _
          -189.0826, -223.691, -255.7146, -279.9642, -192.0268}, {-0.062, _
          -0.0474, 0.0132, -0.0212, -0.0015, 0.0254, 0.1913, 0.1281, 0.0709, _
          0.0025, -0.0086, -0.0182, -0.1288, -0.1302, -0.2172, -0.2232, _
          -0.3131, -0.2984, -0.3699, -0.327, -0.361, -0.4108, -0.4425, _
          -0.3368}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0})

        ' Unmeasured flow; Heerenveen
        IntegrationTestResults = IntegrationTestResults And TestCase("D", _
          Nothing, New DateTime(2015, 8, 21, 1, 0, 0), False, {1.511, 1.511, _
          846, 281, -93, 655}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -565.4391}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6681}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Unmeasured flow (anticorrelation absolute error); Heerenveen/Joure
        IntegrationTestResults = IntegrationTestResults And TestCase("D", "E", _
          New DateTime(2015, 8, 21, 1, 0, 0), False, {-0.996, 1.511, 846, 281, _
          -93, 655}, {24, -0.7001, -0.8208, -34.9959, -39.3783, 61.8779, _
          35.2858, -0.8933, -0.9963, 0.798}, {24, -13.9813, 5.0997, -85.9089, _
          78.9056, -11.7537, 2.0865, -0.3963, 0.8028, 0.157}, {-485.1482, _
          -440.9432, -418.6948, -403.7726, -396.7766, -405.2911, -401.7344, _
          -376.5941, -354.1613, -375.2625, -394.8583, -409.7765, -490.4449, _
          -579.3504, -560.444, -547.4837, -576.4284, -573.7847, -604.1606, _
          -621.2496, -637.0153, -631.5339, -609.5431, -565.4391}, {-0.6325, _
          -0.6043, -0.5926, -0.5816, -0.5789, -0.5774, -0.5581, -0.526, _
          -0.4979, -0.5098, -0.5208, -0.5195, -0.604, -0.6916, -0.6512, _
          -0.6278, -0.6526, -0.6454, -0.6741, -0.6744, -0.6707, -0.6884, _
          -0.6864, -0.6681}, {350.6993, 341.6965, 350.8028, 324.9046, _
          288.6392, 302.3464, 334.5285, 344.8929, 349.464, 367.892, 367.1969, _
          392.318, 393.9427, 413.3453, 432.2467, 441.4588, 445.0821, 467.6547, _
          470.2933, 506.1637, 514.1507, 523.476, 528.9035, 536.1065}, _
          {-4.1426, -4.0553, -4.1104, -5.856, -11.5241, -6.3002, -4.2634, _
          -3.6096, -3.4004, -2.9626, -2.5859, -2.3803, -2.2672, -2.1796, _
          -2.1255, -2.0988, -2.0868, -2.0207, -1.9427, -1.7458, -1.5747, _
          -1.4795, -1.4114, -1.3614})

        ' Pipe burst; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2013, 3, 12, 19, 30, 0), False, {6.793, 6.793, _
          3075, 1140, 855, 1425}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, -1935.136}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6293}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Cache test, Pipe burst; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2013, 3, 12, 19, 35, 0), False, {5.818, 5.818, _
          3042, 1148, 822, 1473}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, -1894.5257}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6227}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Pipe burst 1/2; Leeuwarden/Vitens
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "B", _
          New DateTime(2013, 3, 12, 19, 30, 0), True, {6.793, 6.793, 3075, _
          1140, 855, 1425}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {48.8165, 94.2071, 111.1325, 126.9718, 103.2957, _
          95.8641, 145.2011, 124.2651, 125.8249, 89.8765, 98.8689, 102.9246, _
          66.0764, 69.4908, 53.0277, 50.7359, 61.2197, -204.6909, -1180.754, _
          -1620.7954, -1684.7705, -1798.2238, -1907.0149, -1935.136}, {0.0533, _
          0.1035, 0.1249, 0.1419, 0.1118, 0.1017, 0.1539, 0.1296, 0.1303, _
          0.0881, 0.0961, 0.1003, 0.0614, 0.0639, 0.0473, 0.0444, 0.0536, _
          -0.1456, -0.499, -0.5811, -0.5863, -0.6081, -0.6241, -0.6293}, {0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0})

        ' Pipe burst 2/2; Leeuwarden/Vitens
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "B", _
          New DateTime(2013, 3, 12, 20, 10, 0), True, {6.717, 6.717, 2936, _
          953, 658, 1248}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {125.8249, 89.8765, 98.8689, 102.9246, 66.0764, _
          69.4908, 53.0277, 50.7359, 61.2197, -204.6909, -1180.754, _
          -1620.7954, -1684.7705, -1798.2238, -1907.0149, -1935.136, _
          -1894.5257, -1800.1785, -1843.1128, -1802.7745, -1831.5277, _
          -1916.506, -1953.4516, -1982.9866}, {0.1303, 0.0881, 0.0961, 0.1003, _
          0.0614, 0.0639, 0.0473, 0.0444, 0.0536, -0.1456, -0.499, -0.5811, _
          -0.5863, -0.6081, -0.6241, -0.6293, -0.6227, -0.613, -0.6275, _
          -0.6219, -0.6337, -0.6512, -0.6663, -0.6754}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Pipe burst; Sneek
        IntegrationTestResults = IntegrationTestResults And TestCase("F", _
          Nothing, New DateTime(2016, 1, 4, 4, 0, 0), False, {3.206, 3.206, _
          1320, 390, 100, 680}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, -929.6622}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.7043}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Pipe burst; Sneek/Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("F", "A", _
          New DateTime(2016, 1, 4, 4, 0, 0), False, {3.206, 3.206, 1320, 390, _
          100, 680}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0}, {83.1233, 46.3106, -227.4276, -471.2181, -722.395, -820.036, _
          -907.1308, -929.6377, -990.2518, -965.0819, -972.8458, -914.8979, _
          -899.4269, -900.1418, -954.8946, -935.7147, -953.9033, -933.8995, _
          -950.9154, -958.2613, -954.5203, -976.0492, -946.4839, -929.6622}, _
          {0.2437, 0.1294, -0.3641, -0.5398, -0.6633, -0.6938, -0.7248, _
          -0.7105, -0.7355, -0.7188, -0.7302, -0.7118, -0.6951, -0.6921, _
          -0.7315, -0.7148, -0.7272, -0.7108, -0.7204, -0.726, -0.7231, _
          -0.7394, -0.717, -0.7043}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Pipe burst; Arnhem
        IntegrationTestResults = IntegrationTestResults And TestCase("G", _
          Nothing, New DateTime(2015, 11, 25, 10, 0, 0), False, {15.629, _
          15.629, 3180, 984, 844, 1125}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -2195.2182}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6904}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Pipe burst; Arnhem/Druten
        IntegrationTestResults = IntegrationTestResults And TestCase("G", "H", _
          New DateTime(2015, 11, 25, 10, 0, 0), False, {15.629, 15.629, 3180, _
          984, 844, 1125}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {-1778.7731, -2069.9348, -2245.1073, -2410.9526, _
          -2516.8254, -2569.3224, -2492.0598, -2427.9686, -2394.7571, _
          -2395.1749, -2392.6236, -2388.8825, -2381.0622, -2365.6884, _
          -2356.9603, -2317.8985, -2290.885, -2290.5544, -2248.7011, _
          -2260.2408, -2270.5524, -2293.4979, -2256.7542, -2195.2182}, _
          {-0.6332, -0.667, -0.6825, -0.6983, -0.7039, -0.7054, -0.696, _
          -0.691, -0.6903, -0.6942, -0.6964, -0.7001, -0.701, -0.7005, _
          -0.7007, -0.6972, -0.6953, -0.6945, -0.6881, -0.6908, -0.6934, _
          -0.6972, -0.6949, -0.6904}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Maintenance; Drachten
        IntegrationTestResults = IntegrationTestResults And TestCase("I", _
          Nothing, New DateTime(2015, 4, 9, 13, 30, 0), False, {5.831, 5.831, _
          597, 355, 313, 396}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -242.2435}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.4059}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Maintenance; Drachten/Gorredijk
        IntegrationTestResults = IntegrationTestResults And TestCase("I", "J", _
          New DateTime(2015, 4, 9, 13, 30, 0), False, {5.831, 5.831, 597, 355, _
          313, 396}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0}, {-90.9571, -118.3607, -131.5268, -146.0742, -148.9918, _
          -149.6928, -141.6991, -161.5617, -182.9823, -205.7667, -217.1822, _
          -221.3238, -229.6647, -229.547, -231.3258, -226.0072, -235.5616, _
          -233.342, -235.259, -231.7092, -231.8771, -232.9534, -238.0594, _
          -242.2435}, {-0.1837, -0.2295, -0.2508, -0.2752, -0.2803, -0.2823, _
          -0.2675, -0.2937, -0.3203, -0.3499, -0.3615, -0.3626, -0.372, _
          -0.3727, -0.3717, -0.365, -0.3755, -0.3724, -0.3759, -0.3746, _
          -0.3777, -0.3825, -0.3949, -0.4059}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Week after summer break; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2015, 8, 21, 7, 30, 0), False, {1.399, 1.399, _
          1095, 627, 293, 961}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, -467.6192}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.4271}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Week after summer break (corr. rel. err.); Leeuwarden/Franeker
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "C", _
          New DateTime(2015, 8, 21, 7, 30, 0), False, {0.866, 1.399, 1095, _
          627, 293, 961}, {24, 0.671, 0.8827, 33.8599, 41.4357, -77.4069, _
          51.4748, 0.9221, 0.9246, 0.8502}, {24, 0.7144, 1.0363, 35.5419, _
          46.0207, -0.158, 0.0668, 0.9234, 0.8656, 0.8526}, {80.0693, 71.3444, _
          30.1535, 18.7135, 21.5646, 74.6871, 12.5841, -21.3153, -16.0557, _
          -5.0214, -59.8578, -65.035, -99.315, -106.9292, -100.906, -163.1097, _
          -187.9569, -238.0945, -281.4912, -335.8445, -381.0384, -429.2208, _
          -472.9389, -467.6192}, {0.2812, 0.2457, 0.0929, 0.0508, 0.0596, _
          0.2202, 0.0333, -0.0489, -0.0361, -0.0112, -0.1232, -0.1252, _
          -0.1717, -0.1767, -0.163, -0.242, -0.2574, -0.3005, -0.3331, _
          -0.3753, -0.3971, -0.4242, -0.4395, -0.4271}, {14.8924, -4.6523, _
          -9.9102, 6.5292, -12.3888, -7.7601, -29.7564, -51.0486, -87.0852, _
          -122.4176, -138.2874, -158.3754, -177.1183, -195.7022, -232.4274, _
          -257.1876, -272.0287, -283.1512, -309.0851, -322.058, -340.1488, _
          -343.6312, -317.2904, -302.8298}, {0.051, -0.0153, -0.0312, 0.0199, _
          -0.0374, -0.022, -0.078, -0.128, -0.1986, -0.2617, -0.2824, -0.3042, _
          -0.319, -0.3419, -0.3832, -0.4004, -0.4065, -0.4067, -0.4203, _
          -0.4183, -0.4233, -0.4195, -0.3876, -0.3686})

        ' Christmas morning; Leeuwarden
        IntegrationTestResults = IntegrationTestResults And TestCase("A", _
          Nothing, New DateTime(2015, 12, 25, 9, 0, 0), False, {-1.041, _
          -1.041, 1000, 1189, 1008, 1371}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 189.4183}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.1894}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Christmas morning (correlation relative error); Leeuwarden/Vitens
        IntegrationTestResults = IntegrationTestResults And TestCase("A", "B", _
          New DateTime(2015, 12, 25, 9, 0, 0), True, {0.991, -1.041, 1000, _
          1189, 1008, 1371}, {24, 47.0847, 36.1824, 88.7833, 88.4169, _
          -6689.9305, 2945.1424, 0.9511, 0.9863, 0.9046}, {24, 0.8099, 0.7176, _
          39.0053, 35.6623, -0.1303, 0.1104, 0.9761, 0.9908, 0.9527}, _
          {581.5941, 630.2324, 664.586, 685.5774, 711.5231, 713.7078, _
          759.5657, 772.3307, 767.7746, 721.6509, 726.5312, 691.754, 669.9413, _
          614.1093, 519.1545, 496.1346, 431.2536, 362.6385, 355.4636, _
          334.8511, 326.4106, 273.3801, 248.3248, 189.4183}, {1.5956, 1.7102, _
          1.7719, 1.6668, 1.7297, 1.6322, 1.8436, 1.7487, 1.7204, 1.4266, _
          1.4502, 1.2994, 1.2032, 1.037, 0.7719, 0.7214, 0.5842, 0.4621, _
          0.4415, 0.4063, 0.3844, 0.3016, 0.2681, 0.1894}, {26217.6393, _
          27888.4701, 29092.1221, 29745.8736, 29971.4171, 29777.886, _
          29222.2367, 28438.9962, 27206.1369, 25629.2194, 23881.3918, _
          22187.926, 20408.9651, 18543.0766, 16716.8282, 14863.7979, _
          12990.7359, 11254.5553, 9673.9743, 8311.3091, 7051.7305, 5800.6998, _
          4701.2378, 3638.9989}, {1.2996, 1.3708, 1.4094, 1.413, 1.3926, _
          1.3454, 1.2713, 1.1939, 1.0928, 0.9823, 0.874, 0.7762, 0.6789, _
          0.5879, 0.5081, 0.4327, 0.3618, 0.2997, 0.2472, 0.205, 0.1681, _
          0.1337, 0.1051, 0.0792})

        ' Sequential cache test 1/4; Heerenveen
        IntegrationTestResults = IntegrationTestResults And TestCase("D", _
          Nothing, New DateTime(2015, 8, 21, 2, 0, 0), False, {2.346, 2.346, _
          816, 222, -31, 476}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -593.8604}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.7276}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Sequential cache test 2/4; Heerenveen
        IntegrationTestResults = IntegrationTestResults And TestCase("D", _
          Nothing, New DateTime(2015, 8, 21, 2, 5, 0), False, {2.369, 2.369, _
          814, 216, -36, 469}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -598.0254}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.7344}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Sequential cache test 3/4; Heerenveen
        IntegrationTestResults = IntegrationTestResults And TestCase("D", _
          Nothing, New DateTime(2015, 8, 21, 2, 10, 0), False, {1.928, 1.928, _
          801, 235, -58, 529}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -565.5219}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.7061}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Sequential cache test 4/4; Heerenveen
        IntegrationTestResults = IntegrationTestResults And TestCase("D", _
          Nothing, New DateTime(2015, 8, 21, 2, 15, 0), False, {1.676, 1.676, _
          791, 279, -27, 584}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -512.5031}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6478}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Sequential cache test 1/4 (anticorr. abs. error); Heerenveen/Joure
        IntegrationTestResults = IntegrationTestResults And TestCase("D", "E", _
          New DateTime(2015, 8, 21, 3, 0, 0), False, {-0.998, 1.993, 807, 225, _
          -68, 517}, {24, -0.2941, -1.0118, -16.3902, -45.3371, 428.3739, _
          25.428, -0.4055, -0.9982, 0.1645}, {24, 0.0081, 1.8053, 0.4633, _
          61.0167, -1.297, 0.0499, 0.0053, 0.9984, 0}, {-533.6506, -578.4972, _
          -642.7448, -641.4746, -605.6408, -619.1529, -630.6603, -649.3672, _
          -628.7568, -614.4226, -573.5887, -593.8604, -598.0254, -565.5219, _
          -512.5031, -529.8973, -545.5948, -597.9776, -621.8941, -615.2276, _
          -616.6314, -605.2332, -567.4096, -582.5978}, {-0.6604, -0.684, _
          -0.7261, -0.7552, -0.7388, -0.743, -0.7384, -0.7637, -0.7579, _
          -0.7535, -0.7051, -0.7276, -0.7344, -0.7061, -0.6478, -0.6673, _
          -0.6849, -0.7235, -0.7325, -0.7339, -0.7454, -0.7383, -0.6984, _
          -0.7217}, {542.739, 552.9524, 576.1361, 599.3671, 588.1664, _
          586.1911, 585.1173, 626.8194, 649.2322, 604.3931, 581.9307, _
          609.6077, 614.8427, 603.7487, 591.0562, 590.7233, 607.6901, _
          635.3128, 657.582, 639.803, 618.4015, 610.9732, 601.3713, 604.1553}, _
          {-1.3251, -1.3201, -1.361, -1.413, -1.3772, -1.3447, -1.3228, _
          -1.3116, -1.2771, -1.2768, -1.315, -1.3486, -1.3366, -1.3211, _
          -1.306, -1.2969, -1.3052, -1.2677, -1.2391, -1.2403, -1.2255, _
          -1.232, -1.2398, -1.2655})

        ' Sequential cache test 2/4 (anticorr. abs. error); Heerenveen/Joure
        IntegrationTestResults = IntegrationTestResults And TestCase("D", "E", _
          New DateTime(2015, 8, 21, 3, 5, 0), False, {-0.998, 1.926, 799, 221, _
          -79, 521}, {24, -0.2015, -1.0125, -11.3928, -45.3571, 485.3853, _
          23.4408, -0.2971, -0.9982, 0.0883}, {24, 0.0793, 1.7968, 4.5337, _
          60.9017, -1.2436, 0.0498, 0.0477, 0.9986, 0.0023}, {-578.4972, _
          -642.7448, -641.4746, -605.6408, -619.1529, -630.6603, -649.3672, _
          -628.7568, -614.4226, -573.5887, -593.8604, -598.0254, -565.5219, _
          -512.5031, -529.8973, -545.5948, -597.9776, -621.8941, -615.2276, _
          -616.6314, -605.2332, -567.4096, -582.5978, -578.3391}, {-0.684, _
          -0.7261, -0.7552, -0.7388, -0.743, -0.7384, -0.7637, -0.7579, _
          -0.7535, -0.7051, -0.7276, -0.7344, -0.7061, -0.6478, -0.6673, _
          -0.6849, -0.7235, -0.7325, -0.7339, -0.7454, -0.7383, -0.6984, _
          -0.7217, -0.7236}, {552.9524, 576.1361, 599.3671, 588.1664, _
          586.1911, 585.1173, 626.8194, 649.2322, 604.3931, 581.9307, _
          609.6077, 614.8427, 603.7487, 591.0562, 590.7233, 607.6901, _
          635.3128, 657.582, 639.803, 618.4015, 610.9732, 601.3713, _
          604.1553, 598.2051}, {-1.3201, -1.361, -1.413, -1.3772, -1.3447, _
          -1.3228, -1.3116, -1.2771, -1.2768, -1.315, -1.3486, -1.3366, _
          -1.3211, -1.306, -1.2969, -1.3052, -1.2677, -1.2391, -1.2403, _
          -1.2255, -1.232, -1.2398, -1.2655, -1.2796})

        ' Sequential cache test 3/4 (anticorr. abs. error); Heerenveen/Joure
        IntegrationTestResults = IntegrationTestResults And TestCase("D", "E", _
          New DateTime(2015, 8, 21, 3, 10, 0), False, {-0.998, 2.17, 799, 222, _
          -45, 488}, {24, -0.1723, -1.0158, -9.7788, -45.4486, 504.7356, _
          20.8608, -0.2867, -0.9983, 0.0822}, {24, 0.1302, 1.7912, 7.4189, _
          60.8263, -1.2053, 0.0496, 0.0753, 0.9987, 0.0057}, {-642.7448, _
          -641.4746, -605.6408, -619.1529, -630.6603, -649.3672, -628.7568, _
          -614.4226, -573.5887, -593.8604, -598.0254, -565.5219, -512.5031, _
          -529.8973, -545.5948, -597.9776, -621.8941, -615.2276, -616.6314, _
          -605.2332, -567.4096, -582.5978, -578.3391, -577.4095}, {-0.7261, _
          -0.7552, -0.7388, -0.743, -0.7384, -0.7637, -0.7579, -0.7535, _
          -0.7051, -0.7276, -0.7344, -0.7061, -0.6478, -0.6673, -0.6849, _
          -0.7235, -0.7325, -0.7339, -0.7454, -0.7383, -0.6984, -0.7217, _
          -0.7236, -0.7227}, {576.1361, 599.3671, 588.1664, 586.1911, _
          585.1173, 626.8194, 649.2322, 604.3931, 581.9307, 609.6077, _
          614.8427, 603.7487, 591.0562, 590.7233, 607.6901, 635.3128, 657.582, _
          639.803, 618.4015, 610.9732, 601.3713, 604.1553, 598.2051, _
          599.8287}, {-1.361, -1.413, -1.3772, -1.3447, -1.3228, -1.3116, _
          -1.2771, -1.2768, -1.315, -1.3486, -1.3366, -1.3211, -1.306, _
          -1.2969, -1.3052, -1.2677, -1.2391, -1.2403, -1.2255, -1.232, _
          -1.2398, -1.2655, -1.2796, -1.2878})

        ' Sequential cache test 4/4 (anticorr. abs. error); Heerenveen/Joure
        IntegrationTestResults = IntegrationTestResults And TestCase("D", "E", _
          New DateTime(2015, 8, 21, 3, 15, 0), False, {-0.999, 2.541, 804, _
          226, -2, 453}, {24, -0.2469, -1.0224, -13.8704, -45.636, 461.9018, _
          18.885, -0.4158, -0.9986, 0.1728}, {24, 0.1302, 1.7871, 7.4176, _
          60.77, -1.2018, 0.048, 0.0779, 0.9988, 0.0061}, {-641.4746, _
          -605.6408, -619.1529, -630.6603, -649.3672, -628.7568, -614.4226, _
          -573.5887, -593.8604, -598.0254, -565.5219, -512.5031, -529.8973, _
          -545.5948, -597.9776, -621.8941, -615.2276, -616.6314, -605.2332, _
          -567.4096, -582.5978, -578.3391, -577.4095, -577.9312}, {-0.7552, _
          -0.7388, -0.743, -0.7384, -0.7637, -0.7579, -0.7535, -0.7051, _
          -0.7276, -0.7344, -0.7061, -0.6478, -0.6673, -0.6849, -0.7235, _
          -0.7325, -0.7339, -0.7454, -0.7383, -0.6984, -0.7217, -0.7236, _
          -0.7227, -0.7192}, {599.3671, 588.1664, 586.1911, 585.1173, _
          626.8194, 649.2322, 604.3931, 581.9307, 609.6077, 614.8427, _
          603.7487, 591.0562, 590.7233, 607.6901, 635.3128, 657.582, 639.803, _
          618.4015, 610.9732, 601.3713, 604.1553, 598.2051, 599.8287, _
          599.6043}, {-1.413, -1.3772, -1.3447, -1.3228, -1.3116, -1.2771, _
          -1.2768, -1.315, -1.3486, -1.3366, -1.3211, -1.306, -1.2969, _
          -1.3052, -1.2677, -1.2391, -1.2403, -1.2255, -1.232, -1.2398, _
          -1.2655, -1.2796, -1.2878, -1.2768})

        ' MeanAD/MedianAD test 1/2 (mean); Vlieland
        IntegrationTestResults = IntegrationTestResults And TestCase("K", _
          Nothing, New DateTime(2014, 12, 31, 9, 15, 0), False, {1.049, 1.049, _
          60, 19, -20, 58}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, -40.971}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, -0.6815}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' MeanAD/MedianAD test 2/2 (median); Vlieland
        IntegrationTestResults = IntegrationTestResults And TestCase("K", _
          Nothing, New DateTime(2014, 12, 31, 9, 20, 0), False, {1.082, 1.082, _
          57, 14, -26, 54}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, -43.226}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, -0.7572}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Daylight saving time test 1/5; Wageningseberg
        IntegrationTestResults = IntegrationTestResults And TestCase("L", _
          Nothing, New DateTime(2016, 4, 3, 1, 30, 0), False, {1.112, 1.112, _
          164, 134, 106, 161}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -30.9101}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.188}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Daylight saving time test 2/5; Wageningseberg
        IntegrationTestResults = IntegrationTestResults And TestCase("L", _
          Nothing, New DateTime(2016, 4, 3, 2, 0, 0), False, {1.705, 1.705, _
          164, 136, 119, 153}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -28.5206}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.1735}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Daylight saving time test 3/5; Wageningseberg
        IntegrationTestResults = IntegrationTestResults And TestCase("L", _
          Nothing, New DateTime(2016, 4, 3, 2, 30, 0), False, {3.201, 3.201, _
          164, 133, 123, 143}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -31.3464}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.1907}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Daylight saving time test 4/5; Wageningseberg
        IntegrationTestResults = IntegrationTestResults And TestCase("L", _
          Nothing, New DateTime(2016, 4, 3, 3, 0, 0), False, {2.501, 2.501, _
          164, 132, 119, 145}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -32.3502}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.1968}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Daylight saving time test 5/5; Wageningseberg
        IntegrationTestResults = IntegrationTestResults And TestCase("L", _
          Nothing, New DateTime(2016, 4, 3, 3, 30, 0), False, {3.332, 3.332, _
          164, 132, 122, 142}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, -32.7007}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -0.1989}, {0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Coming/going test 1/3; Eibergen
        IntegrationTestResults = IntegrationTestResults And TestCase("M", _
          Nothing, New DateTime(2016, 4, 28, 17, 50, 0), False, {0, 0, 537, _
          595, 520, 671}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 58.2172}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0.1084}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Coming/going test 2/3; Eibergen
        IntegrationTestResults = IntegrationTestResults And TestCase("M", _
          Nothing, New DateTime(2016, 4, 28, 17, 55, 0), False, {-1.024, _
          -1.024, 535, 605, 537, 673}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 69.5781}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0.1299}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

        ' Coming/going test 3/3; Eibergen
        IntegrationTestResults = IntegrationTestResults And TestCase("M", _
          Nothing, New DateTime(2016, 4, 28, 18, 0, 0), False, {0, 0, 548, _
          609, 538, 679}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 60.9345}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0.1113}, {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}, {0, 0, 0, 0, 0, 0, 0, 0, _
          0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0})

      Next i

      Return IntegrationTestResults

    End Function


    Public Shared Function TestResults As String

      ' Returns a string containing test results (PASSED or FAILED) and
      ' execution time for unit and integration tests.

      Dim Ticks As Int64

      Ticks = Now.Ticks
      TestResults = " - Unit tests " & _
        If(UnitTestResults, "PASSED", "FAILED") & " in " & _
        Round((Now.Ticks-Ticks)/10000).ToString & "ms." & NewLine
      Ticks = Now.Ticks
      TestResults &= " - Integration tests " & _
        If(IntegrationTestResults, "PASSED", "FAILED") & " in " & _
        Round((Now.Ticks-Ticks)/10000).ToString & "ms." & NewLine

      Return TestResults

    End Function


  End Class


End Namespace
