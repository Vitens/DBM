Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStrings


    Public Const UnsignedAssembly As String = "Unsigned assembly"
    Public Const TimestampText As String = "Timestamp"
    Public Const FactorText As String = "Factor"
    Public Const MeasurementText As String = "Measurement"
    Public Const ForecastText As String = "Forecast"
    Public Const LowerControlLimitText As String = "Lower control limit"
    Public Const UpperControlLimitText As String = "Upper control limit"
    Public Const DecimalFormat As String = "G5"
    Public Const PercentageFormat As String = "0.0%"
    Public Const CSVComment As String = "# "
    Public Const StatisticsInsufficientData As String =
      "Insufficient data for calculating model calibration metrics"
    Public Const StatisticsBrief As String =
      "Calibrated: {0} (" &
      "n {1}; " &
      "Systematic error {2:" & PercentageFormat & "}; " &
      "Random error {3:" & PercentageFormat & "}; " &
      "Fit {4:" & PercentageFormat & "})"
    Public Const QualityTests As String =
      "n {0}; C {1:" & PercentageFormat & "}; " &
      "SE {2:" & PercentageFormat & "} (SD={3:" & PercentageFormat & "}); " &
      "RE {4:" & PercentageFormat & "} (SD={5:" & PercentageFormat & "}); " &
      "F {6:" & PercentageFormat & "} (SD={7:" & PercentageFormat & "}); " &
      "Score {8:" & PercentageFormat & "}"
    Public Const ForecastFactorAnnotation As String =
      Factor & " {0:" & DecimalFormat & "}"


  End Class


End Namespace
