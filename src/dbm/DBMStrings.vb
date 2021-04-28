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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStrings


    Public Const sTimestamp As String = "Timestamp"
    Public Const sFactor As String = "Factor"
    Public Const sMeasurement As String = "Measurement"
    Public Const sForecast As String = "Forecast"
    Public Const sLowerControlLimit As String = "Lower control limit"
    Public Const sUpperControlLimit As String = "Upper control limit"
    Public Const sNumberFormat As String = "G5"
    Public Const sPercentageFormat As String = "0.0000000%"
    Public Const sCsvComment As String = "# "
    Public Const sStatisticsInsufficientData As String =
      "Insufficient data for calculating model calibration metrics"
    Public Const sStatisticsBrief As String =
      "Calibrated: {0} (" &
      "n {1}; " &
      "Systematic error {2:" & sPercentageFormat & "}; " &
      "Random error {3:" & sPercentageFormat & "}; " &
      "Fit {4:" & sPercentageFormat & "})"
    Public Const sQualityTests As String =
      "n {0}; C {1:" & sPercentageFormat & "}; " &
      "SE {2:" & sPercentageFormat & "} (SD={3:" & sPercentageFormat & "}); " &
      "RE {4:" & sPercentageFormat & "} (SD={5:" & sPercentageFormat & "}); " &
      "F {6:" & sPercentageFormat & "} (SD={7:" & sPercentageFormat & "})"
    Public Const sForecastFactorAnnotation As String =
      sFactor & " {0:" & sNumberFormat & "}"


  End Class


End Namespace
