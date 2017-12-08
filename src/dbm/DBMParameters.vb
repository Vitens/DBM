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


Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMParameters


    ' This class contains default values for several parameters which DBM uses
    ' for its calculations. The values for these parameters can be changed
    ' at runtime.
    ' The default parameter values were determined by analysing the results
    ' of a particle swarm optimization method using real production data from
    ' Vitens, and then manually tweaking these values for optimal results and
    ' performance.


    ' Time interval at which the calculation is run (seconds).
    ' Default: 5 minutes.
    Public Shared CalculationInterval As Integer = 300

    ' Number of weeks to look back to predict the current value
    ' and control limits.
    ' Default: 17 weeks (PSO Avg:20.8 SD:2.278 Delta:-18.3%), 4 months
    Public Shared ComparePatterns As Integer = CInt(52/12*4)

    ' Number of previous intervals used to smooth the data.
    ' Default: 5 intervals (PSO Avg:22.5 SD:1.658 Delta:-77.8%), 30 minutes
    '  (reduced to minimize lag)
    Public Shared EMAPreviousPeriods As Integer = _
      CInt(0.5*3600/CalculationInterval-1)

    ' Confidence interval used for removing outliers and determining
    ' control limits (ratio).
    ' Default: 0.95 (PSO Avg:0.92834 SD:0.003 Delta:-2.3%)
    Public Shared ConfidenceInterval As Double = 0.95

    ' Number of previous intervals used to calculate prediction error
    ' correlation when an exception is found.
    ' Default: 47 intervals (PSO Avg:55.3 SD:1.785 Delta:-15.0%), 4 hours
    Public Shared CorrelationPreviousPeriods As Integer = _
      CInt(4*3600/CalculationInterval-1)

    ' Absolute correlation lower limit for detecting (anti)correlation.
    ' Default: 0.77460 (PSO Avg:0.77990 SD:0.017 Delta:-0.7%) for a
    '  determination of 0.6
    Public Shared CorrelationThreshold As Double = Sqrt(0.6)

    ' Regression angle range (around -45/+45 degrees) required when suppressing
    ' based on (anti)correlation (degrees).
    ' Default: 21.03751 degrees (PSO Avg:20.77809 SD:1.368 Delta:1.2%), allow
    '  factor 2.25 difference between values
    Public Shared RegressionAngleRange As Double = _
      SlopeToAngle(2.25)-SlopeToAngle(1)


  End Class


End Namespace
