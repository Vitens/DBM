Option Explicit
Option Strict


' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
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


    ' Time interval at which the calculation is run.
    Public Shared CalculationInterval As Integer = 300 ' seconds, 5 minutes

    ' Number of weeks to look back to predict the current value and
    ' control limits.
    Public Shared ComparePatterns As Integer = CInt(52/12*4) ' 17 weeks, 4 mos.

    ' Number of previous intervals used to smooth the data.
    Public Shared EMAPreviousPeriods As Integer = _
      CInt(0.5*3600/CalculationInterval-1) ' 5 intervals, 30 minutes

    ' Confidence interval used for removing outliers and determining
    ' control limits.
    Public Shared ConfidenceInterval As Double = 0.99

    ' Number of previous intervals used to calculate prediction error
    ' correlation when an exception is found.
    Public Shared CorrelationPreviousPeriods As Integer = _
      CInt(4*3600/CalculationInterval-1) ' 47 intervals, 4 hours

    ' Absolute correlation lower limit for detecting (anti)correlation.
    Public Shared CorrelationThreshold As Double = Sqrt(0.6) ' 0.77460

    ' Regression angle range (around -45/+45 degrees) required when suppressing
    ' based on (anti)correlation.
    Public Shared RegressionAngleRange As Double = _
      SlopeToAngle(2.25)-SlopeToAngle(1) ' 21.03751 degrees, factor 2.25


  End Class


End Namespace
