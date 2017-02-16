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


    ' Time interval at which the calculation is run (seconds).
    ' Default: 5 minutes.
    Public Shared CalculationInterval As Integer = 300

    ' Number of weeks to look back to predict the current value
    ' and control limits.
    Public Shared ComparePatterns As Integer = 12

    ' Number of previous intervals used to smooth the data.
    ' Default: 35 minutes, current value inclusive.
    Public Shared EMAPreviousPeriods As Integer = _
      CInt(0.5*3600/CalculationInterval)

    ' Confidence interval used for removing outliers and determining
    ' control limits (ratio).
    Public Shared ConfidenceInterval As Double = 0.99

    ' Number of previous intervals used to calculate prediction error
    ' correlation when an exception is found.
    ' Default: 2 hours, current value inclusive.
    Public Shared CorrelationPreviousPeriods As Integer = _
      CInt(2*3600/CalculationInterval-1)

    ' Absolute correlation lower limit for detecting (anti)correlation.
    ' Default: 0.83666 for a determination of 0.7.
    Public Shared CorrelationThreshold As Double = Sqrt(0.7)

    ' Regression angle range (around -45/+45 degrees) required when suppressing
    ' based on (anti)correlation (degrees).
    ' Default: Allow factor 2 difference between values (18.435 degrees).
    Public Shared RegressionAngleRange As Double = SlopeToAngle(2)-45

    ' Maximum number of cached prediction results per point.
    ' Default: large enough for one day.
    Public Shared MaxPointPredictions As Integer = _
      CInt(24*3600/CalculationInterval+ _
      EMAPreviousPeriods+CorrelationPreviousPeriods)

    ' Maximum number of cached values per point.
    ' Default: large enough for one day.
    Public Shared MaxDataManagerValues As Integer = _
      MaxPointPredictions*(ComparePatterns+1)


  End Class


End Namespace
