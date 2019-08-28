Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
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

    ' Controls the size of the caching algorithms. Use 0 for unlimited items.
    Public Shared CacheSizeFactor As Integer = 6 ' Keep at default (0-16)

    ' Time interval at which the calculation is run.
    Public Shared CalculationInterval As Integer = 300 ' seconds, 5 minutes

    ' Use forecast of the previous Sunday for holidays.
    Public Shared UseSundayForHolidays As Boolean = True

    ' Number of weeks to look back to forecast the current value and
    ' control limits.
    Public Shared ComparePatterns As Integer = 12 ' weeks

    ' Number of previous intervals used to smooth the data.
    Public Shared EMAPreviousPeriods As Integer =
      CInt(0.5*3600/CalculationInterval-1) ' 5 intervals, 30 minutes

    ' Confidence interval used for removing outliers.
    Public Shared OutlierCI As Double = 0.99

    ' Confidence interval used for determining control limits.
    Public Shared BandwidthCI As Double = 0.99

    ' Number of previous intervals used to calculate forecast error correlation
    ' when an event is found.
    Public Shared CorrelationPreviousPeriods As Integer =
      CInt(2*3600/CalculationInterval-1) ' 23 intervals, 2 hours

    ' Absolute correlation lower limit for detecting (anti)correlation.
    Public Shared CorrelationThreshold As Double = Sqrt(0.7) ' 0.83666

    ' Regression angle range (around -45/+45 degrees) required when suppressing
    ' based on (anti)correlation.
    Public Shared RegressionAngleRange As Double =
      SlopeToAngle(2)-SlopeToAngle(1) ' 18.435 degrees, factor 2


  End Class


End Namespace
