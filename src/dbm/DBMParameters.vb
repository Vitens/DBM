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


Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMParameters


    ' This class contains default values for several parameters which DBM uses
    ' for its calculations. The values for these parameters can be changed
    ' at runtime.

    ' Time interval at which the calculation is run.
    Private Shared _calculationInterval As Integer = 300 ' seconds, 5 minutes

    Public Shared Property CalculationInterval As Integer
      Get
        Return _calculationInterval
      End Get
      Set(value As Integer)
        _calculationInterval = value
      End Set
    End Property


    ' Use forecast of the previous Sunday for holidays.
    Private Shared _useSundayForHolidays As Boolean = True

    Public Shared Property UseSundayForHolidays As Boolean
      Get
        Return _useSundayForHolidays
      End Get
      Set(value As Boolean)
        _useSundayForHolidays = value
      End Set
    End Property


    ' Number of weeks to look back to forecast the current value and
    ' control limits.
    Private Shared _comparePatterns As Integer = 12 ' weeks

    Public Shared Property ComparePatterns As Integer
      Get
        Return _comparePatterns
      End Get
      Set(value As Integer)
        _comparePatterns = value
      End Set
    End Property


    ' Number of previous intervals used to smooth the data.
    Private Shared _emaPreviousPeriods As Integer =
      CInt(0.5*3600/CalculationInterval-1) ' 5 intervals, 30 minutes

    Public Shared Property EMAPreviousPeriods As Integer
      Get
        Return _emaPreviousPeriods
      End Get
      Set(value As Integer)
        _emaPreviousPeriods = value
      End Set
    End Property


    ' Confidence interval used for removing outliers.
    Private Shared _outlierCI As Double = 0.99

    Public Shared Property OutlierCI As Double
      Get
        Return _outlierCI
      End Get
      Set(value As Double)
        _outlierCI = value
      End Set
    End Property


    ' Confidence interval used for determining control limits.
    Private Shared _bandwidthCI As Double = 0.99

    Public Shared Property BandwidthCI As Double
      Get
        Return _bandwidthCI
      End Get
      Set(value As Double)
        _bandwidthCI = value
      End Set
    End Property


    ' Number of previous intervals used to calculate forecast error correlation
    ' when an event is found.
    Private Shared _correlationPreviousPeriods As Integer =
      CInt(2*3600/CalculationInterval-1) ' 23 intervals, 2 hours

    Public Shared Property CorrelationPreviousPeriods As Integer
      Get
        Return _correlationPreviousPeriods
      End Get
      Set(value As Integer)
        _correlationPreviousPeriods = value
      End Set
    End Property


    ' Absolute correlation lower limit for detecting (anti)correlation.
    Private Shared _correlationThreshold As Double = Sqrt(0.7) ' 0.83666

    Public Shared Property CorrelationThreshold As Double
      Get
        Return _correlationThreshold
      End Get
      Set(value As Double)
        _correlationThreshold = value
      End Set
    End Property


    ' Regression angle range (around -45/+45 degrees) required when suppressing
    ' based on (anti)correlation.
    Private Shared _regressionAngleRange As Double =
      SlopeToAngle(2)-SlopeToAngle(1) ' 18.435 degrees, factor 2

    Public Shared Property RegressionAngleRange As Double
      Get
        Return _regressionAngleRange
      End Get
      Set(value As Double)
        _regressionAngleRange = value
      End Set
    End Property


  End Class


End Namespace
