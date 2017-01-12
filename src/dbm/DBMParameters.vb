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

Namespace Vitens.DynamicBandwidthMonitor

    Public Class DBMParameters

        Public Shared CalculationInterval As Integer =        300 ' seconds; 300=5 minutes. Time interval at which the calculation is run.
        Public Shared ComparePatterns As Integer =            12 ' weeks. Number of weeks to look back to predict the current value and control limits.
        Public Shared EMAPreviousPeriods As Integer =         CInt(0.5*3600/CalculationInterval) ' intervals; 6=35 minutes, current value inclusive. Number of previous intervals used to smooth the data.
        Public Shared ConfidenceInterval As Double =          0.99 ' percent. Confidence interval used for removing outliers and determining control limits.
        Public Shared CorrelationPreviousPeriods As Integer = CInt(2*3600/CalculationInterval-1) ' intervals; 23=2 hours, current value inclusive. Number of previous intervals used to calculate prediction error correlation when an exception is found.
        Public Shared CorrelationThreshold As Double =        0.83666 ' Absolute correlation lower limit for detecting (anti)correlation.
        Public Shared MaxPointPredictions As Integer =        CInt(24*3600/CalculationInterval+EMAPreviousPeriods+CorrelationPreviousPeriods) ' Maximum number of cached prediction results per point; large enough for at least one day.
        Public Shared MaxDataManagerValues As Integer =       MaxPointPredictions*(ComparePatterns+1) ' Maximum number of cached values per point; large enough for at least one day.

    End Class

End Namespace
