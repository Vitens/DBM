Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

Namespace DBM

    Public Class DBMParameters

        Public Shared CalculationInterval As Integer=           300 ' seconds; 300 = 5 minutes
        Public Shared ComparePatterns As Integer=               12 ' weeks
        Public Shared EMAPreviousPeriods As Integer=            CInt(0.5*3600/CalculationInterval) ' previous periods; 6 = 35 minutes, current value inclusive
        Public Shared ConfidenceInterval As Double=             0.99 ' confidence interval for outlier detection and control limits
        Public Shared CorrelationPreviousPeriods As Integer=    CInt(2*3600/CalculationInterval-1) ' previous periods; 23 = 2 hours, current value inclusive
        Public Shared CorrelationThreshold As Double=           0.83666 ' absolute correlation lower limit for detecting (anti)correlation
        Public Shared MaximumCacheSize As Integer=              CInt((24*3600/CalculationInterval+EMAPreviousPeriods+CorrelationPreviousPeriods)*(ComparePatterns+1)) ' maximum number of cached values per point; large enough for at least one day

    End Class

End Namespace
