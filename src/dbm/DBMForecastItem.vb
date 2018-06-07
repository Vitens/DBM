Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2018  J.H. Fitié, Vitens N.V.
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
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMForecastItem


    Public Measurement, ForecastValue, LowerControlLimit,
      UpperControlLimit As Double


    Public Function Range(p As Double) As Double

      ' Returns the range between the control limits and forecast value scaled
      ' for the requested confidence interval. This function requires that the
      ' results have been calculated for the default range using the Forecast
      ' function before. Always use the full sample size and appropriate
      ' distribution as the calculation is run on an EMA-smoothed series of data
      ' and not on a single forecast result from which outliers might be
      ' removed.

      Return (UpperControlLimit-ForecastValue)/
        ControlLimitRejectionCriterion(2*BandwidthCI-1, ComparePatterns-1)*
        ControlLimitRejectionCriterion(2*p-1, ComparePatterns-1)

    End Function


  End Class


End Namespace
