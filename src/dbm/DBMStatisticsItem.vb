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
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMStrings


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStatisticsItem


    Public Count As Integer
    Public Mean, NMBE, RMSD, CVRMSD, Slope, OriginSlope, Angle, OriginAngle,
      Intercept, StandardError, Correlation, ModifiedCorrelation,
      Determination As Double


    Private Function Calibrated As Boolean

      ' ASHRAE Guideline 14-2014, Measurement of Energy, Demand, and Water
      ' Savings
      ' American Society of Heating, Refrigerating and Air Conditioning
      ' Engineers (ASHRAE). Handbook Fundamentals; American Society of Heating,
      ' Refrigerating and Air Conditioning Engineers: Atlanta, GA, USA, 2013;
      ' Volume 111.
      '   "The computer model shall have an NMBE of 5% and a CV(RMSE) of 15%
      '   relative to monthly calibration data. If hourly calibration data are
      '   used, these requirements shall be 10% and 30%, respectively."

      ' IPMVP, International Performance Measurement and Verification Protocol
      ' Efficiency Valuation Organization. International Performance Measurement
      ' and Verification Protocol: Concepts and Options for Determining Energy
      ' and Water Savings, Volume I; Technical Report; Efficiency Valuation
      ' Organization: Washington, DC, USA, 2012.
      '   "Though there is no universal standard for a minimum acceptable R^2
      '   value, 0.75 is often considered a reasonable indicator of a good
      '   causal relationship amongst the energy and independent variables."

      Return Abs(NMBE)<=0.05 And Abs(CVRMSD)<=0.15 And Determination>=0.75

    End Function


    Public Function Brief As String

      Return String.Format(sStatisticsBrief,
        Calibrated, Count, Mean, NMBE*100, RMSD, CVRMSD*100, Determination)

    End Function


  End Class


End Namespace
