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
    Public TotalWeight, Mean, NMBE, RMSD, CVRMSD, Slope, OriginSlope, Angle,
      OriginAngle, Intercept, StandardError, Correlation, ModifiedCorrelation,
      Determination As Double


    Public Function Calibrated As Boolean

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
      '   "Though there is no universal standard for a minimum acceptable R²
      '   value, 0.75 is often considered a reasonable indicator of a good
      '   causal relationship amongst the energy and independent variables."

      Return Abs(NMBE) <= 0.1 And
        Abs(CVRMSD) <= 0.3 And
        Determination >= 0.75

    End Function


    Public Function SystematicError As Double

      ' The normalized mean bias error is used as a measure of the systematic
      ' error.

      Return NMBE

    End Function


    Public Function RandomError As Double

      ' For the random error, the difference between the absolute normalized
      ' mean bias error and the absolute coefficient of variation of the
      ' root-mean-square deviation is used.

      Return Abs(CVRMSD)-Abs(SystematicError)

    End Function


    Public Function Fit As Double

      ' The determination, R², as a measure of fit.

      Return Determination

    End Function


    Public Function Brief As String

      ' Model calibration metrics: systematic error, random error, and fit
      ' DBM can calculate model calibration metrics. This information is exposed
      ' in the DBMTester utility and the DBMDataRef data reference. The model is
      ' considered to be calibrated if all of the following conditions are met:
      '   * the absolute normalized mean bias error (NMBE, as a measure of bias)
      '     is 10% or lower,
      '   * the absolute coefficient of variation of the root-mean-square
      '     deviation (CV(RMSD), as a measure of error) is 30% or lower,
      '   * the determination (R², as a measure of fit) is 0.75 or higher.
      ' The normalized mean bias error is used as a measure of the systematic
      ' error. For the random error, the difference between the absolute
      ' normalized mean bias error and the absolute coefficient of variation of
      ' the root-mean-square deviation is used.
      ' There are several agencies that have developed guidelines and
      ' methodologies to establish a measure of the accuracy of models. We
      ' decided to follow the guidelines as documented in ASHRAE Guideline
      ' 14-2014, Measurement of Energy, Demand, and Water Savings, by the
      ' American Society of Heating, Refrigerating and Air Conditioning
      ' Engineers, and International Performance Measurement and Verification
      ' Protocol: Concepts and Options for Determining Energy and Water Savings,
      ' Volume I, by the Efficiency Valuation Organization.

      If Count < 3 Then Return sStatisticsInsufficientData ' Need at least 3 pts

      Return String.Format(sStatisticsBrief,
        Calibrated, Count, SystematicError, RandomError, Fit)

    End Function


  End Class


End Namespace
