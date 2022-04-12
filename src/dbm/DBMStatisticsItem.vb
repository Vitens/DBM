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


Imports System
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMStrings


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMStatisticsItem


    Private _count As Integer
    Private _mean, _nmbe, _rmsd, _cvrmsd, _slope, _originSlope, _angle,
      _originAngle, _intercept, _standardError, _correlation,
      _modifiedCorrelation, _determination As Double


    Public Property Count As Integer
      Get
        Return _count
      End Get
      Set(value As Integer)
        _count = value
      End Set
    End Property


    Public Property Mean As Double
      Get
        Return _mean
      End Get
      Set(value As Double)
        _mean = value
      End Set
    End Property


    Public Property NMBE As Double
      Get
        Return _nmbe
      End Get
      Set(value As Double)
        _nmbe = value
      End Set
    End Property


    Public Property RMSD As Double
      Get
        Return _rmsd
      End Get
      Set(value As Double)
        _rmsd = value
      End Set
    End Property


    Public Property CVRMSD As Double
      Get
        Return _cvrmsd
      End Get
      Set(value As Double)
        _cvrmsd = value
      End Set
    End Property


    Public Property Slope As Double
      Get
        Return _slope
      End Get
      Set(value As Double)
        _slope = value
      End Set
    End Property


    Public Property OriginSlope As Double
      Get
        Return _originSlope
      End Get
      Set(value As Double)
        _originSlope = value
      End Set
    End Property


    Public Property Angle As Double
      Get
        Return _angle
      End Get
      Set(value As Double)
        _angle = value
      End Set
    End Property


    Public Property OriginAngle As Double
      Get
        Return _originAngle
      End Get
      Set(value As Double)
        _originAngle = value
      End Set
    End Property


    Public Property Intercept As Double
      Get
        Return _intercept
      End Get
      Set(value As Double)
        _intercept = value
      End Set
    End Property


    Public Property StandardError As Double
      Get
        Return _standardError
      End Get
      Set(value As Double)
        _standardError = value
      End Set
    End Property


    Public Property Correlation As Double
      Get
        Return _correlation
      End Get
      Set(value As Double)
        _correlation = value
      End Set
    End Property


    Public Property ModifiedCorrelation As Double
      Get
        Return _modifiedCorrelation
      End Get
      Set(value As Double)
        _modifiedCorrelation = value
      End Set
    End Property


    Public Property Determination As Double
      Get
        Return _determination
      End Get
      Set(value As Double)
        _determination = value
      End Set
    End Property


    Public Function HasInsufficientData As Boolean

      Return Me.Count < 3 ' Need at least 3 data points.

    End Function


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

      Return Abs(Me.NMBE) <= 0.1 And
        Abs(Me.CVRMSD) <= 0.3 And
        Me.Determination >= 0.75

    End Function


    Public Function SystematicError As Double

      ' The normalized mean bias error is used as a measure of the systematic
      ' error.

      Return Me.NMBE

    End Function


    Public Function RandomError As Double

      ' For the random error, the difference between the absolute normalized
      ' mean bias error and the absolute coefficient of variation of the
      ' root-mean-square deviation is used.

      Return Abs(Me.CVRMSD)-Abs(SystematicError)

    End Function


    Public Function Fit As Double

      ' The determination, R², as a measure of fit.

      Return Me.Determination

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

      If HasInsufficientData Then Return StatisticsInsufficientData

      Return String.Format(StatisticsBrief,
        Calibrated, Me.Count, SystematicError, RandomError, Fit)

    End Function


  End Class


End Namespace
