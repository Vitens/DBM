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

Imports System.Double
Imports System.Math
Imports Vitens.DynamicBandwidthMonitor.DBMMath

Namespace Vitens.DynamicBandwidthMonitor

  Public Class DBMStatistics

    Public Count As Integer
    Public Slope, Angle, Intercept, StandardError, Correlation, _
      Determination, ModifiedSlope, ModifiedAngle, _
      ModifiedCorrelation As Double

    Public Function ShallowCopy As DBMStatistics
      Return DirectCast(Me.MemberwiseClone, DBMStatistics)
    End Function

    Public Sub Calculate(ValuesY() As Double, _
      Optional ValuesX() As Double = Nothing)
      Dim i As Integer
      Dim SumX, SumY, SumXX, SumYY, SumXY As Double
      If ValuesX Is Nothing Then
        ReDim ValuesX(ValuesY.Length-1)
        For i = 0 To ValuesX.Length-1
          ValuesX(i) = i
        Next i
      End If
      Count = 0
      For i = 0 To ValuesY.Length-1
        If Not IsNaN(ValuesX(i)) And Not IsNaN(ValuesY(i)) Then
          SumX += ValuesX(i)
          SumY += ValuesY(i)
          SumXX += ValuesX(i)^2
          SumYY += ValuesY(i)^2
          SumXY += ValuesX(i)*ValuesY(i)
          Count += 1
        End If
      Next i
      Slope = (Count*SumXY-SumX*SumY)/(Count*SumXX-SumX^2)
      Angle = SlopeToAngle(Slope) ' Angle in degrees
      Intercept = (SumX*SumXY-SumY*SumXX)/(SumX^2-Count*SumXX)
      ' Standard error of the predicted y-value for each x in the regression.
      ' The standard error is a measure of the amount of error in the
      ' prediction of y for an individual x
      StandardError = 0
      For i = 0 to ValuesY.Length-1
        If Not IsNaN(ValuesX(i)) And Not IsNaN(ValuesY(i)) Then
          StandardError += (ValuesY(i)-ValuesX(i)*Slope-Intercept)^2
        End If
      Next i
      ' n-2 is used because two parameters (slope and intercept) were
      ' estimated in order to estimate the sum of squares
      StandardError = Sqrt(StandardError/(Count-2))
      ' Wikipedia: A number that quantifies some type of correlation and
      ' dependence, meaning statistical relationships between two or more
      ' random variables or observed data values
      Correlation = (Count*SumXY-SumX*SumY)/ _
        Sqrt((Count*SumXX-SumX^2)*(Count*SumYY-SumY^2))
      ' Wikipedia: A number that indicates the proportion of the variance in
      ' the dependent variable that is predictable from the
      ' independent variable
      Determination = Correlation^2
      ' Average is not removed, as expected average is zero
      ModifiedSlope = SumXY/SumXX
      ModifiedAngle = SlopeToAngle(ModifiedSlope) ' Angle in degrees
      ModifiedCorrelation = SumXY/Sqrt(SumXX)/Sqrt(SumYY)
    End Sub

  End Class

End Namespace
