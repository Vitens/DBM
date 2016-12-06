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

Public Class DBMPoint

    Public DBMDataManager As DBMDataManager

    Public Sub New(DBMPointDriver As DBMPointDriver)
        DBMDataManager=New DBMDataManager(DBMPointDriver)
    End Sub

    Public Function Calculate(Timestamp As DateTime,IsInputDBMPoint As Boolean,HasCorrelationDBMPoint As Boolean,Optional SubstractDBMPoint As DBMPoint=Nothing) As DBMResult
        Dim CorrelationCounter,EMACounter,PatternCounter As Integer
        Dim Pattern(DBMParameters.ComparePatterns),CurrValueEMA(DBMParameters.EMAPreviousPeriods),PredValueEMA(DBMParameters.EMAPreviousPeriods),LowContrLimitEMA(DBMParameters.EMAPreviousPeriods),UppContrLimitEMA(DBMParameters.EMAPreviousPeriods) As Double
        Dim DBMStatistics As New DBMStatistics
        Calculate=New DBMResult
        For CorrelationCounter=0 To DBMParameters.CorrelationPreviousPeriods
            If CorrelationCounter=0 Or (IsInputDBMPoint And Calculate.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                For EMACounter=DBMParameters.EMAPreviousPeriods To 0 Step -1
                    If CorrelationCounter=0 Or (CorrelationCounter>0 And EMACounter=DBMParameters.EMAPreviousPeriods) Then
                        If CorrelationCounter>0 And EMACounter=DBMParameters.EMAPreviousPeriods Then ' Reuse calculation results when moving back for correlation calculation
                            CurrValueEMA=DBMFunctions.ArrayRotateLeft(CurrValueEMA)
                            PredValueEMA=DBMFunctions.ArrayRotateLeft(PredValueEMA)
                            LowContrLimitEMA=DBMFunctions.ArrayRotateLeft(LowContrLimitEMA)
                            UppContrLimitEMA=DBMFunctions.ArrayRotateLeft(UppContrLimitEMA)
                        End If
                        For PatternCounter=DBMParameters.ComparePatterns To 0 Step -1
                            Pattern(DBMParameters.ComparePatterns-PatternCounter)=DBMDataManager.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*DBMParameters.CalculationInterval,Timestamp)))
                            If SubstractDBMPoint IsNot Nothing Then
                                Pattern(DBMParameters.ComparePatterns-PatternCounter)-=SubstractDBMPoint.DBMDataManager.Value(DateAdd("d",-PatternCounter*7,DateAdd("s",-(EMACounter+CorrelationCounter)*DBMParameters.CalculationInterval,Timestamp)))
                            End If
                        Next PatternCounter
                        DBMStatistics.Calculate(DBMMath.RemoveOutliers(Pattern.Take(Pattern.Length-1).ToArray)) ' Calculate statistics for data after removing outliers
                        CurrValueEMA(EMACounter)=Pattern(DBMParameters.ComparePatterns)
                        PredValueEMA(EMACounter)=DBMParameters.ComparePatterns*DBMStatistics.Slope+DBMStatistics.Intercept ' Extrapolate linear regression
                        LowContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)-DBMMath.ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,DBMStatistics.Count-1)*DBMStatistics.StDevSLinReg
                        UppContrLimitEMA(EMACounter)=PredValueEMA(EMACounter)+DBMMath.ControlLimitRejectionCriterion(DBMParameters.ConfidenceInterval,DBMStatistics.Count-1)*DBMStatistics.StDevSLinReg
                    End If
                Next EMACounter
                Calculate.AbsoluteErrors(DBMParameters.CorrelationPreviousPeriods-CorrelationCounter)=DBMMath.CalculateExpMovingAvg(PredValueEMA)-DBMMath.CalculateExpMovingAvg(CurrValueEMA) ' Absolute error compared to prediction
                Calculate.RelativeErrors(DBMParameters.CorrelationPreviousPeriods-CorrelationCounter)=DBMMath.CalculateExpMovingAvg(PredValueEMA)/DBMMath.CalculateExpMovingAvg(CurrValueEMA)-1 ' Relative error compared to prediction
                If CorrelationCounter=0 Then
                    Calculate.CurrValue=DBMMath.CalculateExpMovingAvg(CurrValueEMA)
                    Calculate.PredValue=DBMMath.CalculateExpMovingAvg(PredValueEMA)
                    Calculate.LowContrLimit=DBMMath.CalculateExpMovingAvg(LowContrLimitEMA)
                    Calculate.UppContrLimit=DBMMath.CalculateExpMovingAvg(UppContrLimitEMA)
                    If Calculate.CurrValue<Calculate.LowContrLimit Then ' Lower control limit exceeded
                        Calculate.Factor=(Calculate.PredValue-Calculate.CurrValue)/(Calculate.LowContrLimit-Calculate.PredValue)
                    ElseIf Calculate.CurrValue>Calculate.UppContrLimit Then ' Upper control limit exceeded
                        Calculate.Factor=(Calculate.CurrValue-Calculate.PredValue)/(Calculate.UppContrLimit-Calculate.PredValue)
                    End If
                    Calculate.OriginalFactor=Calculate.Factor ' Store original factor before possible suppression
                End If
            End If
        Next CorrelationCounter
        Return Calculate
    End Function

End Class
