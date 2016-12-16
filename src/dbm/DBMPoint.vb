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

    Public Class DBMPoint

        Public DBMDataManager As DBMDataManager
        Private DBMPredictions As New Collections.Generic.Dictionary(Of DateTime,DBMPrediction)
        Private PrevSubstractDBMPoint As DBMPoint

        Public Sub New(DBMPointDriver As DBMPointDriver)
            DBMDataManager=New DBMDataManager(DBMPointDriver)
        End Sub

        Public Function Calculate(Timestamp As DateTime,IsInputDBMPoint As Boolean,HasCorrelationDBMPoint As Boolean,Optional SubstractDBMPoint As DBMPoint=Nothing) As DBMResult
            Dim CorrelationCounter,EMACounter,PatternCounter As Integer
            Dim CalcTimestamp As DateTime
            Dim Pattern(DBMParameters.ComparePatterns),MeasValueEMA(DBMParameters.EMAPreviousPeriods),PredValueEMA(DBMParameters.EMAPreviousPeriods),LowContrLimitEMA(DBMParameters.EMAPreviousPeriods),UppContrLimitEMA(DBMParameters.EMAPreviousPeriods) As Double
            Dim DBMPrediction As New DBMPrediction
            Calculate=New DBMResult
            If SubstractDBMPoint IsNot PrevSubstractDBMPoint Then ' Can we reuse stored results?
                DBMPredictions.Clear
                PrevSubstractDBMPoint=SubstractDBMPoint
            End If
            For CorrelationCounter=0 To DBMParameters.CorrelationPreviousPeriods
                If CorrelationCounter=0 Or (IsInputDBMPoint And Calculate.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                    For EMACounter=DBMParameters.EMAPreviousPeriods To 0 Step -1
                        CalcTimestamp=DateAdd("s",-(EMACounter+CorrelationCounter)*DBMParameters.CalculationInterval,Timestamp)
                        If DBMPredictions.ContainsKey(CalcTimestamp) Then ' From cache
                            DBMPrediction=DBMPredictions.Item(CalcTimestamp).ShallowCopy
                        Else ' Calculate data
                            For PatternCounter=DBMParameters.ComparePatterns To 0 Step -1
                                Pattern(DBMParameters.ComparePatterns-PatternCounter)=DBMDataManager.Value(DateAdd("d",-PatternCounter*7,CalcTimestamp))
                                If SubstractDBMPoint IsNot Nothing Then
                                    Pattern(DBMParameters.ComparePatterns-PatternCounter)-=SubstractDBMPoint.DBMDataManager.Value(DateAdd("d",-PatternCounter*7,CalcTimestamp))
                                End If
                            Next PatternCounter
                            DBMPrediction.Calculate(Pattern)
                            Do While DBMPredictions.Count>=DBMParameters.MaxDBMPointCacheSize ' Limit cache size
                                DBMPredictions.Remove(DBMPredictions.ElementAt(CInt(Math.Floor(Rnd*DBMPredictions.Count))).Key) ' Remove random cached value
                            Loop
                            DBMPredictions.Add(CalcTimestamp,DBMPrediction.ShallowCopy) ' Add to cache
                        End If
                        MeasValueEMA(EMACounter)=DBMPrediction.MeasValue
                        PredValueEMA(EMACounter)=DBMPrediction.PredValue
                        LowContrLimitEMA(EMACounter)=DBMPrediction.LowContrLimit
                        UppContrLimitEMA(EMACounter)=DBMPrediction.UppContrLimit
                    Next EMACounter
                    DBMPrediction=New DBMPrediction(DBMMath.CalculateExpMovingAvg(MeasValueEMA),DBMMath.CalculateExpMovingAvg(PredValueEMA),DBMMath.CalculateExpMovingAvg(LowContrLimitEMA),DBMMath.CalculateExpMovingAvg(UppContrLimitEMA))
                    Calculate.AbsoluteErrors(DBMParameters.CorrelationPreviousPeriods-CorrelationCounter)=DBMPrediction.PredValue-DBMPrediction.MeasValue ' Absolute error compared to prediction
                    Calculate.RelativeErrors(DBMParameters.CorrelationPreviousPeriods-CorrelationCounter)=DBMPrediction.PredValue/DBMPrediction.MeasValue-1 ' Relative error compared to prediction
                    If CorrelationCounter=0 Then
                        If DBMPrediction.MeasValue<DBMPrediction.LowContrLimit Then ' Lower control limit exceeded
                            Calculate.Factor=(DBMPrediction.PredValue-DBMPrediction.MeasValue)/(DBMPrediction.LowContrLimit-DBMPrediction.PredValue)
                        ElseIf DBMPrediction.MeasValue>DBMPrediction.UppContrLimit Then ' Upper control limit exceeded
                            Calculate.Factor=(DBMPrediction.MeasValue-DBMPrediction.PredValue)/(DBMPrediction.UppContrLimit-DBMPrediction.PredValue)
                        End If
                        Calculate.OriginalFactor=Calculate.Factor ' Store original factor before possible suppression
                        Calculate.DBMPrediction=DBMPrediction.ShallowCopy
                    End If
                End If
            Next CorrelationCounter
            Return Calculate
        End Function

    End Class

End Namespace
