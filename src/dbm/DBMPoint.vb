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
            Dim DBMPrediction As New DBMPrediction
            Dim Patterns(DBMParameters.ComparePatterns),MeasValues(DBMParameters.EMAPreviousPeriods),PredValues(DBMParameters.EMAPreviousPeriods),LowContrLimits(DBMParameters.EMAPreviousPeriods),UppContrLimits(DBMParameters.EMAPreviousPeriods) As Double
            Calculate=New DBMResult
            If SubstractDBMPoint IsNot PrevSubstractDBMPoint Then ' Can we reuse stored results?
                DBMPredictions.Clear 'No, so clear results
                PrevSubstractDBMPoint=SubstractDBMPoint
            End If
            For CorrelationCounter=0 To DBMParameters.CorrelationPreviousPeriods
                If Calculate.DBMPrediction Is Nothing Or (IsInputDBMPoint And Calculate.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                    For EMACounter=DBMParameters.EMAPreviousPeriods To 0 Step -1
                        CalcTimestamp=DateAdd("s",-(EMACounter+CorrelationCounter)*DBMParameters.CalculationInterval,Timestamp)
                        If DBMPredictions.ContainsKey(CalcTimestamp) Then ' From cache
                            DBMPrediction=DBMPredictions.Item(CalcTimestamp).ShallowCopy
                        Else ' Calculate data
                            For PatternCounter=DBMParameters.ComparePatterns To 0 Step -1
                                Patterns(DBMParameters.ComparePatterns-PatternCounter)=DBMDataManager.Value(DateAdd("d",-PatternCounter*7,CalcTimestamp))
                                If SubstractDBMPoint IsNot Nothing Then
                                    Patterns(DBMParameters.ComparePatterns-PatternCounter)-=SubstractDBMPoint.DBMDataManager.Value(DateAdd("d",-PatternCounter*7,CalcTimestamp))
                                End If
                            Next PatternCounter
                            DBMPrediction.Calculate(Patterns)
                            Do While DBMPredictions.Count>=DBMParameters.MaxDBMPointCacheSize ' Limit cache size
                                DBMPredictions.Remove(DBMPredictions.ElementAt(CInt(Math.Floor(Rnd*DBMPredictions.Count))).Key) ' Remove random cached value
                            Loop
                            DBMPredictions.Add(CalcTimestamp,DBMPrediction.ShallowCopy) ' Add to cache
                        End If
                        MeasValues(EMACounter)=DBMPrediction.MeasValue
                        PredValues(EMACounter)=DBMPrediction.PredValue
                        LowContrLimits(EMACounter)=DBMPrediction.LowContrLimit
                        UppContrLimits(EMACounter)=DBMPrediction.UppContrLimit
                    Next EMACounter
                    Calculate.Calculate(CorrelationCounter,DBMMath.CalculateExpMovingAvg(MeasValues),DBMMath.CalculateExpMovingAvg(PredValues),DBMMath.CalculateExpMovingAvg(LowContrLimits),DBMMath.CalculateExpMovingAvg(UppContrLimits))
                End If
            Next CorrelationCounter
            Return Calculate
        End Function

    End Class

End Namespace
