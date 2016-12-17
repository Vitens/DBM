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

        Public DataManager As DBMDataManager
        Private Predictions As New Collections.Generic.Dictionary(Of DateTime,DBMPrediction)
        Private PredictionsSubstractPoint As DBMPoint

        Public Sub New(PointDriver As DBMPointDriver)
            DataManager=New DBMDataManager(PointDriver)
        End Sub

        Public Function Result(Timestamp As DateTime,IsInputDBMPoint As Boolean,HasCorrelationDBMPoint As Boolean,Optional SubstractPoint As DBMPoint=Nothing) As DBMResult
            Dim CorrelationCounter,EMACounter,PatternCounter As Integer
            Dim PredictionTimestamp As DateTime
            Dim Prediction As New DBMPrediction
            Dim Patterns(DBMParameters.ComparePatterns),MeasuredValues(DBMParameters.EMAPreviousPeriods),PredictedValues(DBMParameters.EMAPreviousPeriods),LowerControlLimits(DBMParameters.EMAPreviousPeriods),UpperControlLimits(DBMParameters.EMAPreviousPeriods) As Double
            Result=New DBMResult
            If SubstractPoint IsNot PredictionsSubstractPoint Then ' Can we reuse stored results?
                Predictions.Clear 'No, so clear results
                PredictionsSubstractPoint=SubstractPoint
            End If
            For CorrelationCounter=0 To DBMParameters.CorrelationPreviousPeriods
                If Result.Prediction Is Nothing Or (IsInputDBMPoint And Result.Factor<>0 And HasCorrelationDBMPoint) Or Not IsInputDBMPoint Then
                    For EMACounter=0 To DBMParameters.EMAPreviousPeriods
                        PredictionTimestamp=DateAdd("s",-(DBMParameters.EMAPreviousPeriods-EMACounter+CorrelationCounter)*DBMParameters.CalculationInterval,Timestamp)
                        If Predictions.ContainsKey(PredictionTimestamp) Then ' From cache
                            Prediction=Predictions.Item(PredictionTimestamp).ShallowCopy
                        Else ' Calculate data
                            For PatternCounter=0 To DBMParameters.ComparePatterns
                                Patterns(PatternCounter)=DataManager.Value(DateAdd("d",-(DBMParameters.ComparePatterns-PatternCounter)*7,PredictionTimestamp))
                                If SubstractPoint IsNot Nothing Then
                                    Patterns(PatternCounter)-=SubstractPoint.DataManager.Value(DateAdd("d",-(DBMParameters.ComparePatterns-PatternCounter)*7,PredictionTimestamp))
                                End If
                            Next PatternCounter
                            Prediction.Calculate(Patterns)
                            Do While Predictions.Count>=DBMParameters.MaxPointPredictions ' Limit cache size
                                Predictions.Remove(Predictions.ElementAt(CInt(Math.Floor(Rnd*Predictions.Count))).Key) ' Remove random cached value
                            Loop
                            Predictions.Add(PredictionTimestamp,Prediction.ShallowCopy) ' Add to cache
                        End If
                        MeasuredValues(EMACounter)=Prediction.MeasuredValue
                        PredictedValues(EMACounter)=Prediction.PredictedValue
                        LowerControlLimits(EMACounter)=Prediction.LowerControlLimit
                        UpperControlLimits(EMACounter)=Prediction.UpperControlLimit
                    Next EMACounter
                    Result.Calculate(CorrelationCounter,DBMMath.CalculateExpMovingAvg(MeasuredValues),DBMMath.CalculateExpMovingAvg(PredictedValues),DBMMath.CalculateExpMovingAvg(LowerControlLimits),DBMMath.CalculateExpMovingAvg(UpperControlLimits))
                End If
            Next CorrelationCounter
            Return Result
        End Function

    End Class

End Namespace
