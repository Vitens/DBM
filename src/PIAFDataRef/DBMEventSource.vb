Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2020  J.H. Fitié, Vitens N.V.
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
Imports System.Collections.Generic
Imports System.DateTime
Imports System.Threading
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMEventSource
    Inherits AFEventSource


    Private Shared ValueResults As New List(Of ValueResult)
    Private LastTime As AFTime = Now


    Private Structure ValueResult

      Public Attribute As AFAttribute
      Public Value As AFValue

    End Structure


    Private Structure RetrievalInfo

      Public StartTime As AFTime
      Public EndTime As AFTime
      Public Attribute As AFAttribute

    End Structure


    Private Shared Sub RetrieveValues(RetrievalInfo As Object)

      Dim ValueResult As ValueResult
      Dim Value As AFValue

      ValueResult.Attribute = DirectCast(RetrievalInfo, RetrievalInfo).Attribute

      For Each Value In
        DirectCast(RetrievalInfo, RetrievalInfo).Attribute.GetValues(
        New AFTimeRange(DirectCast(RetrievalInfo, RetrievalInfo).StartTime,
        DirectCast(RetrievalInfo, RetrievalInfo).EndTime), 0, Nothing)

        ValueResult.Value = Value
        ValueResults.Add(ValueResult)

      Next

    End Sub


    Protected Overrides Function GetEvents As Boolean

      Dim EvalTime As AFTime = Now
      Dim RetrievalInfo As RetrievalInfo
      Dim Attribute As AFAttribute
      Dim Threads As New List(Of Thread)
      Dim Thread As Thread
      Dim ValueResult As ValueResult

      EvalTime = New AFTime(AlignPreviousInterval(EvalTime.UtcSeconds,
        CalculationInterval))

      RetrievalInfo.StartTime = LastTime
      RetrievalInfo.EndTime = EvalTime

      For Each Attribute In MyBase.Signups

        If Attribute IsNot Nothing And LastTime < EvalTime Then

          RetrievalInfo.Attribute = Attribute
          Threads.Add(New Thread(
            New ParameterizedThreadStart(AddressOf RetrieveValues)))
          Threads(Threads.Count-1).Start(RetrievalInfo)

        End If

      Next

      For Each Thread In Threads
        Thread.Join
      Next

      For Each ValueResult In ValueResults
        MyBase.PublishEvent(ValueResult.Attribute,
          New AFDataPipeEvent(AFDataPipeAction.Add, ValueResult.Value))
      Next
      ValueResults.Clear

      LastTime = EvalTime

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)

      LastTime = Nothing
      ValueResults = Nothing

    End Sub


  End Class


End Namespace
