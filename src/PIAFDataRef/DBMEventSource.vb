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


    Private Shared Events As New List(Of Event)
    Private LastTime As AFTime = Now


    Private Structure Event

      ' This structure is used to store events retrieved in a separate thread
      ' per attribute.

      Public Attribute As AFAttribute
      Public Value As AFValue

    End Structure


    Private Structure RetrievalInfo

      ' This structure is used to pass information required to retrieve values
      ' to the thread started for each attribute.

      Public StartTime As AFTime
      Public EndTime As AFTime
      Public Attribute As AFAttribute

    End Structure


    Private Shared Sub RetrieveValues(RetrievalInfo As Object)

      Dim Event As Event
      Dim Value As AFValue

      Event.Attribute = DirectCast(RetrievalInfo, RetrievalInfo).Attribute

      ' Call the GetValues method for the attribute to retrieve values to store
      ' in the shared events list.
      For Each Value In
        DirectCast(RetrievalInfo, RetrievalInfo).Attribute.GetValues(
        New AFTimeRange(DirectCast(RetrievalInfo, RetrievalInfo).StartTime,
        DirectCast(RetrievalInfo, RetrievalInfo).EndTime), 0, Nothing)

        Event.Value = Value
        Events.Add(Event)

      Next

    End Sub


    Protected Overrides Function GetEvents As Boolean

      ' The GetEvents method is designed to get data pipe events from the System
      ' of record.

      Dim EvalTime As AFTime = Now
      Dim RetrievalInfo As RetrievalInfo
      Dim Attribute As AFAttribute
      Dim Threads As New List(Of Thread)
      Dim Thread As Thread
      Dim Event As Event

      ' The evaluation time is the current time aligned to the previous
      ' calculation interval.
      EvalTime = New AFTime(AlignPreviousInterval(EvalTime.UtcSeconds,
        CalculationInterval))

      RetrievalInfo.StartTime = LastTime
      RetrievalInfo.EndTime = EvalTime

      For Each Attribute In MyBase.Signups

        ' Only check for new events once per calculation interval.
        If Attribute IsNot Nothing And LastTime < EvalTime Then

          RetrievalInfo.Attribute = Attribute
          ' Start a new thread for each attribute, passing information about the
          ' attribute and the time range to retrieve data for.
          Threads.Add(New Thread(
            New ParameterizedThreadStart(AddressOf RetrieveValues)))
          Threads(Threads.Count-1).Start(RetrievalInfo)

        End If

      Next

      ' Wait for all threads to finish.
      For Each Thread In Threads
        Thread.Join
      Next

      ' Publish all events to the data pipe.
      For Each Event In Events
        MyBase.PublishEvent(Event.Attribute,
          New AFDataPipeEvent(AFDataPipeAction.Add, Event.Value))
      Next
      Events.Clear

      LastTime = EvalTime

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)

      ' Clean up objects.

      Events = Nothing
      LastTime = Nothing

    End Sub


  End Class


End Namespace
