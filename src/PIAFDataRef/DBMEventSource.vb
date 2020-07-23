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


    Private Shared PreviousSnapshots As New Dictionary(Of AFAttribute, AFTime)
    Private Shared DataPipeEvents As New List(Of DataPipeEvent)
    Private PreviousInterval As AFTime = AFTime.MinValue


    Private Structure DataPipeEvent

      ' This structure is used to store events retrieved in a separate thread
      ' per attribute.

      Public Attribute As AFAttribute
      Public Value As AFValue

    End Structure


    Private Shared Sub RetrieveEvent(Attribute As Object)

      Dim DataPipeEvent As DataPipeEvent
      Dim Snapshot As AFValue

      ' In a new interval, retrieve snapshot value.
      DataPipeEvent.Attribute = DirectCast(Attribute, AFAttribute)
      Snapshot = DataPipeEvent.Attribute.GetValue

      Monitor.Enter(PreviousSnapshots) ' Lock
      Try
        If Not PreviousSnapshots.ContainsKey(DataPipeEvent.Attribute) Then
          ' Store the initial previous snapshot timestamp for this attribute.
          PreviousSnapshots.Add(DataPipeEvent.Attribute, Snapshot.Timestamp)
        Else
          ' If newer, store snapshot timestamp as previous snapshot timestamp
          ' for this attribute.
          If Snapshot.Timestamp >
            PreviousSnapshots.Item(DataPipeEvent.Attribute) Then
            PreviousSnapshots.Item(DataPipeEvent.Attribute) = Snapshot.Timestamp
          Else
            Snapshot = Nothing ' There is no new event.
          End If
        End If
      Finally
        Monitor.Exit(PreviousSnapshots)
      End Try

      If Snapshot IsNot Nothing Then
        ' Store new snapshot in shared events list.
        DataPipeEvent.Value = Snapshot
        Monitor.Enter(DataPipeEvents) ' Lock
        Try
          DataPipeEvents.Add(DataPipeEvent)
        Finally
          Monitor.Exit(DataPipeEvents)
        End Try
      End If

    End Sub


    Protected Overrides Function GetEvents As Boolean

      ' The GetEvents method is designed to get data pipe events from the System
      ' of record.

      Dim CurrentInterval As AFTime = Now
      Dim Attribute As AFAttribute
      Dim Threads As New List(Of Thread)
      Dim Thread As Thread
      Dim DataPipeEvent As DataPipeEvent

      ' Determine the start of the current time interval. For this event source,
      ' we only return at most one new event per attribute.
      CurrentInterval = New AFTime(AlignPreviousInterval(
        CurrentInterval.UtcSeconds, CalculationInterval))

      ' Only check for new events once per calculation interval.
      If PreviousInterval < CurrentInterval Then

        ' Iterate over all signed up attributes in this data pipe. We use
        ' parallel processing of the attributes, since we need to reuse the DBM
        ' object to store all cached data. If the shared DBM object cannot be
        ' locked after half a calculation interval, use the attribute-specific
        ' instance. If the attribute-specific DBM object cannot be locked as
        ' well, use a new object. For these last two cases, we need to query the
        ' attributes in parallel, since we don't want to have the waiting step
        ' for each attribute in serial. So since data is essentially retrieved
        ' in serial, this means that for some services signing up to many
        ' attributes, initialization might take some time. After data is
        ' retrieved for all attributes for the first time, only small amounts of
        ' new data are needed for all consequent evaluations, greatly speeding
        ' them up.
        For Each Attribute In MyBase.Signups

          ' Check if we need to perform an action on this attribute. This is
          ' only needed if the attribute is an instance of an object.
          If Attribute IsNot Nothing Then

            ' Start a new thread for each attribute, passing the attribute.
            Threads.Add(New Thread(
              New ParameterizedThreadStart(AddressOf RetrieveEvent)))
            Threads(Threads.Count-1).Start(Attribute)

          End If

        Next

        ' Wait for all threads to finish.
        For Each Thread In Threads
          Thread.Join
        Next

        ' Publish all events to the data pipe and clear the events list.
        Monitor.Enter(DataPipeEvents) ' Lock
        Try

          For Each DataPipeEvent In DataPipeEvents
            MyBase.PublishEvent(DataPipeEvent.Attribute,
              New AFDataPipeEvent(AFDataPipeAction.Add, DataPipeEvent.Value))
          Next
          DataPipeEvents.Clear

        Finally
          Monitor.Exit(DataPipeEvents)
        End Try

        ' Store current interval as previous interval.
        PreviousInterval = CurrentInterval

      End If

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)
    End Sub


  End Class


End Namespace
