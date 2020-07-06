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
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMEventSource
    Inherits AFEventSource


    Private PreviousInterval As AFTime = AFTime.MinValue
    Private PreviousSnapshots As New Dictionary(Of AFAttribute, AFTime)


    Protected Overrides Function GetEvents As Boolean

      ' The GetEvents method is designed to get data pipe events from the System
      ' of record.

      Dim CurrentInterval As AFTime = Now
      Dim Attribute As AFAttribute
      Dim Snapshot As AFValue

      ' Determine the start of the current time interval. For this event source,
      ' we only return at most one new event per attribute.
      CurrentInterval = New AFTime(AlignPreviousInterval(
        CurrentInterval.UtcSeconds, CalculationInterval))

      ' Only check for new events once per calculation interval.
      If PreviousInterval < CurrentInterval Then

        ' Iterate over all signed up attributes in this data pipe. We use serial
        ' processing of the attributes, since we need to reuse the DBM object to
        ' store all cached data. This means that for some services signing up to
        ' many attributes, initialization might take some time. After data is
        ' retrieved for all attributes for the first time, only small amounts of
        ' new data are needed for all consequent evaluations, greatly speeding
        ' them up.
        For Each Attribute In MyBase.Signups

          ' Check if we need to perform an action on this attribute. This is
          ' only needed if the attribute is an instance of an object.
          If Attribute IsNot Nothing Then

            If Not PreviousSnapshots.ContainsKey(Attribute) Then

              ' Store the initial previous snapshot for this attribute.
              PreviousSnapshots.Add(Attribute, AFTime.MinValue)

            Else

              ' In a new interval, retrieve snapshot value. If newer, store
              ' snapshot timestamp as previous snapshot timestamp for this
              ' attribute, and publish snapshot value as new event in the data
              ' pipe.
              Snapshot = Attribute.GetValue
              If Snapshot.Timestamp > PreviousSnapshots.Item(Attribute) Then
                PreviousSnapshots.Item(Attribute) = Snapshot.Timestamp
                MyBase.PublishEvent(Attribute,
                  New AFDataPipeEvent(AFDataPipeAction.Add, Snapshot))
              End

            End If

          End If

        Next

        ' Store current interval as previous interval.
        PreviousInterval = CurrentInterval

      End If

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)
    End Sub


  End Class


End Namespace
