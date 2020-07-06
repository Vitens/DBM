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


    Private PreviousEndTimestamps As New Dictionary(Of AFAttribute, AFTime)


    Protected Overrides Function GetEvents As Boolean

      ' The GetEvents method is designed to get data pipe events from the System
      ' of record.

      Dim Attribute As AFAttribute
      Dim EndTimestamp As AFTime
      Dim Value As AFValue

      ' Iterate over all signed up attributes in this data pipe.
      For Each Attribute In MyBase.Signups

        ' Check if we need to perform an action on this attribute. This is only
        ' needed if the attribute and it's parent (input source) are an instance
        ' of an object.
        If Attribute IsNot Nothing AndAlso Attribute.Parent IsNot Nothing Then

          ' Use parent attribute snapshot as snapshot timestamp for this
          ' attribute, as this is the input source. Aligned to the next
          ' calculation interval. Do not query the attribute itself, as this
          ' would trigger a DBM calculation while we only need the timestamp
          ' here.
          EndTimestamp = New AFTime(AlignNextInterval(Attribute.Parent.GetValue.
            Timestamp.UtcSeconds, CalculationInterval))

          ' If there is no previous end timestamp for this attribute yet, use
          ' the one just retrieved.
          If Not PreviousEndTimestamps.ContainsKey(Attribute) Then
            PreviousEndTimestamps.Add(Attribute, EndTimestamp)
          End If

          ' Check if there are new values to retrieve.
          If PreviousEndTimestamps.Item(Attribute) < EndTimestamp Then

            ' Retrieve values since last event.
            For Each Value In Attribute.GetValues(New AFTimeRange(
              PreviousEndTimestamps.Item(Attribute), EndTimestamp), 0, Nothing)

              ' Publish new values as events in the data pipe.
              MyBase.PublishEvent(Attribute,
                New AFDataPipeEvent(AFDataPipeAction.Add, Value))

            Next

            ' Store end timestamp as previous end timestamp for this attribute.
            PreviousEndTimestamps.Item(Attribute) = EndTimestamp

          End If

        End If

      Next

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)
    End Sub


  End Class


End Namespace
