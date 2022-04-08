Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2022  J.H. Fitié, Vitens N.V.
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


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMEventSource
    Inherits AFEventSource


    Const UpdateSeconds As Integer = 60 ' Time limiter for update checking.


    Private _previousTimestamp As DateTime
    Private _previousSnapshots As New Dictionary(Of AFAttribute, DateTime)


    Protected Overrides Function GetEvents As Boolean

      ' The GetEvents method is designed to get data pipe events from the System
      ' of record.

      Dim attribute As AFAttribute
      Dim snapshot As AFValue

      ' Limit update checking to once every minute.
      If (Now-_previousTimestamp).TotalSeconds < UpdateSeconds Then Return False
      _previousTimestamp = Now

      ' Iterate over all signed up attributes in this data pipe. For some
      ' services signing up to many attributes, initialization might take some
      ' time. After data is retrieved for all attributes for the first time,
      ' only small amounts of new data are needed for all consequent
      ' evaluations, greatly speeding them up.
      For Each attribute In MyBase.Signups

        ' Check if we need to perform an action on this attribute. This is only
        ' needed if the attribute is an instance of an object.
        If attribute IsNot Nothing Then

          snapshot = attribute.GetValue ' Retrieve latest value.

          If _previousSnapshots.ContainsKey(attribute) Then
            ' If newer, store snapshot timestamp as previous snapshot timestamp
            ' for this attribute and publish new snapshot to the data pipe.
            If snapshot.Timestamp.LocalTime >
              _previousSnapshots.Item(attribute) Then
              _previousSnapshots.Item(attribute) = snapshot.Timestamp.LocalTime
              MyBase.PublishEvent(attribute,
                New AFDataPipeEvent(AFDataPipeAction.Add, snapshot))
            End If
          Else
            ' Store the initial previous snapshot timestamp for this attribute.
            _previousSnapshots.Add(attribute, snapshot.Timestamp.LocalTime)
          End If

        End If

      Next attribute

      Return False ' Required.

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)
    End Sub


  End Class


End Namespace
