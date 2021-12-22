Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2021  J.H. Fitié, Vitens N.V.
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
Imports System.Double
Imports System.Threading


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMDataStore


    Private _lock As New Object ' Object for exclusive lock on critical section.
    Private DataStore As New Dictionary(Of DateTime, Double) ' In-memory data


    Public Sub AddData(Timestamp As DateTime, Data As Object)

      ' Make sure that the retrieved data type is a Double and also that the
      ' timestamp is not already stored in memory (could happen because of DST
      ' time overlap).

      If TypeOf Data Is Double Then

        Monitor.Enter(_lock) ' Block
        Try

          If Not DataStore.ContainsKey(Timestamp) Then
            DataStore.Add(Timestamp, DirectCast(Data, Double))
          End If

        Finally
          Monitor.Exit(_lock) ' Unblock
        End Try

      End If

    End Sub


    Public Sub RemoveData(Timestamp As DateTime)

      Monitor.Enter(_lock) ' Block
      Try

        DataStore.Remove(Timestamp)

      Finally
        Monitor.Exit(_lock) ' Unblock
      End Try

    End Sub


    Public Sub ClearData

      Monitor.Enter(_lock) ' Block
      Try

        DataStore.Clear ' Clear all

      Finally
        Monitor.Exit(_lock) ' Unblock
      End Try

    End Sub


    Public Function GetData(Timestamp As DateTime) As Double

      ' Retrieves data from the DataStore dictionary. If there is no data for
      ' the timestamp, return Not a Number.

      Monitor.Enter(_lock) ' Block
      Try

        GetData = Nothing
        If DataStore.TryGetValue(Timestamp, GetData) Then ' In dictionary.
          Return GetData ' Return value from dictionary.
        Else
          Return NaN ' No data in dictionary for timestamp, return Not a Number.
        End If

      Finally
        Monitor.Exit(_lock) ' Unblock
      End Try

    End Function


  End Class


End Namespace
