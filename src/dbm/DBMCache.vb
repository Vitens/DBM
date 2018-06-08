Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2018  J.H. Fitié, Vitens N.V.
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
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMCache


    Private MaximumItems As Integer
    Private Lock As New Object
    Private CacheItems As New Dictionary(Of Object, DBMCacheItem)


    Public Sub New(MaximumItems As Integer)

      Me.MaximumItems = MaximumItems

    End Sub


    Public Sub Clear

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        CacheItems.Clear

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Private Sub LimitSize

      ' Limit number of cached forecast results per point. Cache size is limited
      ' using random eviction policy.

      ' SyncLock: Access to this method does not have to be synchronized because
      '           this private method is only called from the AddItem method
      '           where the lock is already obtained. Items can be removed from
      '           the dictionary in this method.

      If CacheItems.Count >= MaximumItems Then
        CacheItems.Remove(CacheItems.ElementAt(
          RandomNumber(0, CacheItems.Count-1)).Key)
      End If

    End Sub


    Public Sub AddItem(Index As Object, Item As Object)

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        LimitSize

        CacheItems.Add(Index, New DBMCacheItem(Item))

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Public Function GetItem(Index As Object) As Object

      Dim CacheItem As New DBMCacheItem

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        If Not CacheItems.TryGetValue(Index, CacheItem) Then
          CacheItem = New DBMCacheItem ' Not found, return default empty value
        End If

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

      Return CacheItem.Item

    End Function


  End Class


End Namespace
