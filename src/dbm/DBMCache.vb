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
Imports System.DateTime
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMCache


    Const KeyNothingAlternative = Int32.MinValue


    Private MaximumItems As Integer
    Private ItemStaleInterval As Integer
    Private Lock As New Object
    Private NextStaleItemsCheck As DateTime
    Private CacheItems As New Dictionary(Of Object, DBMCacheItem)


    Public Sub New(Optional MaximumItems As Integer = 0,
      Optional ItemStaleInterval As Integer = 0)

      Me.MaximumItems = MaximumItems ' Use default (0) value for unlimited items
      Me.ItemStaleInterval = ItemStaleInterval ' Default val. (0) for no timeout
      UpdateCheck

    End Sub


    Private Sub UpdateCheck

      ' Update timestamp after which stale items have to be checked.

      If ItemStaleInterval > 0 Then
      
        ' SyncLock: Access to this method has to be synchronized because the
        '           NextStaleItemsCheck variable is modified here and should be
        '           available in the RemoveStaleItems method.

        Monitor.Enter(Lock) ' Request the lock, and block until obtained.
        Try

          NextStaleItemsCheck = NextInterval(Now)

        Finally
          Monitor.Exit(Lock) ' Ensure that the lock is released.
        End Try

      Else

        NextStaleItemsCheck = DateTime.MaxValue ' Never

      End If

    End Sub


    Private Sub RemoveStaleItems

      Dim Pair As KeyValuePair(Of Object, DBMCacheItem)
      Dim StaleItems As New List(Of Object)
      Dim StaleItem As Object

      ' SyncLock: Access to this method does not have to be synchronized because
      '           this private method is only called from the AddItem method
      '           where the lock is already obtained.

      If Now >= NextStaleItemsCheck Then

        UpdateCheck

        For Each Pair In CacheItems
          If Pair.Value.IsStale Then ' Find stale items
            StaleItems.Add(Pair.Key)
          End If
        Next

        For Each StaleItem In StaleItems
          CacheItems.Remove(StaleItem) ' Remove stale items
        Next

      End If

    End Sub


    Private Sub LimitSize

      ' Limit number of cached forecast results per point. Cache size is limited
      ' using random eviction policy.

      ' SyncLock: Access to this method does not have to be synchronized because
      '           this private method is only called from the AddItem method
      '           where the lock is already obtained. Items can be removed from
      '           the dictionary in this method.

      If MaximumItems > 0 Then
        If CacheItems.Count >= MaximumItems Then
          CacheItems.Remove(CacheItems.ElementAt(
            RandomNumber(0, CacheItems.Count-1)).Key)
        End If
      End If

    End Sub


    Private Function ValidatedKey(Key As Object) As Object

      ' Keys cannot be Nothing, as the dictionary requires an instance of an
      ' object to be used for the hash code.

      If Key Is Nothing Then
        Return KeyNothingAlternative
      Else
        Return Key
      End If

    End Function


    Public Function HasItem(Key As Object) As Boolean

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        Return CacheItems.ContainsKey(ValidatedKey(Key))

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Function


    Public Sub AddItem(Key As Object, Item As Object)

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        RemoveStaleItems
        LimitSize

        CacheItems.Add(ValidatedKey(Key),
          New DBMCacheItem(Item, ItemStaleInterval))

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Public Sub AddItemIfNotExists(Key As Object, Item As Object)

      ' SyncLock: Access to this method has to be synchronized because multiple
      '           threads could try to add a new item simultaneously. HasItem
      '           will then return False on each thread and the same item might
      '           be added multiple times through AddItem.

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        If Not HasItem(Key) Then AddItem(Key, Item)

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Sub


    Public Function GetItem(Key As Object) As Object

      Dim CacheItem As New DBMCacheItem

      Monitor.Enter(Lock) ' Request the lock, and block until it is obtained.
      Try

        If Not CacheItems.TryGetValue(ValidatedKey(Key), CacheItem) Then
          CacheItem = New DBMCacheItem ' Not found, return default empty value
        End If

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

      Return CacheItem.GetItem

    End Function


  End Class


End Namespace
