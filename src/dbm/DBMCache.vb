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
Imports Vitens.DynamicBandwidthMonitor.DBMMath


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMCache


    Const KeyNothingAlternative As Integer = Int32.MinValue


    Private CacheItems As New Dictionary(Of Object, DBMCacheItem)


    Private Function ValidatedKey(Key As Object) As Object

      ' Keys cannot be Nothing, as the dictionary requires an instance of an
      ' object to be used for the hash code.

      If Key Is Nothing Then
        Return KeyNothingAlternative
      Else
        Return Key
      End If

    End Function


    Public Sub Clear

      CacheItems.Clear

    End Sub


    Public Sub AddItem(Key As Object, Item As Object)

      CacheItems.Add(ValidatedKey(Key), New DBMCacheItem(Item))

    End Sub


    Public Function HasItem(Key As Object) As Boolean

      Return CacheItems.ContainsKey(ValidatedKey(Key))

    End Function


    Public Function GetItem(Key As Object) As Object

      Dim CacheItem As New DBMCacheItem

      If Not CacheItems.TryGetValue(ValidatedKey(Key), CacheItem) Then
        CacheItem = New DBMCacheItem ' Not found, return default empty value
      End If

      Return CacheItem.Item

    End Function


  End Class


End Namespace
