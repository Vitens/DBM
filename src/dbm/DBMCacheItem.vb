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
Imports System.DateTime
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMCacheItem


    Private Item As Object
    Private ItemStaleInterval As Integer
    Private Lock As New Object
    Private TimeOut As DateTime


    Public Sub New(Optional Item As Object = Nothing,
      Optional ItemStaleInterval As Integer = 0)

      Me.Item = Item
      Me.ItemStaleInterval = ItemStaleInterval ' Default val. (0) for no timeout
      UpdateTimeOut

    End Sub


    Private Sub UpdateTimeOut

      ' Update timestamp after which item turns stale.

      If ItemStaleInterval > 0 Then
      
        ' SyncLock: Access to this method has to be synchronized because the
        '           TimeOut variable is modified here and should be available in
        '           the IsStale method.

        Monitor.Enter(Lock) ' Request the lock, and block until obtained.
        Try

          TimeOut = NextInterval(Now, ItemStaleInterval)

        Finally
          Monitor.Exit(Lock) ' Ensure that the lock is released.
        End Try

      Else

        TimeOut = DateTime.MaxValue ' Never

      End If

    End Sub


    Public Function GetItem As Object

      UpdateTimeOut
      Return Item

    End Function


    Public Function IsStale As Boolean

      ' Returns true if this item has turned stale.

      ' SyncLock: Access to this method has to be synchronized because the
      '           TimeOut variable should be available here and is modified in
      '           the UpdateTimeOut method.

      Monitor.Enter(Lock) ' Request the lock, and block until obtained.
      Try

        Return Now >= TimeOut ' Returns True if the point has timed out.

      Finally
        Monitor.Exit(Lock) ' Ensure that the lock is released.
      End Try

    End Function


  End Class


End Namespace
