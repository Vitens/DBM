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
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMPointDriverAbstract


    ' DBM drivers should inherit from this base class. In Sub New,
    ' MyBase.New(point) should be called. RetrieveData is used to retrieve and
    ' store data in bulk from a source using the DataStore.AddData method. Use
    ' DataStore.GetData to fetch results from memory.


    Private _point As Object
    Private _dataStore As New DBMDataStore ' In-memory data
    Private _lock As New Object ' Object for exclusive lock on critical section.
    Private _previousStartTimestamp As DateTime = DateTime.MaxValue
    Private _previousEndTimestamp As DateTime


    Public Property Point As Object
      Get
        Return _point
      End Get
      Set(value As Object)
        _point = value
      End Set
    End Property


    Public Property DataStore As DBMDataStore
      Get
        Return _dataStore
      End Get
      Set(value As DBMDataStore)
        _dataStore = value
      End Set
    End Property


    Public Sub New(point As Object)

      ' When inheriting from this base class, call MyBase.New(Point) from Sub
      ' New to store the unique identifier in the Point object.

      Me.Point = point

    End Sub


    Public MustOverride Overrides Function ToString As String
      ' Return the name of this point as a string.


    Public Overridable Function SnapshotTimestamp As DateTime

      ' Return the latest data timestamp (snapshot) for which the source of data
      ' has information available. If there is no limit, return
      ' DateTime.MaxValue.

      Return DateTime.MaxValue

    End Function


    Public Function CalculationTimestamp As DateTime

      ' Return the latest possible calculation timestamp for which the source of
      ' data has information available. Compensated for exponential moving
      ' average (EMA) time shift.

      Return SnapshotTimestamp.AddSeconds(EMATimeOffset(EMAPreviousPeriods+1))

    End Function


    Public MustOverride Sub PrepareData(startTimestamp As DateTime,
      endTimestamp As DateTime)
      ' Must retrieve and store data in bulk for the passed time range from a
      ' source of data, to be used in the DataStore.GetData method. Use
      ' DataStore.AddData to store data in memory.


    Public Sub RetrieveData(startTimestamp As DateTime,
      endTimestamp As DateTime)

      Dim snapshot As DateTime = SnapshotTimestamp

      ' Data preparation timestamps
      startTimestamp = DataPreparationTimestamp(startTimestamp)
      endTimestamp = PreviousInterval(endTimestamp)

      ' If set, never retrieve values beyond the snapshot time aligned to the
      ' next interval.
      If startTimestamp > snapshot Then startTimestamp = NextInterval(snapshot)
      If endTimestamp > snapshot Then endTimestamp = NextInterval(snapshot)

      ' Exit this sub if there is no data to retrieve or when the start
      ' timestamp is not before the end timestamp.
      If Not startTimestamp < endTimestamp Then Exit Sub

      Monitor.Enter(_lock) ' Block
      Try

        ' Determine what data stored in memory can be reused, what needs to be
        ' removed, and what needs to be retrieved from the data source and
        ' stored in memory. Here is a simplified overview of how the different
        ' cases (S(tart)-E(nd)) are handled, compared to the time range stored
        ' previously (PS-PE):
        '             PS*==========*PE  S...PS E...PE Action
        '   Case 1:     S==========E      =      =    Do nothing
        '   Case 2:  S++|==========E      <      =    Add backward
        '   Case 3:     ---S=======E      >      =    Do nothing
        '   Case 4:     S=======E---      =      <    Do nothing
        '   Case 5:  S++|=======E---      <      <    Remove forward, add backwd
        '   Case 6:     ---S====E---      >      <    Do nothing
        '   Case 7:     S==========|++E   =      >    Add forward
        '   Case 8:  S++|==========|++E   <      >    Clear all
        '   Case 9:     ---S=======|++E   >      >    Remove backward, add forwd
        '   Case 10: S==E a) or b) S==E   E<=PS S>=PE Clear all
        If (startTimestamp < _previousStartTimestamp And
          endTimestamp > _previousEndTimestamp) Or
          endTimestamp <= _previousStartTimestamp Or
          startTimestamp >= _previousEndTimestamp Then ' Cases 8, 10a), 10b)
          DataStore.ClearData ' Clear all
          _previousStartTimestamp = startTimestamp
          _previousEndTimestamp = endTimestamp
        Else If startTimestamp >= _previousStartTimestamp And
          endTimestamp <= _previousEndTimestamp Then ' Cases 1, 3, 4, 6
          Exit Sub ' Do nothing
        Else If startTimestamp < _previousStartTimestamp Then ' Cases 2, 5
          If endTimestamp < _previousEndTimestamp Then ' Case 5
            Do While endTimestamp < _previousEndTimestamp ' Remove forward
              _previousEndTimestamp = _previousEndTimestamp.
                AddSeconds(-CalculationInterval)
              DataStore.RemoveData(_previousEndTimestamp)
            Loop
          End If
          endTimestamp = _previousStartTimestamp ' Add backward
          _previousStartTimestamp = startTimestamp
        Else If endTimestamp > _previousEndTimestamp Then ' Cases 7, 9
          If startTimestamp > _previousStartTimestamp Then ' Case 9
            Do While _previousStartTimestamp < startTimestamp ' Remove backward
              DataStore.RemoveData(_previousStartTimestamp)
              _previousStartTimestamp = _previousStartTimestamp.
                AddSeconds(CalculationInterval)
            Loop
          End If
          startTimestamp = _previousEndTimestamp ' Add forward
          _previousEndTimestamp = endTimestamp
        End If

        Try
          ' Retrieve all data from the data source. Will pass start and end
          ' timestamps. The driver can then prepare the dataset for which
          ' calculations are required in the next step. The (aligned) end time
          ' itself is excluded.
          PrepareData(startTimestamp, endTimestamp)
        Catch ex As Exception
          DBM.Logger.LogWarning(ex.ToString, Me.ToString)
        End Try

      Finally
        Monitor.Exit(_lock) ' Unblock
      End Try

    End Sub


  End Class


End Namespace
