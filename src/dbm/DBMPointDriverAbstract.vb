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
Imports System.Double
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMPointDriverAbstract


    ' DBM drivers should inherit from this base class. In Sub New,
    ' MyBase.New(Point) should be called. GetValues is used to retrieve, store,
    ' and return values in bulk from a source of data using the AddValue method.


    Public Point As Object
    Private Values As New Dictionary(Of DateTime, Double) ' In-memory data store
    Private PreviousStartTimestamp As DateTime = DateTime.MaxValue
    Private PreviousEndTimestamp As DateTime = DateTime.MinValue


    Public Sub New(Point As Object)

      ' When inheriting from this base class, call MyBase.New(Point) from Sub
      ' New to store the unique identifier in the Point object.

      Me.Point = Point

    End Sub


    Private Function AddValue(Timestamp As DateTime, Value As Object)

      ' Make sure that the retrieved data type is a Double and also that the
      ' timestamp is not already stored in memory (could happen because of DST
      ' time overlap).
      If TypeOf Value Is Double AndAlso
        Not Values.ContainsKey(Timestamp) Then
        Values.Add(Timestamp, DirectCast(Value, Double))
      End If

    End Function


    Private MustOverride Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)
      ' Retrieve and store values in bulk for the passed time range from a
      ' source of data, to be used in the GetData method. Use AddValue to store
      ' values in memory.


    Public Function GetValues(StartTimestamp As DateTime,
      EndTimestamp As DateTime) As Dictionary(Of DateTime, Double)

      Dim ShouldPrepareData As Boolean = True

      ' Data preparation timestamps
      StartTimestamp = DataPreparationTimestamp(StartTimestamp)
      EndTimestamp = AlignTimestamp(EndTimestamp, CalculationInterval)

      Monitor.Enter(Point) ' Lock
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
        If (StartTimestamp < PreviousStartTimestamp And
          EndTimestamp > PreviousEndTimestamp) Or
          EndTimestamp <= PreviousStartTimestamp Or
          StartTimestamp >= PreviousEndTimestamp Then ' Cases 8, 10a), 10b)
          Values.Clear ' Clear all
          PreviousStartTimestamp = StartTimestamp
          PreviousEndTimestamp = EndTimestamp
        Else If StartTimestamp >= PreviousStartTimestamp And
          EndTimestamp <= PreviousEndTimestamp Then ' Cases 1, 3, 4, 6
          ShouldPrepareData = False ' Do nothing
        Else If StartTimestamp < PreviousStartTimestamp Then ' Cases 2, 5
          If EndTimestamp < PreviousEndTimestamp Then ' Case 5
            Do While EndTimestamp < PreviousEndTimestamp ' Remove forward
              PreviousEndTimestamp = PreviousEndTimestamp.
                AddSeconds(-CalculationInterval)
              Values.Remove(PreviousEndTimestamp)
            Loop
          End If
          EndTimestamp = PreviousStartTimestamp ' Add backward
          PreviousStartTimestamp = StartTimestamp
        Else If EndTimestamp > PreviousEndTimestamp Then ' Cases 7, 9
          If StartTimestamp > PreviousStartTimestamp Then ' Case 9
            Do While PreviousStartTimestamp < StartTimestamp ' Remove backward
              Values.Remove(PreviousStartTimestamp)
              PreviousStartTimestamp = PreviousStartTimestamp.
                AddSeconds(CalculationInterval)
            Loop
          End If
          StartTimestamp = PreviousEndTimestamp ' Add forward
          PreviousEndTimestamp = EndTimestamp
        End If

        Try
          ' Retrieve all data from the data source. Will pass start and end
          ' timestamps. The driver can then prepare the dataset for which
          ' calculations are required in the next step. The (aligned) end time
          ' itself is excluded.
          If ShouldPrepareData Then PrepareData(StartTimestamp, EndTimestamp)
        Catch
        End Try

      Return New Dictionary(Of DateTime, Double)(Values) ' Return a copy.

      Finally
        Monitor.Exit(Point)
      End Try

    End Function


  End Class


End Namespace
