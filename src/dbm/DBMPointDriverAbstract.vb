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
Imports System.Double
Imports System.Threading
Imports Vitens.DynamicBandwidthMonitor.DBMDate
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMPointDriverAbstract


    ' DBM drivers should inherit from this base class. In Sub New,
    ' MyBase.New(Point) should be called. At a minimum, the GetData function
    ' must be overridden and should return a value for the Timestamp passed.
    ' PrepareData is used to retrieve and store values in bulk from a source of
    ' data, to be used in the GetData function.


    Public Point As Object


    Public Sub New(Point As Object)

      ' When inheriting from this base class, call MyBase.New(Point) from Sub
      ' New to store the unique identifier in the Point object.

      Me.Point = Point

    End Sub


    Public Overridable Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' Retrieve and store values in bulk for the passed time range from a
      ' source of data, to be used in the GetData method.

    End Sub


    Public Function PrepareAndGetData(StartTimestamp As DateTime,
      EndTimestamp As DateTime) As Double

      ' The PrepareAndGetData function is called from the DBMPoint class to
      ' retrieve data using the overridden GetData function. If there is an
      ' exception when calling this function, NaN is returned instead. Data will
      ' be prepared for the time range and a value will be returned for the
      ' start timestamp.

      StartTimestamp = NextInterval(StartTimestamp,
        -EMAPreviousPeriods-CorrelationPreviousPeriods).
        AddDays(ComparePatterns*-7)
      If UseSundayForHolidays Then StartTimestamp =
        PreviousSunday(StartTimestamp)
      EndTimestamp = AlignTimestamp(EndTimestamp, CalculationInterval)

      Monitor.Enter(Point) ' Lock
      Try

        Try
          ' Retrieve all data from the data source. Will pass start and end
          ' timestamps. The driver can then prepare the dataset for which
          ' calculations are required in the next step. The (aligned) end time
          ' itself is excluded.
          PrepareData(StartTimestamp, EndTimestamp)
        Catch
        End Try

        Try
          PrepareAndGetData = GetData(StartTimestamp)
        Catch
          PrepareAndGetData = NaN ' Error getting data, return Not a Number
        End Try

        Return PrepareAndGetData

      Finally
        Monitor.Exit(Point)
      End Try

    End Function


    Public MustOverride Function GetData(Timestamp As DateTime) As Double
    ' GetData must be overridden and should return a value for the passed
    ' timestamp.


  End Class


End Namespace
