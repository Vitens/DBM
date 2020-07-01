Option Explicit
Option Strict


' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
' Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
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


Namespace Vitens.DynamicBandwidthMonitor


  Public MustInherit Class DBMPointDriverAbstract


    ' DBM drivers should inherit from this base class. In Sub New,
    ' MyBase.New(Point) should be called. At a minimum, the GetData function
    ' must be overridden and should return a value for the Timestamp passed.
    ' If required, PrepareData can be used to retrieve and store values in bulk
    ' from a source of data, to be used in the GetData function.


    Public Point As Object


    Public Sub New(Point As Object)

      ' When inheriting from this base class, call MyBase.New(Point) from Sub
      ' New to store the unique identifier in the Point object.

      Me.Point = Point

    End Sub


    Public Sub TryPrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' The TryPrepareData function is called from the DBM class to retrieve
      ' data using the overridden PrepareData function.

      Try
        PrepareData(StartTimestamp, EndTimestamp)
      Catch
      End Try

    End Sub


    Public Overridable Sub PrepareData(StartTimestamp As DateTime,
      EndTimestamp As DateTime)

      ' Retrieve and store values in bulk for the passed time range from a
      ' source of data, to be used in the GetData method.

    End Sub


    Public Function TryGetData(Timestamp As DateTime) As Double

      ' The TryGetData function is called from the DBMPoint class to retrieve
      ' data using the overridden GetData function. If there is an exception
      ' when calling this function, NaN is returned instead.

      Try
        TryGetData = GetData(Timestamp)
      Catch
        TryGetData = NaN ' Error getting data, return Not a Number
      End Try

      Return TryGetData

    End Function


    Public MustOverride Function GetData(Timestamp As DateTime) As Double
    ' GetData must be overridden and should return a value for the passed
    ' timestamp.


  End Class


End Namespace
