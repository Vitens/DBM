Option Explicit
Option Strict


' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
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


Imports System.Double
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMDataManager


    ' The DBMDataManager is responsible for retrieving input data.
    ' It stores and uses a DBMPointDriverAbstract object, which has a GetData
    ' method used for retrieving data.


    Public PointDriver As DBMPointDriverAbstract


    Public Sub New(PointDriver As DBMPointDriverAbstract)

      Me.PointDriver = PointDriver

    End Sub


    Public Function Value(Timestamp As DateTime) As Double

      Try
        Value = PointDriver.GetData(Timestamp, _
          Timestamp.AddSeconds(CalculationInterval))
      Catch
        Value = NaN ' Error getting data, return Not a Number
      End Try

      Return Value

    End Function


  End Class


End Namespace
