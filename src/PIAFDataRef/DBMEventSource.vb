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
Imports System.DateTime
Imports OSIsoft.AF.Asset
Imports OSIsoft.AF.Data
Imports OSIsoft.AF.Time
Imports Vitens.DynamicBandwidthMonitor.DBMMath
Imports Vitens.DynamicBandwidthMonitor.DBMParameters


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMEventSource
    Inherits AFEventSource


    Private LastTime As AFTime = Now


    Protected Overrides Function GetEvents As Boolean

      Dim EvalTime As AFTime = Now
      Dim Attribute As AFAttribute
      Dim Value As AFValue

      EvalTime = New AFTime(AlignPreviousInterval(EvalTime.UtcSeconds,
        CalculationInterval))

      For Each Attribute In MyBase.Signups

        If Attribute IsNot Nothing And LastTime < EvalTime Then

          For Each Value In Attribute.GetValues(
            New AFTimeRange(LastTime, EvalTime), 0, Nothing)

            MyBase.PublishEvent(Attribute,
              New AFDataPipeEvent(AFDataPipeAction.Add, Value))

          Next

        End If

      Next

      LastTime = EvalTime

      Return False

    End Function


    Protected Overrides Sub Dispose(disposing As Boolean)

      LastTime = Nothing

    End Sub


  End Class


End Namespace
