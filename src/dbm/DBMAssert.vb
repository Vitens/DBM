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
Imports System.Convert
Imports System.Double
Imports System.Math


Namespace Vitens.DynamicBandwidthMonitor


  Public Class DBMAssert


    ' In computer programming, specifically when using the imperative
    ' programming paradigm, an assertion is a predicate (a Boolean-valued
    ' function over the state space, usually expressed as a logical proposition
    ' using the variables of a program) connected to a point in the program,
    ' that always should evaluate to true at that point in code execution.
    ' Assertions can help a programmer read the code, help a compiler compile
    ' it, or help the program detect its own defects.
    ' For the latter, some programs check assertions by actually evaluating the
    ' predicate as they run. Then, if it is not in fact true - an assertion
    ' failure -, the program considers itself to be broken and typically
    ' deliberately crashes or throws an assertion failure exception.


    Public Shared Function Hash(Values() As Double) As Double

      ' Simple hash function for checking array contents.

      Dim Value As Double

      Hash = 1
      For Each Value In Values
        If Not IsNaN(Value) Then
          Hash = (Hash+Value+1)/3
        End If
      Next

      Return Hash

    End Function


    Public Shared Sub AssertEqual(a As Object, b As Object)

      If TypeOf a Is Boolean And TypeOf b Is Boolean AndAlso
        ToBoolean(a) = ToBoolean(b) Then Exit Sub
      If TypeOf a Is DateTime And TypeOf b Is DateTime AndAlso
        ToDateTime(a) = ToDateTime(b) Then Exit Sub
      If TypeOf a Is Double And TypeOf b Is Double AndAlso
        (IsNaN(ToDouble(a)) And IsNaN(ToDouble(b))) Then Exit Sub
      If (TypeOf a Is Integer Or TypeOf a Is Double Or TypeOf a Is Decimal) And
        (TypeOf b Is Integer Or TypeOf b Is Double Or
        TypeOf b Is Decimal) AndAlso
        ToDouble(a) = ToDouble(b) Then Exit Sub
      Throw New Exception("Assert failed a=" & a.ToString & " b=" & b.ToString)

    End Sub


    Public Shared Sub AssertArrayEqual(a() As Double, b() As Double)

      AssertEqual(Hash(a), Hash(b))

    End Sub


    Public Shared Sub AssertAlmostEqual(a As Object, b As Object,
      Optional Digits As Integer = 4)

      AssertEqual(Round(ToDouble(a), Digits), Round(ToDouble(b), Digits))

    End Sub


    Public Shared Sub AssertTrue(a As Boolean)

      AssertEqual(a, True)

    End Sub


    Public Shared Sub AssertFalse(a As Boolean)

      AssertEqual(a, False)

    End Sub


    Public Shared Sub AssertNaN(a As Double)

      AssertEqual(a, NaN)

    End Sub


  End Class


End Namespace
