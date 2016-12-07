Option Explicit
Option Strict

' DBM
' Dynamic Bandwidth Monitor
' Leak detection method implemented in a real-time data historian
'
' Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.
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

<assembly:System.Reflection.AssemblyTitle("DBMUnitTests")>

Module DBMUnitTests

    Dim _DBM As DBM.DBM

    Private Sub UnitTest(ExpFact As Double,ExpOrigFact As Double,ExpCurr As Double,ExpPred As Double,ExpLCL As Double,ExpUCL As Double,InputPoint As Object,CorrelationPoint As Object,Timestamp As DateTime,SubstractSelf As Boolean)
        Dim InputDBMPointDriver,CorrelationDBMPointDriver As DBM.DBMPointDriver
        Dim DBMCorrelationPoints As New Collections.Generic.List(Of DBM.DBMCorrelationPoint)
        Dim DBMResult As DBM.DBMResult
        InputDBMPointDriver=New DBM.DBMPointDriver(CStr(InputPoint))
        If CorrelationPoint IsNot Nothing Then
            CorrelationDBMPointDriver=New DBM.DBMPointDriver(CStr(CorrelationPoint))
            DBMCorrelationPoints.Add(New DBM.DBMCorrelationPoint(CorrelationDBMPointDriver,SubstractSelf))
        Else
            CorrelationDBMPointDriver=Nothing
        End If
        DBMResult=_DBM.Calculate(InputDBMPointDriver,DBMCorrelationPoints,Timestamp)
        If Not (Math.Round(DBMResult.Factor,3)=ExpFact And Math.Round(DBMResult.OriginalFactor,3)=ExpOrigFact And Math.Round(DBMResult.CurrValue)=ExpCurr And Math.Round(DBMResult.PredValue)=ExpPred And Math.Round(DBMResult.LowContrLimit)=ExpLCL And Math.Round(DBMResult.UppContrLimit)=ExpUCL And (DBMResult.SuppressedBy Is Nothing Or (DBMResult.SuppressedBy IsNot Nothing And CorrelationDBMPointDriver Is DBMResult.SuppressedBy))) Then
            Throw New System.Exception("Unit tests failed.")
        End If
    End Sub

    Public Sub Main
        Console.WriteLine(DBM.DBMFunctions.DBMVersion & vbCrLf)
        _DBM=New DBM.DBM
        For i=0 To 1 ' Run all cases twice to test cache
            UnitTest(0,0,1154,1192,1023,1361,"A",Nothing,New DateTime(2016,2,5,10,0,0),False)
            UnitTest(0,0,1154,1192,1023,1361,"A",Nothing,New DateTime(2016,2,5,10,0,0),False)
            UnitTest(0,0,1141,1194,1021,1367,"A",Nothing,New DateTime(2016,2,5,10,5,0),False)
            UnitTest(0,0,27087,27408,25605,29212,"B",Nothing,New DateTime(2015,5,23,2,0,0),False)
            UnitTest(0,0,27087,27408,25605,29212,"B",Nothing,New DateTime(2015,5,23,2,0,0),False)
            UnitTest(1.206,1.206,570,378,219,537,"A",Nothing,New DateTime(2016,1,1,1,50,0),False)
            UnitTest(0.849,1.206,570,378,219,537,"A","C",New DateTime(2016,1,1,1,50,0),False)
            UnitTest(1,1.206,570,378,219,537,"A","A",New DateTime(2016,1,1,1,50,0),False)
            UnitTest(1.206,1.206,570,378,219,537,"A","A",New DateTime(2016,1,1,1,50,0),True)
            UnitTest(1.511,1.511,846,281,-93,655,"D",Nothing,New DateTime(2015,8,21,1,0,0),False)
            UnitTest(-0.996,1.511,846,281,-93,655,"D","E",New DateTime(2015,8,21,1,0,0),False)
            UnitTest(6.793,6.793,3075,1140,855,1425,"A",Nothing,New DateTime(2013,3,12,19,30,0),False)
            UnitTest(5.818,5.818,3042,1148,822,1473,"A",Nothing,New DateTime(2013,3,12,19,35,0),False)
            UnitTest(6.793,6.793,3075,1140,855,1425,"A","B",New DateTime(2013,3,12,19,30,0),True)
            UnitTest(0.846,6.717,2936,953,658,1248,"A","B",New DateTime(2013,3,12,20,10,0),False)
            UnitTest(6.717,6.717,2936,953,658,1248,"A","B",New DateTime(2013,3,12,20,10,0),True)
            UnitTest(3.206,3.206,1320,390,100,680,"F",Nothing,New DateTime(2016,1,4,4,0,0),False)
            UnitTest(3.206,3.206,1320,390,100,680,"F","A",New DateTime(2016,1,4,4,0,0),False)
            UnitTest(15.629,15.629,3180,984,844,1125,"G",Nothing,New DateTime(2015,11,25,10,0,0),False)
            UnitTest(15.629,15.629,3180,984,844,1125,"G","H",New DateTime(2015,11,25,10,0,0),False)
            UnitTest(5.831,5.831,597,355,313,396,"I",Nothing,New DateTime(2015,4,9,13,30,0),False)
            UnitTest(5.831,5.831,597,355,313,396,"I","J",New DateTime(2015,4,9,13,30,0),False)
            UnitTest(1.399,1.399,1095,627,293,961,"A",Nothing,New DateTime(2015,8,21,7,30,0),False)
            UnitTest(0.866,1.399,1095,627,293,961,"A","C",New DateTime(2015,8,21,7,30,0),False)
            UnitTest(-1.041,-1.041,1000,1189,1008,1371,"A",Nothing,New DateTime(2015,12,25,9,0,0),False)
            UnitTest(0.991,-1.041,1000,1189,1008,1371,"A","B",New DateTime(2015,12,25,9,0,0),True)
            UnitTest(2.346,2.346,816,222,-31,476,"D",Nothing,New DateTime(2015,8,21,2,0,0),False)
            UnitTest(2.369,2.369,814,216,-36,469,"D",Nothing,New DateTime(2015,8,21,2,5,0),False)
            UnitTest(1.928,1.928,801,235,-58,529,"D",Nothing,New DateTime(2015,8,21,2,10,0),False)
            UnitTest(1.676,1.676,791,279,-27,584,"D",Nothing,New DateTime(2015,8,21,2,15,0),False)
            UnitTest(-0.998,1.993,807,225,-68,517,"D","E",New DateTime(2015,8,21,3,0,0),False)
            UnitTest(-0.998,1.926,799,221,-79,521,"D","E",New DateTime(2015,8,21,3,5,0),False)
            UnitTest(-0.998,2.17,799,222,-45,488,"D","E",New DateTime(2015,8,21,3,10,0),False)
            UnitTest(-0.999,2.541,804,226,-2,453,"D","E",New DateTime(2015,8,21,3,15,0),False)
            UnitTest(1.049,1.049,60,19,-20,58,"K",Nothing,New DateTime(2014,12,31,9,15,0),False)
            UnitTest(1.082,1.082,57,14,-26,54,"K",Nothing,New DateTime(2014,12,31,9,20,0),False)
            UnitTest(1.112,1.112,164,134,106,161,"L",Nothing,New DateTime(2016,4,3,1,30,0),False)
            UnitTest(1.705,1.705,164,136,119,153,"L",Nothing,New DateTime(2016,4,3,2,0,0),False)
            UnitTest(3.201,3.201,164,133,123,143,"L",Nothing,New DateTime(2016,4,3,2,30,0),False)
            UnitTest(2.501,2.501,164,132,119,145,"L",Nothing,New DateTime(2016,4,3,3,0,0),False)
            UnitTest(3.332,3.332,164,132,122,142,"L",Nothing,New DateTime(2016,4,3,3,30,0),False)
            UnitTest(0,0,537,595,520,671,"M",Nothing,New DateTime(2016,4,28,17,50,0),False)
            UnitTest(-1.024,-1.024,535,605,537,673,"M",Nothing,New DateTime(2016,4,28,17,55,0),False)
            UnitTest(0,0,548,609,538,679,"M",Nothing,New DateTime(2016,4,28,18,0,0),False)
        Next i
        Console.WriteLine("Unit tests passed.")
    End Sub

End Module
