# DBM
Dynamic Bandwidth Monitor  
Leak detection method implemented in a real-time data historian  
J.H. Fitié, Vitens N.V. <[johan.fitie@vitens.nl](mailto:johan.fitie@vitens.nl)>  
E.A. Trietsch, Vitens N.V. <[eelco.trietsch@vitens.nl](mailto:eelco.trietsch@vitens.nl)>  
[www.github.com/jfitie/DBM](https://www.github.com/jfitie/DBM)

## Samples
### Sample 1 - Prediction
![Sample 1](https://raw.githubusercontent.com/jfitie/DBM/master/docs/sample1.png)
[Sample 1 data](https://github.com/jfitie/DBM/blob/master/docs/sample1.csv)
```
Option Explicit
Option Strict
Module Module1
	Public Sub Main
		Dim _DBM As New DBM
		Dim _PISDK As PISDK.PISDK=New PISDK.PISDK
		Dim s,e As DateTime
		Dim r As New DBMResult
		s=DateAdd("d",-2,DateTime.Today)
		e=DateAdd("d",5,s)
		Do While s<e
			r=_DBM.Calculate(New DBMPointDriver(_PISDK.Servers("sr-16635").PIPoints("ACE-FR-Deelbalansgebied-Leeuwarden-levering")),Nothing,s,False)
			Console.WriteLine(s.ToString & vbTab & r.Factor & vbTab & r.CurrValue & vbTab & r.PredValue & vbTab & r.LowContrLimit & vbTab & r.UppContrLimit)
			s=DateAdd("s",DBMConstants.CalculationInterval,s)
		Loop
	End Sub
End Module
```
### Sample 2 - Exception
![Sample 2](https://raw.githubusercontent.com/jfitie/DBM/master/docs/sample2.png)
[Sample 2 data](https://github.com/jfitie/DBM/blob/master/docs/sample2.csv)
```
Option Explicit
Option Strict
Module Module1
	Public Sub Main
		Dim _DBM As New DBM
		Dim _PISDK As PISDK.PISDK=New PISDK.PISDK
		Dim s,e As DateTime
		Dim r As New DBMResult
		s=New DateTime(2013,3,12)
		e=DateAdd("d",1,s)
		Do While s<e
			r=_DBM.Calculate(New DBMPointDriver(_PISDK.Servers("sr-16635").PIPoints("ACE-FR-Deelbalansgebied-Leeuwarden-levering")),Nothing,s,False)
			Console.WriteLine(s.ToString & vbTab & r.Factor & vbTab & r.CurrValue & vbTab & r.PredValue & vbTab & r.LowContrLimit & vbTab & r.UppContrLimit)
			s=DateAdd("s",DBMConstants.CalculationInterval,s)
		Loop
	End Sub
End Module
```
### Sample 3 - Suppressed exception
![Sample 3a](https://raw.githubusercontent.com/jfitie/DBM/master/docs/sample3a.png)
![Sample 3b](https://raw.githubusercontent.com/jfitie/DBM/master/docs/sample3b.png)
[Sample 3 data](https://github.com/jfitie/DBM/blob/master/docs/sample3.csv)
```
Option Explicit
Option Strict
Module Module1
	Public Sub Main
		Dim _DBM As New DBM
		Dim _PISDK As PISDK.PISDK=New PISDK.PISDK
		Dim s,e As DateTime
		Dim r As New DBMResult
		s=New DateTime(2016,1,1)
		e=DateAdd("d",1,s)
		Do While s<e
			r=_DBM.Calculate(New DBMPointDriver(_PISDK.Servers("sr-16635").PIPoints("ACE-FR-Deelbalansgebied-Leeuwarden-levering")),New DBMPointDriver(_PISDK.Servers("sr-16634").PIPoints("Reinwaterafgifte")),s,False)
			Console.Write(s.ToString & vbTab & r.Factor & vbTab & r.CurrValue & vbTab & r.PredValue & vbTab & r.LowContrLimit & vbTab & r.UppContrLimit)
			r=_DBM.Calculate(New DBMPointDriver(_PISDK.Servers("sr-16634").PIPoints("Reinwaterafgifte")),Nothing,s,False)
			Console.WriteLine(vbTab & r.CurrValue & vbTab & r.PredValue & vbTab & r.LowContrLimit & vbTab & r.UppContrLimit)
			s=DateAdd("s",DBMConstants.CalculationInterval,s)
		Loop
	End Sub
End Module
```

## About Vitens
Vitens is the largest drinking water company in The Netherlands. We deliver top quality drinking water to 5.6 million people and companies in the provinces Flevoland, Fryslân, Gelderland, Utrecht and Overijssel and some municipalities in Drenthe and Noord-Holland. Annually we deliver 350 million m³ water with 1,400 employees, 100 water treatment works and 49,000 kilometres of water mains.  
One of our main focus points is using advanced water quality, quantity and hydraulics models to further improve and optimize our treatment and distribution processes.  
[www.vitens.nl](https://www.vitens.nl/)
