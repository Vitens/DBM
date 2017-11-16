![Logo](img/dbmlogo.png)

# DBM
Dynamic Bandwidth Monitor  
Leak detection method implemented in a real-time data historian

Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.  
Auteursrecht (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.

## Continuous integration
| Build status                                                                                                                                                              | Test results                                                                                | Downloads                                                                 |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------- |
| [![Build status](https://ci.appveyor.com/api/projects/status/yop6ltcux5bam7ed/branch/master?svg=true)](https://ci.appveyor.com/project/JohanFiti/dbm-jbrie/branch/master) | [:thumbsup: Tests](https://ci.appveyor.com/project/JohanFiti/dbm-jbrie/branch/master/tests) | [:package: Latest release](https://github.com/Vitens/DBM/releases/latest) |

## Description
Water company Vitens has created a demonstration site called the Vitens Innovation Playground (VIP), in which new technologies and methodologies are developed, tested, and demonstrated. The projects conducted in the demonstration site can be categorized into one of four themes: energy optimization, real-time leak detection, online water quality monitoring, and customer interaction. In the real-time leak detection theme, a method for leak detection based on statistical demand forecasting was developed.

Using historical demand patterns and statistical methods - such as median absolute deviation, linear regression, sample variance, and exponential moving averages - real-time values can be compared to a predicted demand pattern and checked to be within calculated bandwidths. The method was implemented in Vitens’ realtime data historian, continuously comparing measured demand values to be within operational bounds.

One of the advantages of this method is that it doesn’t require manual configuration or training sets. Next to leak detection, unmeasured supply between areas and unscheduled plant shutdowns were also detected. The method was found to be such a success within the company, that it was implemented in an operational dashboard and is now used in day-to-day operations.

### Keywords
Real-time, leak detection, demand forecasting, demand patterns, operational dashboard

## Samples

### Sample 1 - Prediction
![Sample 1](img/sample1.png)

In this example, two days before and after the current day are shown. For historic values, the measured data (black) is shown along with the predicted value (red). The upper and lower control limits (gray) were not crossed, so the DBM factor value (blue) equals zero. For future values, the prediction is shown along with the upper and lower control limits. Reliable predictions can be made for at least seven days in advance.

### Sample 2 - Exception
![Sample 2](img/sample2.png)

In this example, an exception causes the measured value (black) to cross the upper control limit (gray). The DBM factor value (blue) is greater than one during this time (calculated as _(measured value - predicted value)/(upper control limit - predicted value)_).

### Sample 3 - Suppressed exception (correlation)
![Sample 3a](img/sample3a.png)
![Sample 3b](img/sample3b.png)

In this example, an exception causes the measured value (black) to cross the upper and lower control limits (gray). Because the pattern is checked against a similar pattern which has a comparable relative prediction error (calculated as _(predicted value / measured value) - 1_), the exception is suppressed. The DBM factor value is greater than zero and less than, or equal to one (correlation coefficient of the relative prediction error) during this time.

### Sample 4 - Suppressed exception (anticorrelation)
![Sample 4a](img/sample4a.png)
![Sample 4b](img/sample4b.png)

In this example, an exception causes the measured value (black) to cross the lower control limit (gray). Because the pattern is checked against a similar, adjacent, pattern which has a comparable, but inverted, absolute prediction error (calculated as _predicted value - measured value_), the exception is suppressed. The DBM factor value is less than zero and greater than, or equal to negative one (correlation coefficient of the absolute prediction error) during this time.

## Program information

### Requirements
| Priority  | Requirement                                                     | Version |
| --------- | --------------------------------------------------------------- | ------- |
| Mandatory | Microsoft .NET Framework                                        | 4.5     |
| Optional  | OSIsoft PI Software Development Kit (PI SDK)                    |         |
| Optional  | OSIsoft PI Asset Framework Software Development Kit (PI AF SDK) |         |
| Optional  | OSIsoft PI Advanced Computing Engine (PI ACE)                   |         |

### Drivers
DBM uses drivers to read information from a source of data. The following drivers are included:

| Driver                         | Description                             | Identifier (`Point`)             | Remarks                                                                |
| ------------------------------ | --------------------------------------- | -------------------------------- | ---------------------------------------------------------------------- |
| `DBMPointDriverCSV.vb`         | Driver for CSV files (timestamp,value). | `String` (CSV filename)          | Data interval must be the same as the `CalculationInterval` parameter. |
| `DBMPointDriverOSIsoftPI.vb`   | Driver for OSIsoft PI.                  | `PISDK.PIPoint` (PI tag)         | Used by PI ACE module `DBMRt`.                                         |
| `DBMPointDriverOSIsoftPIAF.vb` | Driver for OSIsoft PI Asset Framework.  | `OSIsoft.AF.PI.PIPoint` (PI tag) | Used by PI AF Data Reference `DBMDataRef`.                             |

### Parameters
DBM can be configured using several parameters. The values for these parameters can be changed at runtime in the `Vitens.DynamicBandwidthMonitor.DBMParameters` class.

| Parameter                    | Default value | Units         | Description                                                                                             |
| ---------------------------- | ------------- | ------------- | ------------------------------------------------------------------------------------------------------- |
| `CalculationInterval`        | 300           | seconds       | Time interval at which the calculation is run.                                                          |
| `ComparePatterns`            | 12            | weeks         | Number of weeks to look back to predict the current value and control limits.                           |
| `EMAPreviousPeriods`         | 6             | intervals     | Number of previous intervals used to smooth the data.                                                   |
| `ConfidenceInterval`         | 0.99          | ratio         | Confidence interval used for removing outliers and determining control limits.                          |
| `CorrelationPreviousPeriods` | 23            | intervals     | Number of previous intervals used to calculate prediction error correlation when an exception is found. |
| `CorrelationThreshold`       | 0.83666       | dimensionless | Absolute correlation lower limit for detecting (anti)correlation.                                       |
| `RegressionAngleRange`       | 18.435        | degrees       | Regression angle range (around -45/+45 degrees) required when suppressing based on (anti)correlation.   |

### DBMTester
DBMTester is a command line utility that can be used to quickly calculate DBM results using the CSV driver. The following arguments are available:

| Argument | Count  | Description                                                                                                            |
| -------- | ------ | ---------------------------------------------------------------------------------------------------------------------- |
| `-i=`    | 1      | Specifies the input point (CSV file).                                                                                  |
| `-c=`    | 0..n   | Adds a correlation point (CSV file).                                                                                   |
| `-cs=`   | 0..n   | Adds a correlation point (CSV file) from which the input point is subtracted.                                          |
| `-iv=`   | 0..1   | Changes the `CalculationInterval` parameter.                                                                           |
| `-p=`    | 0..1   | Changes the `ComparePatterns` parameter.                                                                               |
| `-ep=`   | 0..1   | Changes the `EMAPreviousPeriods` parameter.                                                                            |
| `-ci=`   | 0..1   | Changes the `ConfidenceInterval` parameter.                                                                            |
| `-cp=`   | 0..1   | Changes the `CorrelationPreviousPeriods` parameter.                                                                    |
| `-ct=`   | 0..1   | Changes the `CorrelationThreshold` parameter.                                                                          |
| `-ra=`   | 0..1   | Changes the `RegressionAngleRange` parameter.                                                                          |
| `-st=`   | 1      | Start timestamp for calculations.                                                                                      |
| `-et=`   | 0..1   | End timestamp for calculations, all intervals in between are calculated.                                               |
| `-f=`    | 0..1   | Output format. `local` (default) for local formatting, `intl` for UTC time and international formatting (ISO 8601).    |
| `-oc=`   | 0..1   | Output correlation data. Default (`false`) is to only output data during an exception; set to `true` to always output. |

### License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.

### Licentie

Dit programma is vrije software: je mag het herdistribueren en/of wijzigen onder de voorwaarden van de GNU Algemene Publieke Licentie zoals gepubliceerd door de Free Software Foundation, onder versie 3 van de Licentie of (naar jouw keuze) elke latere versie.

Dit programma is gedistribueerd in de hoop dat het nuttig zal zijn maar ZONDER ENIGE GARANTIE; zelfs zonder de impliciete garanties die GEBRUIKELIJK ZIJN IN DE HANDEL of voor BRUIKBAARHEID VOOR EEN SPECIFIEK DOEL.  Zie de GNU Algemene Publieke Licentie voor meer details.

Je hoort een kopie van de GNU Algemene Publieke Licentie te hebben ontvangen samen met dit programma.  Als dat niet het geval is, zie <http://www.gnu.org/licenses/>.

## About Vitens
Vitens is the largest drinking water company in The Netherlands. We deliver top quality drinking water to 5.6 million people and companies in the provinces Flevoland, Fryslân, Gelderland, Utrecht and Overijssel and some municipalities in Drenthe and Noord-Holland. Annually we deliver 350 million m³ water with 1,400 employees, 100 water treatment works and 49,000 kilometres of water mains.

One of our main focus points is using advanced water quality, quantity and hydraulics models to further improve and optimize our treatment and distribution processes.

![Vitens](img/vitens.png)

https://www.vitens.nl/
