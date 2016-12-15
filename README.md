![Logo](docs/dbmlogo.png)

# DBM
Dynamic Bandwidth Monitor  
Leak detection method implemented in a real-time data historian

## Authors
| Name          | Company      | Role        | Contact                  |
| ------------- | ------------ | ----------- | ------------------------ |
| J.H. Fitié    | Vitens N.V.  | Lead author | johan.fitie@vitens.nl    |
| E.A. Trietsch | Vitens N.V.  | Co-author   | eelco.trietsch@vitens.nl |
|               | VORtech B.V. | Contributor |                          |

## Description
Water company Vitens has created a demonstration site called the Vitens Innovation Playground (VIP), in which new technologies and methodologies are developed, tested, and demonstrated. The projects conducted in the demonstration site can be categorized into one of four themes: energy optimization, real-time leak detection, online water quality monitoring, and customer interaction. In the real-time leak detection theme, a method for leak detection based on statistical demand forecasting was developed.

Using historical demand patterns and statistical methods - such as median absolute deviation, linear regression, sample variance, and exponential moving averages - real-time values can be compared to a predicted demand pattern and checked to be within calculated bandwidths. The method was implemented in Vitens’ realtime data historian, continuously comparing measured demand values to be within operational bounds.

One of the advantages of this method is that it doesn’t require manual configuration or training sets. Next to leak detection, unmeasured supply between areas and unscheduled plant shutdowns were also detected. The method was found to be such a success within the company, that it was implemented in an operational dashboard and is now used in day-to-day operations.

### Keywords
Real-time, leak detection, demand forecasting, demand patterns, operational dashboard

## Samples

### Sample 1 - Prediction
![Sample 1](docs/sample1.png)

In this example, two days before and after the current day are shown. For historic values, the measured data (black) is shown along with the predicted value (red). The upper and lower control limits (gray) were not crossed, so the DBM factor value (blue) equals zero. For future values, the prediction is shown along with the upper and lower control limits. Reliable predictions can be made for at least seven days in advance.

### Sample 2 - Exception
![Sample 2](docs/sample2.png)

In this example, an exception causes the measured value (black) to cross the upper control limit (gray). The DBM factor value (blue) is greater than one during this time (calculated as _(measured value - predicted value)/(upper control limit - predicted value)_).

### Sample 3 - Suppressed exception (correlation)
![Sample 3a](docs/sample3a.png)
![Sample 3b](docs/sample3b.png)

In this example, an exception causes the measured value (black) to cross the upper and lower control limits (gray). Because the pattern is checked against a similar pattern which has a comparable relative prediction error (calculated as _(predicted value / measured value) - 1_), the exception is suppressed. The DBM factor value is greater than zero and less than, or equal to one (correlation coefficient of the relative prediction error) during this time.

### Sample 4 - Suppressed exception (anticorrelation)
![Sample 4a](docs/sample4a.png)
![Sample 4b](docs/sample4b.png)

In this example, an exception causes the measured value (black) to cross the lower control limit (gray). Because the pattern is checked against a similar, adjacent, pattern which has a comparable, but inverted, absolute prediction error (calculated as _predicted value - measured value_), the exception is suppressed. The DBM factor value is less than zero and greater than, or equal to negative one (correlation coefficient of the absolute prediction error) during this time.

## Program information

### Requirements
| Priority  | Requirement                                   | Version    |
| --------- | --------------------------------------------- | ---------- |
| Mandatory | Microsoft .NET Framework                      | v4.0.30319 |
| Optional  | OSIsoft PI Software Development Kit (PI SDK)  |            |
| Optional  | OSIsoft PI Advanced Computing Engine (PI ACE) |            |

### Drivers
DBM uses drivers to read information from a source of data. The following drivers are included:

| Driver                  | Description                             | Identifier (`Point`)     | Remarks                                                                |
| ----------------------- | --------------------------------------- | ------------------------ | ---------------------------------------------------------------------- |
| `DBMDriverNull.vb`      | Returns zero for every request.         | `Object` (unused)        | Template for custom developed drivers.                                 |
| `DBMDriverCSV.vb`       | Driver for CSV files (timestamp,value). | `String` (CSV filename)  | Data interval must be the same as the `CalculationInterval` parameter. |
| `DBMDriverOSIsoftPI.vb` | Driver for OSIsoft PI.                  | `PISDK.PIPoint` (PI tag) | Used by PI ACE module `DBMRt`.                                         |

### Parameters
DBM can be configured using several parameters. The values for these parameters can be changed at runtime in the `DBM.DBMParameters` class.

| Parameter                    | Default value | Units         | Description                                                                                             |
| ---------------------------- | ------------- | ------------- | ------------------------------------------------------------------------------------------------------- |
| `CalculationInterval`        | 300           | seconds       | Time interval at which the calculation is run.                                                          |
| `ComparePatterns`            | 12            | weeks         | Number of weeks to look back to predict the current value and control limits.                           |
| `EMAPreviousPeriods`         | 6             | intervals     | Number of previous intervals used to smooth the data.                                                   |
| `ConfidenceInterval`         | 0.99          | percent       | Confidence interval used for removing outliers and determining control limits.                          |
| `CorrelationPreviousPeriods` | 23            | intervals     | Number of previous intervals used to calculate prediction error correlation when an exception is found. |
| `CorrelationThreshold`       | 0.83666       | dimensionless | Absolute correlation lower limit for detecting (anti)correlation.                                       |

### DBMTester
DBMTester is a command line utility that can be used to quickly calculate DBM results using the CSV driver. The following arguments are available:

| Argument | Count  | Description                                                                                                             |
| -------- | ------ | ----------------------------------------------------------------------------------------------------------------------- |
| `-i=`    | 1      | Specifies the input point (CSV file).                                                                                   |
| `-c=`    | 0..n   | Adds a correlation point (CSV file).                                                                                    |
| `-cs=`   | 0..n   | Adds a correlation point (CSV file) from which the input point is substracted.                                          |
| `-iv=`   | 0..1   | Changes the `CalculationInterval` parameter.                                                                            |
| `-p=`    | 0..1   | Changes the `ComparePatterns` parameter.                                                                                |
| `-ep=`   | 0..1   | Changes the `EMAPreviousPeriods` parameter.                                                                             |
| `-ci=`   | 0..1   | Changes the `ConfidenceInterval` parameter.                                                                             |
| `-cp=`   | 0..1   | Changes the `CorrelationPreviousPeriods` parameter.                                                                     |
| `-ct=`   | 0..1   | Changes the `CorrelationThreshold` parameter.                                                                           |
| `-st=`   | 1      | Start timestamp for calculations.                                                                                       |
| `-et=`   | 0..1   | End timestamp for calculations, all intervals in between are calculated.                                                |
| `-f=`    | 0..1   | CSV output format. `local` (default) for local formatting, `intl` for UTC time and international formatting (ISO 8601). |

### License
DBM  
Dynamic Bandwidth Monitor  
Leak detection method implemented in a real-time data historian

Copyright (C) 2014, 2015, 2016 J.H. Fitié, Vitens N.V.

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.

![GPLv3](docs/gplv3.png)

## About Vitens
Vitens is the largest drinking water company in The Netherlands. We deliver top quality drinking water to 5.6 million people and companies in the provinces Flevoland, Fryslân, Gelderland, Utrecht and Overijssel and some municipalities in Drenthe and Noord-Holland. Annually we deliver 350 million m³ water with 1,400 employees, 100 water treatment works and 49,000 kilometres of water mains.

One of our main focus points is using advanced water quality, quantity and hydraulics models to further improve and optimize our treatment and distribution processes.

![Vitens](docs/vitens.png)

https://www.vitens.nl/
