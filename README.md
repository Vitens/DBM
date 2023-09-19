![Logo](img/dbmlogo.png)

# DBM
Dynamic Bandwidth Monitor\
Leak detection method implemented in a real-time data historian\
Copyright (C) 2014-2023  J.H. Fitié, Vitens N.V.

## Description
Water company Vitens has created a demonstration site called the Vitens Innovation Playground (VIP), in which new technologies and methodologies are developed, tested, and demonstrated. The projects conducted in the demonstration site can be categorized into one of four themes: energy optimization, real-time leak detection, online water quality monitoring, and customer interaction. In the real-time leak detection theme, a method for leak detection based on statistical demand forecasting was developed.

Using historical demand patterns and statistical methods - such as median absolute deviation, linear regression, sample variance, and exponential moving averages - real-time values can be compared to a forecast demand pattern and checked to be within calculated bandwidths. The method was implemented in Vitens' realtime data historian, continuously comparing measured demand values to be within operational bounds.

One of the advantages of this method is that it doesn't require manual configuration or training sets. Next to leak detection, unmeasured supply between areas and unscheduled plant shutdowns were also detected. The method was found to be such a success within the company, that it was implemented in an operational dashboard and is now used in day-to-day operations.

### Keywords
Real-time, leak detection, demand forecasting, demand patterns, operational dashboard

## Samples

### Sample 1 - Forecast
![Sample 1](img/sample1.png)

In this example, two days before and after the current day are shown. For historic values, the measured data (black) is shown along with the forecast value (red). The upper and lower control limits (gray) were not crossed, so the DBM factor value (blue) equals zero. For future values, the forecast is shown along with the upper and lower control limits. Reliable forecasts can be made for at least seven days in advance.

### Sample 2 - Event
![Sample 2](img/sample2.png)

In this example, an event causes the measured value (black) to cross the upper control limit (gray). The DBM factor value (blue) is greater than one during this time (calculated as _(measured value - forecast value)/(upper control limit - forecast value)_).

### Sample 3 - Suppressed event (correlation)
![Sample 3a](img/sample3a.png)
![Sample 3b](img/sample3b.png)

In this example, an event causes the measured value (black) to cross the upper and lower control limits (gray). Because the pattern is checked against a similar pattern which has a comparable relative forecast error (calculated as _(forecast value / measured value) - 1_), the event is suppressed. The DBM factor value is reset to zero during this time.

### Sample 4 - Suppressed event (anticorrelation)
![Sample 4a](img/sample4a.png)
![Sample 4b](img/sample4b.png)

In this example, an event causes the measured value (black) to cross the lower control limit (gray). Because the pattern is checked against a similar, adjacent, pattern which has a comparable, but inverted, absolute forecast error (calculated as _forecast value - measured value_), the event is suppressed. The DBM factor value is reset to zero during this time.

## Program information

### Requirements
| Priority  | Requirement                                                   | Version    |
| --------- | ------------------------------------------------------------- | ---------- |
| Mandatory | Microsoft .NET Framework                                      | 4.0.30319  |
| Optional  | AVEVA PI Asset Framework Software Development Kit (PI AF SDK) | 2.10.7.283 |

### Drivers
DBM uses drivers to read information from a source of data. The following drivers are included:

| Driver                       | Description                             | Identifier (`Point`)                             | Remarks                                                                |
| ---------------------------- | --------------------------------------- | ------------------------------------------------ | ---------------------------------------------------------------------- |
| `DBMPointDriverCSV.vb`       | Driver for CSV files (timestamp,value). | `String` (CSV filename)                          | Data interval must be the same as the `CalculationInterval` parameter. |
| `DBMPointDriverAVEVAPIAF.vb` | Driver for AVEVA PI Asset Framework.    | `OSIsoft.AF.Asset.AFAttribute` (PI AF attribute) | Used by PI AF Data Reference `DBMDataRef`.                             |

### Parameters
DBM can be configured using several parameters. The values for these parameters can be changed at runtime in the `Vitens.DynamicBandwidthMonitor.DBMParameters` class.

| Parameter                    | Default value | Units         | Description                                                                                           |
| ---------------------------- | ------------- | ------------- | ----------------------------------------------------------------------------------------------------- |
| `CalculationInterval`        | 300           | seconds       | Time interval at which the calculation is run.                                                        |
| `UseSundayForHolidays`       | True          |               | Use forecast of the previous Sunday for holidays.                                                     |
| `ComparePatterns`            | 12            | weeks         | Number of weeks to look back to forecast the current value and control limits.                        |
| `EMAPreviousPeriods`         | 5             | intervals     | Number of previous intervals used to smooth the data.                                                 |
| `OutlierCI`                  | 0.99          | ratio         | Confidence interval used for removing outliers.                                                       |
| `BandwidthCI`                | 0.99          | ratio         | Confidence interval used for determining control limits.                                              |
| `CorrelationPreviousPeriods` | 23            | intervals     | Number of previous intervals used to calculate forecast error correlation when an event is found.     |
| `CorrelationThreshold`       | 0.83666       | dimensionless | Absolute correlation lower limit for detecting (anti)correlation.                                     |
| `RegressionAngleRange`       | 18.435        | degrees       | Regression angle range (around -45/+45 degrees) required when suppressing based on (anti)correlation. |

### DBMTester
DBMTester is a command line utility that can be used to quickly calculate DBM results using the CSV driver. The following arguments are available:

| Argument | Count  | Description                                                                       |
| -------- | ------ | --------------------------------------------------------------------------------- |
| `-i=`    | 1      | Specifies the input point (CSV file).                                             |
| `-c=`    | 0..n   | Adds a correlation point (CSV file).                                              |
| `-cs=`   | 0..n   | Adds a correlation point (CSV file) from which the input point is subtracted.     |
| `-iv=`   | 0..1   | Changes the `CalculationInterval` parameter.                                      |
| `-us=`   | 0..1   | Changes the `UseSundayForHolidays` parameter.                                     |
| `-p=`    | 0..1   | Changes the `ComparePatterns` parameter.                                          |
| `-ep=`   | 0..1   | Changes the `EMAPreviousPeriods` parameter.                                       |
| `-oi=`   | 0..1   | Changes the `OutlierCI` parameter.                                                |
| `-bi=`   | 0..1   | Changes the `BandwidthCI` parameter.                                              |
| `-cp=`   | 0..1   | Changes the `CorrelationPreviousPeriods` parameter.                               |
| `-ct=`   | 0..1   | Changes the `CorrelationThreshold` parameter.                                     |
| `-ra=`   | 0..1   | Changes the `RegressionAngleRange` parameter.                                     |
| `-st=`   | 1      | Start timestamp for calculations.                                                 |
| `-et=`   | 0..1   | End timestamp for calculations, all intervals in between are calculated.          |
| `-f=`    | 0..1   | Output format. Default for local formatting, `intl` for international formatting. |

### DBMDataRef
DBMDataRef is a custom AVEVA PI Asset Framework data reference which integrates DBM with PI AF. The `register.bat` script automatically registers the data reference and support assemblies when run on the PI AF server. The data reference uses the parent attribute as input and automatically uses attributes from sibling and parent elements based on the same template containing good data for correlation calculations, unless the `NoCorrelation` category is applied to the output attribute. The value returned from the DBM calculation is determined by the applied property/trait:

| Property/trait | Return value                              | Description                                                                                                   |
| -------------- | ----------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| None           | Factor                                    |                                                                                                               |
| `Target`       | Measured value, forecast if not available | Indicates the aimed-for measurement value or process output.                                                  |
| `Forecast`     | Forecast value                            |                                                                                                               |
| `Minimum`      | Lower control limit (p = 0.9999)          | Indicates the very lowest possible measurement value or process output.                                       |
| `LoLo`         | Lower control limit (default)             | Indicates a very low measurement value or process output, typically an abnormal one that initiates an alarm.  |
| `Lo`           | Lower control limit (p = 0.95)            | Indicates a low measurement value or process output, typically one that initiates a warning.                  |
| `Hi`           | Upper control limit (p = 0.95)            | Indicates a high measurement value or process output, typically one that initiates a warning.                 |
| `HiHi`         | Upper control limit (default)             | Indicates a very high measurement value or process output, typically an abnormal one that initiates an alarm. |
| `Maximum`      | Upper control limit (p = 0.9999)          | Indicates the very highest possible measurement value or process output.                                      |

The `Target` trait returns the original, valid raw measurements, augmented with forecast data for invalid and future data. Data quality is marked as `Substituted` for invalid data replaced by forecast values and for future forecast values, or as `Questionable` for data exceeding `Minimum` and `Maximum` control limits.
![Augmented](img/augmented.png)

For the `Target` trait, DBM also detects and removes incorrect flatlines in the data and replaces them with forecast values, which are then weight adjusted to the original time range total. Data quality is marked as `Substituted` for this data.

Beginning with PI AF 2018 SP3 Patch 2, all AF plugins must be signed with a valid certificate. Users must ensure any 3rd party or custom plugins are signed with a valid certificate. Digitally signing plugins increases the users' confidence that it is from a trusted entity. AF data references that are not signed could have been tampered with and are potentially dangerous. Therefore, in order to protect its users, AVEVA software enforces that all AF 2.x data reference plugins be signed. Use `sign.bat` to sign the data reference and support assemblies with your pfx certificate file before registering the data reference with PI AF.

### Model calibration metrics: systematic error, random error, and fit
DBM can calculate model calibration metrics. This information is exposed in the DBMTester utility and the DBMDataRef data reference. The model is considered to be calibrated if all of the following conditions are met:

* the absolute normalized mean bias error (NMBE, as a measure of bias) is 10% or lower,
* the absolute coefficient of variation of the root-mean-square deviation (CV(RMSD), as a measure of error) is 30% or lower,
* the determination (R², as a measure of fit) is 0.75 or higher.

The normalized mean bias error is used as a measure of the systematic error. For the random error, the difference between the absolute normalized mean bias error and the absolute coefficient of variation of the root-mean-square deviation is used.

There are several agencies that have developed guidelines and methodologies to establish a measure of the accuracy of models. We decided to follow the guidelines as documented in _ASHRAE Guideline 14-2014, Measurement of Energy, Demand, and Water Savings_, by the American Society of Heating, Refrigerating and Air Conditioning Engineers, and _International Performance Measurement and Verification Protocol: Concepts and Options for Determining Energy and Water Savings, Volume I_, by the Efficiency Valuation Organization.

An example of the model calibration metrics:\
`Calibrated: True (n 2016; Systematic error 2.2%; Random error 4.2%; Fit 97.9%)`

### License
This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>.

## About Vitens
Vitens is the largest drinking water company in The Netherlands. We deliver top quality drinking water to 5.7 million people and companies in the provinces Flevoland, Fryslan, Gelderland, Utrecht and Overijssel and some municipalities in Drenthe and Noord-Holland. Annually we deliver 350 million m3 water with 1,400 employees, 100 water treatment works and 50,000 kilometres of water mains.

One of our main focus points is using advanced water quality, quantity and hydraulics models to further improve and optimize our treatment and distribution processes.

![Vitens](img/vitens.png)

https://www.vitens.nl/

## About AVEVA
AVEVA, the AVEVA logo and logotype, OSIsoft, the OSIsoft logo and logotype, Managed PI, OSIsoft Advanced Services, OSIsoft Cloud Services, OSIsoft Connected Services, OSIsoft EDS, PI ACE, PI Advanced Computing Engine, PI AF SDK, PI API, PI Asset Framework, PI Audit Viewer, PI Builder, PI Cloud Connect, PI Connectors, PI Data Archive, PI DataLink, PI DataLink Server, PI Developers Club, PI Integrator for Business Analytics, PI Interfaces, PI JDBC Driver, PI Manual Logger, PI Notifications, PI ODBC Driver, PI OLEDB Enterprise, PI OLEDB Provider, PI OPC DA Server, PI OPC HDA Server, PI ProcessBook, PI SDK, PI Server, PI Square, PI System, PI System Access, PI Vision, PI Visualization Suite, PI Web API, PI WebParts, PI Web Services, RLINK, and RtReports are all trademarks of AVEVA Group plc or its subsidiaries. All other trademarks or trade names used herein are the property of their respective owners.
