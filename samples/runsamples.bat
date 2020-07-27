@echo off

rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
rem
rem This file is part of DBM.
rem
rem DBM is free software: you can redistribute it and/or modify
rem it under the terms of the GNU General Public License as published by
rem the Free Software Foundation, either version 3 of the License, or
rem (at your option) any later version.
rem
rem DBM is distributed in the hope that it will be useful,
rem but WITHOUT ANY WARRANTY; without even the implied warranty of
rem MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
rem GNU General Public License for more details.
rem
rem You should have received a copy of the GNU General Public License
rem along with DBM.  If not, see <http://www.gnu.org/licenses/>.

cd /d %~dp0

..\build\DBMTester.exe -f=intl -i=sample1input.csv -st=11-24-2016 > sample1.csv
..\build\DBMTester.exe -f=intl -i=sample2input.csv -st=03-12-2013 -et=03-13-2013 > sample2.csv
..\build\DBMTester.exe -f=intl -us=False -i=sample3input.csv -c=sample3correlation.csv -st=01-01-2016 -et=01-02-2016 > sample3.csv
..\build\DBMTester.exe -f=intl -i=sample4input.csv -c=sample4correlation.csv -st=11-13-2014 -et=11-14-2014 > sample4.csv
