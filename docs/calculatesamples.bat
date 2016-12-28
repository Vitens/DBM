@echo off

rem DBM
rem Dynamic Bandwidth Monitor
rem Leak detection method implemented in a real-time data historian
rem
rem Copyright (C) 2014, 2015, 2016, 2017  J.H. Fitié, Vitens N.V.
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

%~d0
cd %~dp0

..\build\DBMTester.exe -i=sample1input.csv -st=24-11-2016 -et=29-11-2016 -f=intl | sort > sample1.csv
..\build\DBMTester.exe -i=sample2input.csv -st=12-3-2013 -et=13-3-2013 -f=intl | sort > sample2.csv
..\build\DBMTester.exe -i=sample3input.csv -c=sample3correlation.csv -st=1-1-2016 -et=2-1-2016 -f=intl | sort > sample3.csv
..\build\DBMTester.exe -i=sample4input.csv -c=sample4correlation.csv -st=13-11-2014 -et=14-11-2014 -f=intl | sort > sample4.csv
