#!/bin/bash

# Dynamic Bandwidth Monitor
# Leak detection method implemented in a real-time data historian
# Copyright (C) 2014-2019  J.H. Fitié, Vitens N.V.
#
# This file is part of DBM.
#
# DBM is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# DBM is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with DBM.  If not, see <http://www.gnu.org/licenses/>.

# Update copyright years
grep -q \\-$(date +%Y)\ \  $0 || find . -type f -exec \
sed -i -E "s/^(.*Copyright \(C\) [0-9]{4}\-)[0-9]{4}(  .*)$/\
\1$(date +%Y)\2/g" {} \;

# Run build script using Wine
export WINEDEBUG=fixme-all,err-ole
export WINEDLLOVERRIDES=mscoree=n
wine cmd /c build.bat
