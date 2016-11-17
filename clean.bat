@echo off
%~d0
cd %~dp0

if not exist build mkdir build

del /Q build\*
