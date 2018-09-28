@echo off
setlocal EnableDelayedExpansion
set InputFile=path to your CSV file
for /f "tokens=1-8 delims=," %%A in ('type "%InputFile%"') do (
  py sync.py %%A:%%B:%%C:%%D %%E:%%F:%%G:%%H --from=2016-1-1 --zone="Africa/Nairobi"
)