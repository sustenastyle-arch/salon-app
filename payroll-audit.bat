@echo off
rem Payroll audit for the pay period that just ended (run by Task Scheduler at 9pm on the 15th and the last day of the month).
rem Double-click to run it by hand. The report opens in Notepad and is saved on the Desktop.
cd /d "%~dp0"
node scripts\payroll-audit.mjs %*
