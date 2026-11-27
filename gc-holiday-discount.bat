@echo off
rem Turns the Square gift card 10% discount on or off (run by Task Scheduler). Usage: gc-holiday-discount.bat on|off
cd /d "%~dp0"
node scripts\gc-holiday-discount.mjs %*
