@echo off
tools\sqlite3.exe tools\opendatatest.db ".read tools/commandstest.txt"
pause