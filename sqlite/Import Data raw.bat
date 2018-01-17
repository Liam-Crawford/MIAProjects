@echo off
tools\sqlite3.exe tools\opendataraw.db ".read tools/commandsraw.txt"
pause