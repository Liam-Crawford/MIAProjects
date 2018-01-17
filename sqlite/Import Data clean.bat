@echo off
tools\sqlite3.exe tools\opendataclean.db ".read tools/commandsclean.txt"
pause