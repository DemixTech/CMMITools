@echo OFF
echo Copyright (c) 2020, 2021 Demxi (Pty) Ltd
echo Mapping s: drive...
subst s: /D
subst s: "%CD%"
s:
@echo ON
pause