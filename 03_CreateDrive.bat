@echo OFF
echo Copyright (c) 2020, 2021 Demxi (Pty) Ltd
echo Mapping s: drive...
subst g: /D
subst g: "%CD%"
g:
@echo ON
pause