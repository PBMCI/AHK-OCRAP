@echo off
for /f "tokens=1" %%A in ('assoc .ahk 2^>nul') do (
    set "file_association=%%A"
)

echo "%file_association%"

if not defined file_association (
    echo No association
)