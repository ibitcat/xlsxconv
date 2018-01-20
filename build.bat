echo off

rsrc -ico="go.ico" -manifest="go.manifest"
go build -v -ldflags "-w -s -H windowsgui" -o ./xlsxconv.exe
upx xlsxconv.exe

pause