@IF [%1]==[] (ECHO No version specified && EXIT /B 1)
@IF [%2]==[] (ECHO Outdir missing && EXIT /B 1)
@IF [%3]==[] (ECHO Processor architecture missing && EXIT /B 1)

@SET Version=%1
@candle -ext WixUtilExtension.dll -ext WixUIExtension.dll Main.wxs -dProcessorArchitecture=%3
@IF NOT %ERRORLEVEL% == 0 (ECHO Build failed && EXIT /B 1)
@light -ext WixUtilExtension.dll -ext WixUIExtension.dll Main.wixobj -out %2\Sunet.Mailfilter.OutlookPlugin-%1_%3.msi
@echo Running setupbld...
@setupbld.exe -out %2\Sunet.Mailfilter.OutlookPlugin-%1_%3.exe -msu %2\Sunet.Mailfilter.OutlookPlugin-%1_%3.msi -setup setup1.exe
