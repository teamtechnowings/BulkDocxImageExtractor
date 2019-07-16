
setlocal
SET documents=F:\work\docx\
SET images=F:\work\images\

CD %documents%
%documents:~0,2%
FOR %%i IN (*.docx) DO (
set "filedrive=%%~di"
set "filepath=%%~pi"
set "filename=%%~ni"
set "fileextension=%%~xi"


mkdir "%images%/%%~ni%%~xi"
rename %%~ni%%~xi "%%~ni%%~xi.zip"
mkdir "%images%%%~ni%%~xi\"
Call :UnZipFile "%images%%%~ni%%~xi\" "%documents%%%~ni%%~xi.zip"
xcopy /Y "%images%%%~ni%%~xi\word\media" "%images%%%~ni%%~xi_images\" 
rmdir /s /Q "%images%%%~ni%%~xi\"
rename %%~ni%%~xi.zip "%%~ni%%~xi"
)

)

exit /b

:UnZipFile <ExtractTo> <newzipfile>
set vbs="%temp%\_.vbs"
if exist %vbs% del /f /q %vbs%
>%vbs%  echo Set fso = CreateObject("Scripting.FileSystemObject")
>>%vbs% echo If NOT fso.FolderExists(%1) Then
>>%vbs% echo fso.CreateFolder(%1)
>>%vbs% echo End If
>>%vbs% echo set objShell = CreateObject("Shell.Application")
>>%vbs% echo set FilesInZip=objShell.NameSpace(%2).items
>>%vbs% echo objShell.NameSpace(%1).CopyHere(FilesInZip)
>>%vbs% echo Set fso = Nothing
>>%vbs% echo Set objShell = Nothing
cscript //nologo %vbs%
if exist %vbs% del /f /q %vbs%