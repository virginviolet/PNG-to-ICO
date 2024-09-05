@ECHO off
REM		Name :
REM					PNG to ICO
REM		Author :
REM					▄▄▄▄▄▄▄  ▄ ▄▄ ▄▄▄▄▄▄▄
REM					█ ▄▄▄ █ ██ ▀▄ █ ▄▄▄ █
REM					█ ███ █ ▄▀ ▀▄ █ ███ █
REM					█▄▄▄▄▄█ █ ▄▀█ █▄▄▄▄▄█
REM					▄▄ ▄  ▄▄▀██▀▀ ▄▄▄ ▄▄
REM					 ▀█▄█▄▄▄█▀▀ ▄▄▀█ █▄▀█
REM					 █ █▀▄▄▄▀██▀▄ █▄▄█ ▀█
REM					▄▄▄▄▄▄▄ █▄█▀ ▄ ██ ▄█
REM					█ ▄▄▄ █  █▀█▀ ▄▀▀  ▄▀
REM					█ ███ █ ▀▄  ▄▀▀▄▄▀█▀█
REM					█▄▄▄▄▄█ ███▀▄▀ ▀██ ▄
REM Console title
TITLE PNG to ICO
REM Script folder path
SET _directoryPath=%~dp0
REM Console height / width
MODE 65,30 | ECHO off
ECHO.
ECHO   -------------------------------------------------------------
ECHO                            PNG to ICO :
ECHO   -------------------------------------------------------------
ECHO.

MKDIR "%TEMP%\PNG-to-ICO" 2> nul

@REM SetLocal EnableDelayedExpansion
REM First command line argument
SET _argPath=%1
SET "_cubic_blur=1.05"
set "_normalSizes=256,128"
set "_sharpSizes=,96,64,48,32,24,16"
SET "_parameters=-resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=%_normalSizes%%_sharpSizes% "%%~df%%~pf%%~nf.ico""
REM Initialize variables
SET "_width="
SET "_height="
REM Get width
FOR /f "tokens=*" %%i IN ('magick identify -ping -format %%w %_argPath%') DO (
	SET "_width=%%i"
)
REM Get height
FOR /f "tokens=*" %%i IN ('magick identify -ping -format %%h %_argPath%') DO (
	SET "_height=%%i"
)
SET "_largestDimension="
if %_height% GEQ %_width% (
	SET "_largestDimension=%_height%"
	) ELSE (
		SET "_largestDimension=%_width%"
	)
REM If first argument is a directory
IF EXIST %_argPath%\* (
	REM ECHO Directory : %_argPath%
	REM Iterate through PNG, GIF, BMP, SVG and JPG files in directory
	FOR %%f IN (%_argPath%\*.png %_argPath%\*.bmp %_argPath%\*.gif %_argPath%\*.jpg %_argPath%\*.jpeg %_argPath%\*.svg) DO (
		REM Print file name (with extension)
		ECHO - %%~nf%%~xf
		REM CALL:SOMETHING
		REM TODO Remove below line
		REM Convert file to multi-resolution ICO
		"%_directoryPath%ImageMagick\magick.exe" "%%f" -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=256,128,96,64,48,32,24,16 "%%~df%%~pf%%~nf.ico"
	)
REM If first argument is a file
) ELSE (
	REM ECHO File : %_argPath%
	FOR %%f IN (%_argPath%) DO (
		REM Print file name (with extension)
		ECHO - %%~nf%%~xf
		SET "_nf=%%~nf"
		echo "dp %_directory_path%"
		for %%x in (%_normalSizes%) do (
			CALL :CONVERT_NORMAL %%x %_nf%
		)
		for %%x in (%_sharpSizes%) do (
			CALL :CONVERT_SHARP %%x %_nf% %_width% %_height% %_cubic_blur%
		)
		@REM SET "myvar="
		@REM ECHO _boxfilter !_boxfilter!
		@REM echo testone
		@REM for %%x in (%%)
		@REM CALL :CONVERT_NORMAL _argPath,_sharpSizes,cubic_b,_cubic_c,_boxfilter
		@REM CALL :CONVERT_SHARP _boxfilter
		@REM ECHO myvar %_boxfilter%
		REM Convert file to multi-resolution ICO
		REM TODO Remove below line
		"%_directoryPath%ImageMagick\magick.exe" %_argPath% -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=256,128,96,64,48,32,24,16 "%%~df%%~pf%%~nf.ico"
	)
)

GOTO :EOF

:CONVERT_NORMAL
	SET "_size=%~1"
	"%_directoryPath%ImageMagick\magick.exe" %_argPath% -resize 256x256^> -background none -gravity center -extent 256x256 "%TEMP%\PNG-to-ICO\%_nf%-%_size%.ico"
	EXIT /B

:CONVERT_SHARP
	SET "_size=%~1"
	REM Initialize variables
	SET "_scaleFactor="
	SET "_cubic_b="
	SET "_cubic_c="
	SET "_boxfilter="
	echo %_size%
	echo %_largestDimension%
	SET /A "_scaleFactor=_largestDimension/_size"
	echo %_scaleFactor% 
	@REM echo %_cubic_blur%
	@REM echo ok
	EXIT /B