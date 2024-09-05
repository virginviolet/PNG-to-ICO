<#
 	Name :
				PNG to ICO
	Author :
				▄▄▄▄▄▄▄  ▄ ▄▄ ▄▄▄▄▄▄▄
				█ ▄▄▄ █ ██ ▀▄ █ ▄▄▄ █
				█ ███ █ ▄▀ ▀▄ █ ███ █
				█▄▄▄▄▄█ █ ▄▀█ █▄▄▄▄▄█
				▄▄ ▄  ▄▄▀██▀▀ ▄▄▄ ▄▄
				 ▀█▄█▄▄▄█▀▀ ▄▄▀█ █▄▀█
				 █ █▀▄▄▄▀██▀▄ █▄▄█ ▀█
				▄▄▄▄▄▄▄ █▄█▀ ▄ ██ ▄█
				█ ▄▄▄ █  █▀█▀ ▄▀▀  ▄▀
				█ ███ █ ▀▄  ▄▀▀▄▄▀█▀█
				█▄▄▄▄▄█ ███▀▄▀ ▀██ ▄ 
#>

# Console title
$host.ui.RawUI.WindowTitle = "PNG to ICO"

# Script directory path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Console height / width
$host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(65, 30)

ECHO "`n"
ECHO  " -------------------------------------------------------------"
ECHO  "                          PNG to ICO :"
ECHO  " -------------------------------------------------------------"
ECHO "`n"

# First command line argument
$argPath = $args[0]

# ImageMagick executable
$magick = Join-Path $scriptDir -ChildPath "ImageMagick\magick.exe"

# If first argument is a directory
IF ([bool](Test-Path $argPath -PathType container)) {
	
	# ECHO "Directory : $argPath"

	# Get images in directory
	$dirContents = Join-Path $argPath -ChildPath '*'
	$images = Get-ChildItem ($dirContents) -Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg

	# Iterate through the images
	FOREACH ($i IN $images) {

		# Print file name (with extension)
		$fileName = $i.Name
		ECHO "- $fileName"

		# Convert file to multi-resolution ICO
		$dir = Resolve-Path $argPath
		$fileBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
		$icon = Join-Path -Path $dir -ChildPath "$fileBaseName.ico"
		& $magick $argPath -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=256,128,96,64,48,32,24,16 $icon
	}
# If first argument is a file
} ELSE {
	
	# ECHO File : $argPath

	# Print file name (with extension)
	$fileName = (Get-Item -Path $argPath).Name
	ECHO "- $fileName"

	# Convert file to multi-resolution ICO
	$dir = (Get-Item -Path $argPath).Directory
	$fileBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
	$icon = Join-Path -Path $dir -ChildPath "$fileBaseName.ico"
	echo $icon
	& $magick $argPath -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=256,128,96,64,48,32,24,16 $icon
}