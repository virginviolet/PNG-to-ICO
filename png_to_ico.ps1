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
# Check if RawUI is available before trying to set window properties
if ($null -ne $host.UI.RawUI) {
	$host.UI.RawUI.WindowTitle = "PNG to ICO"
    
	try {
		# Attempt to change window size
		$host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(65, 30)
	} catch {
		Write-Debug "Window size change is not supported in this environment."
	}

}

Write-Output "`n"
Write-Output  " -------------------------------------------------------------"
Write-Output  "                          PNG to ICO :"
Write-Output  " -------------------------------------------------------------"
Write-Output "`n"

# Script directory path
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# First command line argument
$argPath = $args[0]

# ImageMagick executable
$magick = Join-Path $scriptDir -ChildPath "ImageMagick\magick.exe"

function ConvertTo-Ico($image, $icon) {
	$sizes = "256, 128, 96, 64, 48, 32, 24, 16"
	& $magick $image -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=$sizes $icon
}

# If first argument is a directory
IF ([bool](Test-Path $argPath -PathType container)) {
	
	Write-Verbose "Directory : $argPath"

	# Get images in directory
	$dirContents = Join-Path $argPath -ChildPath '*'
	$images = Get-ChildItem ($dirContents) -Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg

	# Iterate through the images
	FOREACH ($i IN $images) {

		# Print file name (with extension)
		$fileName = $i.Name
		Write-Output "- $fileName"

		# Convert file to multi-resolution ICO
		$image = Resolve-Path -Path $i
		$dir = Resolve-Path -Path $argPath
		$fileBaseName = (Get-Item -Path $i).BaseName # Name without extension
		$icon = Join-Path -Path $dir -ChildPath "$fileBaseName.ico"
		ConvertTo-Ico $image $icon
	}
	# If first argument is a file
} ELSE {
	
	Write-Verbose "File : $argPath"

	# Print file name (with extension)
	$fileName = (Get-Item -Path $argPath).Name
	Write-Output "- $fileName"

	# Convert file to multi-resolution ICO
	$image = $argPath
	$dir = (Get-Item -Path $argPath).Directory
	$fileBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
	$icon = Join-Path -Path $dir -ChildPath "$fileBaseName.ico"
	ConvertTo-Ico $image $icon
}