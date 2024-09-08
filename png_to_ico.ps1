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

# Temporary directory path
$tempDir = Join-Path -Path $env:TEMP -ChildPath "PNG-to-ICO"

# First command line argument
$argPath = $args[0]

# ImageMagick executable
$magick = Join-Path $scriptDir -ChildPath "ImageMagick\magick.exe"

function ConvertTo-IcoMultiResOld($image, $icon) {
	$sizes = "256, 128, 96, 64, 48, 32, 24, 16"
	& $magick $image -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=$sizes $icon
}

function ConvertTo-IcoMultiRes($image, $icon) {
	if (-Not [bool](Test-Path -Path $tempDir)) {
		New-Item -Path $thisIconTempDir -ItemType "directory"
	}
	$fileBaseName = (Get-ChildItem -Path $image).BaseName
	$iconTempDirName = $fileBaseName + "_" + (New-Guid).Guid
	$iconTempDir = Join-Path -Path $tempDir -ChildPath $iconTempDirName
	# Create directory (and suppress Success stream output)
	New-Item -Path $iconTempDir -ItemType "directory" 1> $null
	# $sizePng = (256)
	$sizesPngNormal = @(256)
	$sizesPngSharper = @()
	$sizesIcoNormal = @(128, 96)
	$sizesIcoSharper = @(64, 48, 32, 24, 16)
	$width = & $magick identify -ping -format '%w' $image
	$height = & $magick identify -ping -format '%h' $image
	
	# Initialize variables
	$largestDimension = ""
	$iconSized = ""
	$algorithm = ""
	$outputCount = 0
	$outputFiles = @()
	$outputExt = ""
	$outputFile = ""

	if ($height -ge $width) {
		$largestDimension = $height
	}
	else {
		$largestDimension = $width
	}
	# echo $largestDimension
	# if ($largestDimension -le $sizes_) {}
	foreach ($size in $sizesPngNormal) {
		$outputExt = "png"
		# is less than or equal
		if ($size -le $largestDimension) {
			$iconSizedName = "$fileBaseName-$outputCount.$outputExt"
			$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
			Convert-ImageNormal $image $size $iconSized
			$outputCount++
			$outputFiles += $iconSized
		}
	}
	<# foreach ($size in $sizesPngSharper) {
		$outputExt = "png"
		# is less than or equal
		if ($size -le $largestDimension) {
			$iconSizedName = "$fileBaseName-$outputCount.$outputExt"
			$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
			Convert-ImageNormal $image $size $iconSized
			$outputCount++
			$outputFiles += $iconSized
		}
	} #>
	foreach ($size in $sizesIcoNormal) {
		$outputExt = "ico"
		# is less than or equal
		if ($size -le $largestDimension) {
			$iconSizedName = "$fileBaseName-$outputCount.$outputExt"
			$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
			Convert-ImageNormal $image $size $iconSized
			$outputCount++
			$outputFiles += $iconSized
		}
	}
	foreach ($size in $sizesIcoSharper) {
		$outputExt = "ico"
		# is less than or equal
		if ($size -le $largestDimension) {
			$iconSizedName = "$fileBaseName-$outputCount.$outputExt"
			$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
			Convert-ImageNormal $image $size $iconSized
			$outputCount++
			$outputFiles += $iconSized
		}
	}
	# & $magick $image -resize^> $largestDimension
	# foreach ($s in $sizes) {
	# 	if $width 
	# }
	# echo $iconTempDir
	# echo child
	# $test = Get-ChildItem -Path $iconTempDir
	# echo $test
	# & $magick $image -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=$sizes $icon
}

function Convert-ImageNormal ($image, $size, $outputFile) {
	# echo $image $size $outputFile
	& $magick $image -resize $size -background none -gravity center -extent $size $outputFile
}

function Convert-ImageSharp ($image, $size, $outputFile) {

	& $magick $image -resize $size -background none -gravity center -extent $size $outputFile
	return
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
		ConvertTo-IcoMultiRes $image $icon
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
	ConvertTo-IcoMultiRes $image $icon
}