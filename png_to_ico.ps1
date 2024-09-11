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

# If no parameter is specified
if ($null -eq $argPath) {
	# Explain how to use this script
	Write-Output "Usage: png_to_ico.ps1 <file|directory> [<allow upscale>]"
	Write-Output "`nExamples`n"
	Write-Output "png_to_ico.ps1 C:\Users\John\Desktop\my_image.png"
	Write-Output "png_to_ico.ps1 ""C:\Users\John\Desktop\my images"""
	Write-Output "png_to_ico.ps1 ""C:\Users\John\Desktop\my tiny image.png"" true"
	return
}
else {
	# Turn the string into a path
	$argPath = Resolve-Path -Path $argPath
}

# Set to $true to to include sizes in multi-res ICOs that exceeds the original size of the image
$upscale = $args[1]
if ($null -eq $upscale) {
	# Set to $false if not specified
	$upscale = $false
}

# ImageMagick executable
$magick = Join-Path $scriptDir -ChildPath "ImageMagick\magick.exe"

# Uncomment to set custom path
# $magick = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"

# If magick.exe is not found in specified directory
if (-not [bool](Test-Path -Path $magick)) {
	if ([bool](Get-Command magick)) {
		# Use magick.exe from environment variable if found
		$magick = "magick"
	}
	else {
		# Write and throw error
		Write-Error "ImageMagick not found."
		throw
	}
}

# Old method to convert multi-res ico
function ConvertTo-IcoMultiResOld {
	param (
		$image,
		$icon
	)
	$sizes = "256, 128, 96, 64, 48, 32, 24, 16"
	& $magick $image -resize 256x256^> -background none -gravity center -extent 256x256 -define icon:auto-resize=$sizes $icon
}

# TODO Order of functions 
function ConvertTo-IcoMultiRes {
	param (
		$image,
		$icon
	)
	$fileBaseName = (Get-ChildItem -Path $image).BaseName
	$iconTempName = $fileBaseName + "_" + (New-Guid).Guid
	# XXX
	# FIXME It adds icon to $tempDir instead of a subdirectory for all images
	$iconTempDir = Join-Path -Path $tempDir -ChildPath $iconTempName
	# Create temporary directory for this icon, for storing icons in different sizes (and suppress Success stream output)
	if (-Not [bool](Test-Path -Path $iconTempDir)) {
		New-Item -Path $iconTempDir -ItemType "directory" 1> $null
	}
	$sizesPngNormal = @(256)
	$sizesPngSharper = @()
	$sizesIcoNormal = @(128)
	$sizesIcoSharper = @(96, 64, 48, 32, 24, 16)
	$width = [int](& $magick identify -ping -format '%w' $image)
	$height = [int](& $magick identify -ping -format '%h' $image)
	
	# Initialize variables
	$largestDimension = ""
	Set-Variable -Name outputCount -Value 0 -Scope Script
	# $outputFiles = @()

	if ($height -ge $width) {
		$largestDimension = $height
	}
	else {
		$largestDimension = $width
	}
	# echo $largestDimension
	# echo $fileBaseName
	# The canonical way to do this is `$outputCount = ConvertFrom-SizeList <...>`, but that would look misleading in this case, imo
	ConvertFrom-SizeList $sizesPngNormal $largestDimension "normal" $image $iconTempDir $outputCount $fileBaseName "png"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesPngSharper $largestDimension "sharper" $image $iconTempDir $outputCount $fileBaseName "png"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesIcoNormal $largestDimension "normal" $image $iconTempDir $outputCount $fileBaseName "ico"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesIcoSharper $largestDimension "sharper" $image $iconTempDir $outputCount $fileBaseName "ico"
	Get-Variable -Name outputCount -Scope Script 1> $null

	# Take all the icon files created (same image, different sizes) and assemble into a multi-res icon
	Merge-Icons $iconTempDir $iconTempName $icon
	Remove-Item $iconTempDir -Recurse
}

function ConvertFrom-SizeList {
	param (
		$sizeList,
		$dimension,
		$algorithm,
		$inputImage,
		$outputDir,
		$outputCount,
		$outputBaseName,
		$outputExt
	)
	
	$inputExt = (Get-ChildItem -Path $inputImage).Extension
	$outputCountLocal = $outputCount

	# Initialize variables
	$iconSizedName = ""
	$iconSized = ""

	# echo "$dimension $sizesPngNormal[-1] $sizesPngSharper[-1] $sizesIcoNormal[-1]) $sizesIcoSharper[-1]"
	# echo $dimension
	# echo $sizesIcoSharper[-1]
	foreach ($size in $sizeList) {
		# echo "l $outputCountLocal"
		# Only resize if original image is greater than $size
		$iconSizedName = "$outputBaseName-$outputCountLocal.$outputExt"
		# echo "'$outputBaseName'"
		# echo "'$iconSizedName'"
		# echo $size
		$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
		if ($dimension -gt $size) {
			# echo one
			if ($algorithm -eq "sharper") {
				# echo "$size sharp"
				Convert-ImageResizeSharper $inputImage $size $iconSized $dimension
			}
			elseif ($algorithm -eq "normal") {
				# echo "$size normal"
				Convert-ImageResizeNormal $inputImage $size $iconSized
			}
			else {
				Write-Error "Incorrect algorithm '$algorithm'"
			}
			$outputCountLocal++
			# Get-Variable -Name outputCount -Scope Script
			# echo "l $outputCountLocal"
		}
		# If original image is equal to $size, and the format is different, then convert without resize
		elseif (($dimension -eq $size) -and ($outputExt -ne $inputExt)) {
			# echo two
			# [x] Test
			Convert-Image $inputImage $iconSized
			$outputCountLocal++
		}
		# If original image is equal to $size, and the format is the same, then simply copy the file to our temporary location where it will be included
		elseif (($dimension -eq $size) -and ($outputExt -eq $inputExt)) {
			echo three
			# [x] Test
			Copy-Item -Path "$inputImage" -Destination "$iconSized"
			$outputCountLocal++
		}
		# If upscaling is enabled, and the original image is smaller than $size
		elseif (($upscale -eq $true) -and ($dimension -lt $size)) {
			Convert-ImageResizeNormal $inputImage $size $iconSized
			$outputCountLocal++
		}
		# If original image is smaller than anything in the lists (and upscaling is disabled)
		# (`[-1]` gets the last item i the list. The `$null -eq` check is necessary in case the list is empty.)
		elseif (
			($dimension -lt $sizesPngNormal[-1] -or $null -eq $sizesPngNormal[-1]) -and
			($dimension -lt $sizesPngSharper[-1] -or $null -eq $sizesPngSharper[-1]) -and
			($dimension -lt $sizesIcoNormal[-1] -or $null -eq $sizesIcoNormal[-1]) -and
			($dimension -lt $sizesIcoSharper[-1] -or $null -eq $sizesIcoSharper[-1])) {
				# IMPROVE Make this not run more times than necessary. It currently runs onece for each time time this function is called.
				# This piece of code could be moved to ConvertTo-IcoMultiRes, which would solve the problem, but it feels more organized to keep it here. The performance loss is negligible, unless, perhaps, someone tries to batch convert a huge amount of tiny files, but that's not a very plausible scenario, I think.
				Convert-Image $inputImage $iconSized
				$outputCountLocal++
				return
			}

		# Pause
	}
	Set-Variable -Name outputCount -Value $outputCountLocal -Scope Script
	# Get-Variable -Name outputCount -Scope Script
	# echo "gccc $outputCount"

}

function Convert-Image {
	param (
		$image,
		$outputFile
	)
	# echo "command: '$magick ""$image"" ""$outputFile""'"
	& $magick "$image" "$outputFile"
}
function Convert-ImageResizeNormal {
	param (
		$inputFile,
		$outputSize,
		$outputFile
	)
	# Convert+Resize with Image Magick
	$resize = "$outputSize" + "x" + "$outputSize"
	# & $magick "$inputFile" -resize $resize -background none -gravity center -extent $outputSize "$outputFile"
	& $magick "$inputFile" `
	-resize $resize `
	"$outputFile"
}

function Convert-ImageResizeSharper {
	param (
		$inputFile,
		$outputSize,
		$outputFile,
		$inputSize)
	$cubicBValue = 0.0
	$cubicBlurValue = 1.05
	$boxBlurValue = 0.707

	# initialize variables
	$scaleFactor = 0.0
	$cubicCValue = 0.0
	$useBoxFilter = $false
	$boxFilterSize = 0
	$boxFilterCrop1 = ""
	$boxFilterCrop2 = ""
	$boxFilterParameters = ""
	$cubcFiltersParameters = ""
	$parameters = ""

	# Determine scale factor
	# echo "$outputSize / $inputSize"
	$scaleFactor = $outputSize / $inputSize
	# echo $scaleFactor
	if ($scaleFactor -gt 1) {
		# [x] test
		# Image enlargement
		# echo enlargement
		$cubicCValue = 1.0
	}
	elseif ($scaleFactor -ge 0.25 -and $scaleFactor -lt 1) {
		# echo here
		$cubicCValue = 2.6 - (1.6 * $scaleFactor)
	}
	elseif (<# $scaleFactor -le ? -and  #>$scaleFactor -lt 0.25) {
		# echo box
		$useBoxFilter = $true
		$boxFilterSize = 4 * $outputSize
		$cubicCValue = 2.2
		$box_filter_scale_factor_a = $boxFilterSize / $inputSize
		$box_filter_scale_factor_b = 1 - ($boxFilterSize / $inputSize)
		# SRT = ScaleRotateTranslate
		# +distort SRT X,Y ScaleX,ScaleY Angle NewX,NewY
		$boxFilterSrt1 = "0,0 $box_filter_scale_factor_a,1.0 0 $box_filter_scale_factor_b,0.0"
		# -crop: width x height + x offset + y offset
		$boxFilterCrop1 = "$boxFilterSize" + "x" + "$inputSize+0+0"
		# +distort SRT X,Y ScaleX,ScaleY Angle NewX,NewY
		$boxFilterSrt2 = "0,0 1.0,$box_filter_scale_factor_a 0 0.0,$box_filter_scale_factor_a"
		# -crop: width x height + x offset + y offset
		$boxFilterCrop2 = "$boxFilterSize" + "x" + "$boxFilterSize+0+0"
	}
	elseif ($scaleFactor -eq 1) {
		Write-Warning "Attempted to resize with scale factor 1.0."
		Pause
	}

	# "Cubic Filters". https://imagemagick.org/Usage/filter/#cubics
	# "Box". https://imagemagick.org/Usage/filter/#box
	# "Scale-Rotate-Translate (SRT) Distortion". https://imagemagick.org/Usage/distorts/#srt
	# "-crop". https://imagemagick.org/script/command-line-options.php#crop
	# "Transpose and Transverse, Diagonally". https://imagemagick.org/Usage/warping/#transpose
	# "Resizing Images". https://legacy.imagemagick.org/Usage/resize/#resize
	# `-resize` keeps aspect ratio by default.
	$resize = "$outputSize" + "x" + "$outputSize"
	# echo $outputFile
	if ($useBoxFilter) {
		# echo box
		& $magick "$inputFile" `
		-filter box -define filter:blur=$boxBlurValue `
		+distort SRT "$boxFilterSrt1" -crop $boxFilterCrop1 `
		+distort SRT "$boxFilterSrt2" -crop $boxFilterCrop2 `
		-transpose `
		-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
		-resize $resize `
		-transpose `
		"$outputFile"
	}
	else {
		# [x] Compare with and without transpose (It makes no difference)
		# echo "no box"
		# echo "command: '$magick ""$inputFile"" -transpose -filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue -resize $resize -transpose ""$outputFile""'"
		& $magick "$inputFile" `
		-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
		-resize $resize `
		"$outputFile"
	}
}

function Merge-Icons {
	param (
		$inputDir,
		$tempName,
		$outputIcon
	)
	echo "outputicon $outputIcon"
	$files = Join-Path -Path $inputDir -ChildPath "\*"
	$tempFile = Join-Path -Path $tempDir -ChildPath "$tempName.ico"
	& $magick "$files" "$tempFile"
	
	# Check if exists
	if ([bool](Test-Path -Path $outputIcon)) {
		Move-Item -Path $tempFile -Destination $outputIcon -Force -Confirm
	}
	else {
		Move-Item -Path $tempFile -Destination $outputIcon
	}
}

# Create temporary directory for this script (unless it for some reason exists, and suppress Success stream output)
if (-Not [bool](Test-Path -Path $tempDir)) {
	New-Item -Path $tempDir -ItemType "directory" 1> $null
}

# TODO change "fileBaseName" in variables
# If first argument is a directory
IF ([bool](Test-Path $argPath -PathType container)) {
	
	Write-Verbose "Directory : $argPath"

	# Name of the directory
	# In this script, "image" means image that is not yet converted to ico.
	$imagesDirName = (Get-Item -Path $argPath).Name
	# Name of subdirectory to save the icons to
	$iconsDirName = "ICO"
	$iconsDir = Join-Path -Path $argPath -ChildPath $iconsDirName

	# Create temporary directory to store the files in (and suppress Succes stream output)
	# (This enables for overwriting all existing files with "[A]")
	$tempIconsDirName = "$imagesDirName" + "_" + (New-Guid).Guid
	$tempIconsDir = Join-Path -Path $tempDir -ChildPath "$tempIconsDirName"
	New-Item $tempIconsDir -ItemType "directory" 1> $null

	# Get images in directory
	$dirContents = Join-Path $argPath -ChildPath '*'
	# TODO Add more formats? Ico?
	$images = Get-ChildItem ($dirContents) -Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg
	# Iterate through the images
	foreach ($i in $images) {

		# Print file name (with extension)
		$fileName = $i.Name
		Write-Output "- $fileName"

		# Convert file to multi-resolution ICO
		$image = $i
		$dir = $tempIconsDir
		$fileBaseName = (Get-Item -Path $i).BaseName # Name without extension
		$icon = Join-Path -Path $dir -ChildPath "$fileBaseName.ico"
		# echo $icon
		ConvertTo-IcoMultiRes $image $icon

		# Move each file from the temporary directory to a subdirectory in the original directory
		$icons = Get-ChildItem ($tempIconsDir)
	
	foreach ($i in $icons) {
		Move-Item -Path $i -Destination $iconsDir
	}

	# Remove-Item $tempDir -Recurse
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

	# Remove-Item $tempDir -Recurse
}