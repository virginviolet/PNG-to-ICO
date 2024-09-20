#########################################################################################
#
#	Name :
#				PNG to ICO
#	Author :
#				▄▄▄▄▄▄▄  ▄ ▄▄ ▄▄▄▄▄▄▄
#				█ ▄▄▄ █ ██ ▀▄ █ ▄▄▄ █
#				█ ███ █ ▄▀ ▀▄ █ ███ █
#				█▄▄▄▄▄█ █ ▄▀█ █▄▄▄▄▄█
#				▄▄ ▄  ▄▄▀██▀▀ ▄▄▄ ▄▄
#				 ▀█▄█▄▄▄█▀▀ ▄▄▀█ █▄▀█
#				 █ █▀▄▄▄▀██▀▄ █▄▄█ ▀█
#				▄▄▄▄▄▄▄ █▄█▀ ▄ ██ ▄█
#				█ ▄▄▄ █  █▀█▀ ▄▀▀  ▄▀
#				█ ███ █ ▀▄  ▄▀▀▄▄▀█▀█
#				█▄▄▄▄▄█ ███▀▄▀ ▀██ ▄ 
#
#########################################################################################

#########################################################################################
#
#region 1. Script settings and initialization
#
#########################################################################################

# Stop on error
$ErrorActionPreference = "Stop"

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
$scriptDirPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

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
} else {
	# Get full path
	$argPath = Convert-Path -Path $argPath
}

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 2. User configuration
#
#########################################################################################


# Set which subicons to create, based on image size and resize algorithm
$sizesPngNormal = @(256)
$sizesPngSharper = @()
$sizesIcoNormal = @(128)
$sizesIcoSharper = @(96, 64, 48, 32, 24, 16)

# Set to $true to always include selected subicon sizes (resolutions) in multi-res ICOs even when they exceed the original size of the image
$upscale = $args[1]

# ImageMagick executable
$magick = Join-Path $scriptDirPath -ChildPath "ImageMagick\magick.exe"
# Uncomment to set custom path:
# $magick = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"

# Used for converting all images in a directory:
# Subdirectory, relative to input directory, to save all icons in
# Default: Save into a subdirectory named "ICO"
$iconsFinalDirPath = Join-Path -Path $argPath -ChildPath "ICO"
# echo "argPath $argPath"
# echo "iconsFinalDir $iconsFinalDirPath"
# Uncomment to set custom path:
# $iconsFinalDirPath = "C:\Users\John\My icons"

# Used for single file conversion:
# Subdirectory, relative to input image's directory, to save the icon in
# Default: Save icon to the same folder as input image
$singleIconFinalDirName = ""
# Uncomment to set custom path:
# $singleIconFinalDirPath = "C:\Users\John\My icons" 

# Temporary directory path
$tempDirPath = Join-Path -Path $env:TEMP -ChildPath "PNG-to-ICO"

$subiconsDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Unfinished\Subicons"

# Temporary directory for unfinished icon directories, used before all images has been converted (only for when first argument is a directory)
# (Saving all files to one directory and move all files at the end, allows for the user to overwrite all files with "[A]", in the edge case where there are existing files with the same names and extensions)
# Default: Unfinished icon directories go into "Unfinished" parent directory.
$unfinishedDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Unfinished"

# Temporary directory for finished icons (when first argument is a file)
# Default: Icons go directly in "Finished" directory.
$finishedSingleIconDirPath = Join-Path -Path $tempDirPath -ChildPath "Finished"

# Temporary parent directory for a finished icons subdirectory (when first argument is a directory)
# Default: icon directories go into "Finished"
$finishedDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Finished"

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 3. Finalization
#
#########################################################################################

if ($null -eq $upscale) {
	# Set to $false if not specified
	$upscale = $false
}

# If magick.exe is not found in specified directory
if (-not [bool](Test-Path -Path $magick)) {
	if ([bool](Get-Command magick)) {
		# Use magick.exe from environment variable if found
		$magick = "magick"
	} else {
		# Write and throw error
		Write-Error "ImageMagick not found."
		throw
	}
}

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 4. Function definitions
#
#########################################################################################

# Functions are defined in the the order they are first called.

function ConvertTo-IcoMultiRes {
	param (
		# Add aliases that mimics offcical cmdlet parameters,
		# for no reason
		[Alias("Path")][string]$InputPath, # takes string or FileInfo object
		[string]$BaseName,
		[Alias("Destination")][string]$OutputPath # takes string or FileInfo object
	)

	$subiconsDirPath = Join-Path -Path $subiconsDirParentPath -ChildPath $BaseName
	
	# Create temporary directory for this icon, for storing subicons
	if (-Not [bool](Test-Path -Path $subiconsDirPath)) {
		New-Item -Path $subiconsDirPath -ItemType "directory" 1> $null
	}

	# Get width and height of input image
	$width = [int](& $magick identify -ping -format '%w' $InputPath)
	$height = [int](& $magick identify -ping -format '%h' $InputPath)

	# Only the largest dimension is needed
	if ($height -ge $width) {
		$imageSize = $height
	} else {
		$imageSize = $width
	}

	# echo $imageSize
	# echo $BaseName

	# Make subicons
	# We need the function to update the value of $subiconCount. The canonical way to do this is probably:
	# `$subiconCount = ConvertFrom-SubiconsArray <...>`
	# But that would make it look like the purpose of that function is only to update the subiconCount value. So instead, we set variable in the Script scope here, and then update it in the other function.
	# $subiconCount is suffixed to the subicon file names.
	$script:subiconCount = 0 # Initialize variable
	# echo $subiconsDirPath
	ConvertFrom-SubiconsArray -SubiconsArray $sizesPngNormal `
		-InputSize $imageSize `
		-Algorithm "Normal" `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".png"
	# Pause
	ConvertFrom-SubiconsArray -SubiconsArray $sizesPngSharper `
		-InputSize $imageSize `
		-Algorithm "Sharper" `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".png"
	# Pause
	ConvertFrom-SubiconsArray -SubiconsArray $sizesIcoNormal `
		-InputSize $imageSize `
		-Algorithm "Normal" `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".ico"
	ConvertFrom-SubiconsArray -SubiconsArray $sizesIcoSharper `
		-InputSize $imageSize `
		-Algorithm "Sharper" `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".ico"
	
	New-MultiResIcon -InputDirectoryPath $subiconsDirPath -OutputPath $OutputPath
	# Pause
	# Take all the subicons created and assemble into a multi-res icon
	Remove-Item -Path $subiconsDirParentPath -Recurse
}

function ConvertFrom-SubiconsArray {
	param (
		[Alias("InputObject")][array]$SubiconsArray,
		[Alias("Size")][int]$InputSize,
		[string]$Algorithm,
		[Alias("Path")][string]$InputPath,
		[Alias("Destination")][string]$OutputPath,
		[string]$BaseName,
		[Alias("Extension")][string]$OutputExt
	)

	# TODO Algorithm string to Enum

	# Get input image extension
	$inputExt = (Get-Item -Path $InputPath).Extension

	# echo "$InputSize $sizesPngNormal[-1] $sizesPngSharper[-1] $sizesIcoNormal[-1]) $sizesIcoSharper[-1]"
	# echo $InputSize
	# echo $sizesIcoSharper[-1]
	foreach ($subiconSize in $SubiconsArray) {
		# echo "SIC $subiconCount"
		# Only resize if original image is greater than $subiconSize
		$subiconName = "$BaseName-$subiconCount" + $OutputExt
		# echo "'$BaseName'"
		# echo "'$subiconName'"
		# echo $subiconSize
		$subiconPath = Join-Path -Path $OutputPath -ChildPath $subiconName
		if ($InputSize -gt $subiconSize) {
			# echo one
			if ($Algorithm -eq "Sharper") {
				# echo "$subiconSize sharp"
				Convert-ImageResizeSharper -InputPath $InputPath -InputSize $InputSize -OutputSize $subiconSize -OutputPath $subiconPath
			} elseif ($Algorithm -eq "Normal") {
				# echo "$subiconSize normal"
				Convert-ImageResizeNormal -InputPath $InputPath -OutputSize $subiconSize -OutputPath $subiconPath
			} else {
				Write-Error "Incorrect algorithm '$Algorithm'"
			}
			$script:subiconCount++
			# Get-Variable -Name subiconCount -Scope Script
			# echo "SIC $subiconCount"
		}
		# If original image is equal to $subiconSize, and the format is different, then convert without resize
		elseif (($InputSize -eq $subiconSize) -and ($OutputExt -ne $inputExt)) {
			# echo two
			Convert-Image -InputPath $InputPath -OutputPath $subiconPath
			$script:subiconCount++
		}
		# If original image is equal to $subiconSize, and the format is the same, then simply copy the image to the subicons directory
		elseif (($InputSize -eq $subiconSize) -and ($OutputExt -eq $inputExt)) {
			# echo three
			Copy-Item -Path "$InputPath" -Destination "$subiconPath"
			$script:subiconCount++
		}
		# If upscaling is enabled, and the original image is smaller than $subiconSize
		elseif (($upscale -eq $true) -and ($InputSize -lt $subiconSize)) {
			Convert-ImageResizeNormal -InputPath $InputPath -OutputSize $subiconSize -OutputPath $subiconPath
			$script:subiconCount++
		}
		# If original image is smaller than anything in the lists (and upscaling is disabled)
		# (`[-1]` gets the last item i the list. The `$null -eq` check is necessary in case the list is empty.)
		elseif (
			($InputSize -lt $sizesPngNormal[-1] -or $null -eq $sizesPngNormal[-1]) -and
			($InputSize -lt $sizesPngSharper[-1] -or $null -eq $sizesPngSharper[-1]) -and
			($InputSize -lt $sizesIcoNormal[-1] -or $null -eq $sizesIcoNormal[-1]) -and
			($InputSize -lt $sizesIcoSharper[-1] -or $null -eq $sizesIcoSharper[-1])) {
			# IMPROVE Make this not run more times than necessary. It currently runs onece for each time time this function is called.
			# This piece of code could be moved to ConvertTo-IcoMultiRes, which would solve the problem, but it feels more organized to keep it here. The performance loss is negligible, unless, perhaps, someone tries to batch convert a huge amount of tiny files, but that's not a very plausible scenario, I think.
			Convert-Image $InputPath $subiconPath
			$script:subiconCount++
			# echo returning
			return
		}

		# Pause
	}
	# Get-Variable -Name subiconCount -Scope Script
	# echo "gccc $subiconCount"

}

function Convert-Image {
	param (
		[Alias("Path")]$InputPath,
		[Alias("Destination")]$OutputPath
	)

	# echo "command: '$magick ""$InputPath"" ""$OutputPath""'"
	& $magick "$InputPath" "$OutputPath"
}

function Convert-ImageResizeNormal {
	param (
		[Alias("Path")][string]$InputPath,
		[Alias("Size")][string]$OutputSize,
		[Alias("Destination")][string]$OutputPath
	)

	# Convert+Resize with Image Magick
	$resize = "$OutputSize" + "x" + "$OutputSize"
	# & $magick "$InputPath" -resize $resize -background none -gravity center -extent $OutputSize "$OutputPath"
	& $magick "$InputPath" `
		-resize $resize `
		"$OutputPath"
}

function Convert-ImageResizeSharper {
	param (
		[Alias("Path")][string]$InputPath,
		[int]$InputSize,
		[int]$OutputSize,
		[Alias("Destination")][string]$OutputPath
	)

	# "Cubic Filters". https://imagemagick.org/Usage/filter/#cubics
	# "Box". https://imagemagick.org/Usage/filter/#box
	# "Scale-Rotate-Translate (SRT) Distortion". https://imagemagick.org/Usage/distorts/#srt
	# "-crop". https://imagemagick.org/script/command-line-options.php#crop
	# "Transpose and Transverse, Diagonally". https://imagemagick.org/Usage/warping/#transpose
	# "Resizing Images". https://legacy.imagemagick.org/Usage/resize/#resize
	
	# Parameters for resizing the two algorithms
	$cubicBValue = 0.0
	$cubicBlurValue = 1.05
	$boxBlurValue = 0.707

	# Initialize variables
	$scaleFactor = 0.0
	$cubicCValue = 0.0
	$useBoxFilter = $false
	$boxFilterSize = 0
	$boxFilterCrop1 = ""
	$boxFilterCrop2 = ""

	# Determine scale factor
	# echo "$OutputSize / $inputSize"
	$scaleFactor = $OutputSize / $inputSize
	# echo $scaleFactor
	if ($scaleFactor -gt 1) {
		# Image enlargement
		# echo enlargement
		$cubicCValue = 1.0
	} elseif ($scaleFactor -ge 0.25 -and $scaleFactor -lt 1) {
		# echo here
		$cubicCValue = 2.6 - (1.6 * $scaleFactor)
	} elseif (<# $scaleFactor -le ? -and  #>$scaleFactor -lt 0.25) {
		# echo box
		$useBoxFilter = $true
		$boxFilterSize = 4 * $OutputSize
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
	} elseif ($scaleFactor -eq 1) {
		Write-Warning "Attempted to resize with scale factor 1.0."
		Pause
	}
	
	# `-resize` keeps aspect ratio by default.
	$resize = "$OutputSize" + "x" + "$OutputSize"

	# echo $OutputPath
	if ($useBoxFilter) {
		# echo box
		& $magick "$InputPath" `
			-filter box -define filter:blur=$boxBlurValue `
			+distort SRT "$boxFilterSrt1" -crop $boxFilterCrop1 `
			+distort SRT "$boxFilterSrt2" -crop $boxFilterCrop2 `
			-transpose `
			-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
			-resize $resize `
			-transpose `
			"$OutputPath"
	} else {
		# echo "command: '$magick ""$InputPath"" -transpose -filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue -resize $resize -transpose ""$OutputPath""'"
		& $magick "$InputPath" `
			-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
			-resize $resize `
			"$OutputPath"
	}
}

function New-MultiResIcon {
	param (
		[Alias("Path")][string]$InputDirectoryPath,
		[Alias("Destination")][string]$OutputPath
	)

	$filesPath = Join-Path -Path $InputDirectoryPath -ChildPath "\*"
	& $magick "$filesPath" "$OutputPath"
}

function Move-Files {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		$Files, # Takes Object[] that represents files
		[string]$Destination
	)

	# Documentation on ShouldProcess:
	# sdwheeler. "Everything You Wanted to Know about ShouldProcess - PowerShell", 17 november 2022. https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-shouldprocess?view=powershell-7.4.

	# Move each file to destination, and prompt on conflict

	# Initialize variables
	$yesToAll = $false
	$noToAll = $false
	$continue = $true

	# echo "files: $Files"

	# Get the number of files
	$fileCount = $Files.Count

	foreach ($f in $Files) {
		# Check if file/directory already exists at destination
		$fileFilename = $f.Name
		# echo "fp: " $f.FullName
		# echo "fn: $fileFilename"
		# echo "destination: $Destination"
		$filePath = $f.FullName
		$fileDestination = Join-Path -Path $Destination -ChildPath $fileFilename
		# Check if destination exists and path type (directory/file) matches
		# Path type matches if both $filePath and $fileDestination are directories, or if both are files
		# If both or either does not exist, it will return $false, so we do not need to check separetly a if a file/directory exists at the destination path
		$destinationExists = (
			([bool](Test-Path -Path $filePath -PathType Container) -and [bool](
				Test-Path -Path $fileDestination -PathType Container)) -or 
			([bool](Test-Path -Path $filePath -PathType Leaf) -and [bool](
				Test-Path $fileDestination -PathType Leaf)))
		# echo "fc $fileCount"
		# If only one file
		if ($destinationExists -and $fileCount -eq 1 -and -not $yesToAll -and -not $noToAll) {
			# IMPROVE Skip if file/directory is identical
			# Do not give the options "Yes to All" and "No to All"
			# (only give options "Yes", "No", "Suspend", and "Help")
			# echo here
			$continue = $PSCmdlet.ShouldContinue(
				"""$fileDestination""",
				'Replace the file in the destination?'
			)
		} elseif ($destinationExists -and -not $yesToAll -and -not $noToAll) {
			# Prompt to replace file
			# (Options: "Yes", "Yes to All", "No", "No to All", "Suspend", and "Help")
			Write-Output ""
			Write-Warning "The destination already has a file named ""$fileFilename""."
			$continue = $PSCmdlet.ShouldContinue(
				"""$fileDestination""",
				'Replace the file in the destination?',
				[ref]$yesToAll,
				[ref]$noToAll
			)
		}
		# If user has selected "Yes to All"
		elseif ($destinationExists -and $yesToAll) {
			Write-Warning "Replacing ""$fileFilename""."
		}
		# If user has selected "No to All"
		elseif ($destinationExists -and $noToAll) {
			Write-Warning "Skipping ""$fileFilename""."
		}

		# echo "continue: $continue"
		# Continue if file does not exist at destination,
		# or if user selected "Yes" or "Yes to All"
		if ($continue) {
			# Move file/directory (replace if exists)
			# echo "fileDestination: $fileDestination"
			# echo "file: $file"
			# echo "will move"
			Move-Item -Path "$filePath" -Destination "$fileDestination" -Force
			$fileCount--
		}
	}
}

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 5. Convert to ICO
#
#########################################################################################

# If first argument is a directory
IF ([bool](Test-Path $argPath -PathType container)) {
	
	Write-Verbose "Directory : $argPath"

	# Name of the directory provided by first argument
	$imagesDirName = (Get-Item -Path $argPath).Name
	
	# Name for $unfinishedDirPath and $finishedDir
	# (GUID is just a precaution and perhaps for easier debugging)
	$tempDirName = "$imagesDirName" + "_" + (New-Guid).Guid

	# Create temporary directory to store icons in before the directory is completed 
	$unfinishedDirPath = Join-Path -Path $unfinishedDirParentPath -ChildPath $tempDirName
	New-Item $unfinishedDirPath -ItemType "directory" 1> $null

	# Get images in argPath
	$dirContentsPath = Join-Path $argPath -ChildPath '*'
	# IMPROVE Add more formats https://imagemagick.org/script/formats.php
	# $images = Get-ChildItem -Path $dirContentsPath
	$images = Get-ChildItem -Path $dirContentsPath `
		-Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg, *.bmp, *.pcx, *.nef, *.ico
	# Iterate through the images
	# Write-Host "file count: $($images.Count)"
	foreach ($image in $images) {
		# FIXME If there two files with different extensions but the same file name, the alphabetically latter file will get used without prompt or warning
		# echo "images $images"

		# Print image file name (with extension)
		# Pause
		$imageName = $image.Name
		Write-Output "- $imageName"

		# Convert image to multi-resolution ICO
		# echo "image: $imageName"
		$dirPath = $unfinishedDirPath
		$imagePath = $image.FullName
		$imageBaseName = (Get-Item -Path $image).BaseName # Name without extension
		$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
		$iconBaseName = $imageBaseName
		# $iconTemp = Join-Path -Path $dirPath -ChildPath "$iconTempBaseName.ico"
		$iconPath = Join-Path -Path $dirPath -ChildPath "$iconBaseName.ico"
		# echo "iconPath $iconPath"
		ConvertTo-IcoMultiRes -InputPath $imagePath -BaseName $iconTempBaseName -OutputPath $iconPath
		# pause
	}

	# Pause
	# Move directory with icons to a directory for finished directories, where icons will be before they are moved to their final destination
	# (This could technically be skipped, but it makes it arguably easier to debug; adds more intuitive file structure)
	# First set the path
	# echo $finishedDirParentPath
	# echo (-not [bool](Test-Path -Path $finishedDirParentPath))
	$finishedDirPath = Join-Path -Path $finishedDirParentPath -ChildPath $tempDirName
	# echo "finishedDir $finishedDirPath"
	# And create parent directory, or else Move-Item will fail (skip if the directory already exists)
	if (-not [bool](Test-Path -Path $finishedDirParentPath)) {
		New-Item $finishedDirParentPath -ItemType "directory" 1> $null
	}
	# Then move the directory
	# echo "unfinishedDir $unfinishedDirPath"
	Move-Item -Path $unfinishedDirPath -Destination $finishedDirPath
	
	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirParentPath

	# Create final destination directory for the icons (unless the directory already exists)
	# echo "iconsFinalDir $iconsFinalDirPath"
	if (-not [bool](Test-Path -Path $iconsFinalDirPath)) {
		New-Item -Path $iconsFinalDirPath -ItemType "directory" 1> $null
	}
	# Pause
	# Move each multi-res icon file, from the temporary completed directory, into $iconsFinalDirPath
	$icons = Get-ChildItem -Path $finishedDirPath
	Move-Files -Files $icons -Destination $iconsFinalDirPath

	# Pause
	# Remove the (now normally empty) temporary completed directory
	# echo "will remove"
	Remove-Item $finishedDirPath -Recurse
	# Remove-Item $finishedDirPath -Recurse -Force
	# TODO Remove each directory individually?

	# Pause
	Remove-Item $tempDirPath -Recurse

	# If first argument is a file
} ELSE {
	Write-Verbose "File : $argPath"

	# Print image file name (with extension)
	$imageFilename = (Get-Item -Path $argPath).Name
	Write-Output "- $imageFilename"

	$unfinishedDirPath = $unfinishedDirParentPath

	# Convert image to multi-resolution ICO
	$imagePath = $argPath
	$dirPath = $unfinishedDirPath
	$imageBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
	$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
	$iconBaseName = $imageBaseName
	$iconPath = Join-Path -Path $dirPath -ChildPath "$iconBaseName.ico"
	ConvertTo-IcoMultiRes -InputPath $imagePath -BaseName $iconTempBaseName -OutputPath $iconPath
	# echo $iconPath

	# Move icon to a temporary directory for finished icon.
	# (This could technically be skipped, but it makes it arguably easier to debug; adds more intuitive file structure)
	$finishedDirPath = $finishedSingleIconDirPath
	$finishedIconPath = Join-Path -Path $finishedDirPath -ChildPath "$iconBaseName.ico"
	# echo "fd $finishedDirPath"
	# echo "fi $finishedIconPath"
	if (-not [bool](Test-Path -Path $finishedDirPath)) {
		New-Item $finishedDirPath -ItemType "directory" 1> $null
	}
	# Pause
	# Move and replace without prompting if existing file found (it would just be confusing for the user if they are got to choose whether to replace a file in a temporary directory that they did not know about.)
	Move-Item -Path $iconPath -Destination $finishedIconPath -Force

	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirParentPath

	# Move icon to its final destination
	$imageDirPath = Convert-Path -Path $((Get-Item $argPath).Directory)
	# echo "im dir $imageDirPath"
	$singleIconFinalDirPath = Join-Path -Path $imageDirPath -ChildPath "$singleIconFinalDirName" # Set full path
	if (-not [bool](Test-Path -Path $singleIconFinalDirPath)) {
		New-Item $singleIconFinalDirPath -ItemType "directory" 1> $null
	}

	# Check if icon already exists
	
	$finishedIcon = Get-Item -Path $finishedIconPath
	Move-Files $finishedIcon $singleIconFinalDirPath

	# Remove the (now normally empty) temporary completed directory
	# echo "will remove"
	Remove-Item $finishedDirPath -Recurse

	# Remove the now empty temporary directory
	Remove-Item $tempDirPath -Recurse

	# Remove-Item $tempDirPath -Recurse
}

# TODO Remove debugging
# TODO Finish code comment documentation
# IMPROVE Error handling
# IMPROVE Would this be an improvement?: Check if file(s) exists in temporary directory when starting, prompt user for action if found
# IMPROVE Get-Help data

#########################################################################################
#
#endregion
#
#########################################################################################