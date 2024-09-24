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

# Check if RawUI is available before trying to set window properties
if ($null -ne $host.UI.RawUI) {
	# Set console title
	$host.UI.RawUI.WindowTitle = "PNG to ICO"
    
	try {
		# Attempt to change window size
		$host.UI.RawUI.WindowSize = New-Object Management.Automation.Host.Size(65, 30)
	} catch {
		Write-Debug "Window size change is not supported in this environment."
	}

}

Write-Host "`n"
Write-Host " -------------------------------------------------------------"
Write-Host "                          PNG to ICO :"
Write-Host " -------------------------------------------------------------"
Write-Host "`n"

# Script directory path
$scriptDirPath = Split-Path -Parent $MyInvocation.MyCommand.Definition

# First command line argument
$argPath = $args[0]

# If no parameter is specified
if ($null -eq $argPath) {
	# Explain how to use this script
	Write-Host "Usage: png_to_ico.ps1 <file|directory> [<allow upscale>]"
	Write-Host "`nExamples`n"
	Write-Host "png_to_ico.ps1 C:\Users\John\Desktop\my_image.png"
	Write-Host "png_to_ico.ps1 ""C:\Users\John\Desktop\my images"""
	Write-Host "png_to_ico.ps1 ""C:\Users\John\Desktop\my tiny image.png"" true"
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
#region 2. Enum definitions
#
#########################################################################################

enum Algorithm {
	Normal
	Sharper
}

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 3. User configuration
#
#########################################################################################

# Set which subicons each multi-res icon should have to have, based on image size and resize algorithm
$sizesPngNormal = @(256)
$sizesPngSharper = @()
$sizesIcoNormal = @(128)
$sizesIcoSharper = @(96, 64, 48, 32, 24, 16)

# Set to $true to always include the selected subicon sizes (resolutions) in multi-res ICOs even when they exceed the original size of the image
# By default, it will be set to $false unless set to true in command line parameter.
$upscale = $args[1]

# ImageMagick executable
# If the executable is not found here, the script will attempt to use `magick` from the Path environment variable.
$magick = Join-Path $scriptDirPath -ChildPath "ImageMagick\magick.exe"
# Uncomment to set custom path:
# $magick = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"

# (Used for converting all images in a directory)
# Subdirectory, relative to input directory, to save all icons in
# Default: Save into a subdirectory named "ICO"
$iconsFinalDirPath = Join-Path -Path $argPath -ChildPath "ICO"
# Uncomment to set custom path:
# $iconsFinalDirPath = "C:\Users\John\My icons"

# (Used for single file conversion)
# Subdirectory, relative to input image's directory, to save the icon in
# Default: Save icon to the same folder as input image
$singleIconFinalDirName = ""
# Uncomment to set custom path:
# $singleIconFinalDirPath = "C:\Users\John\My icons" 

# Temporary directory path
$tempDirPath = Join-Path -Path $env:TEMP -ChildPath "PNG-to-ICO"

# Temporary directory for subicons
$subiconsDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Unfinished\Subicons"

# (Used for converting all images in a directory)
# Temporary directory for unfinished icon directory, used before all images has been converted
# (Saving all files to one directory and moveing all files at the end, allows the user to overwrite all files with "[A]", in the edge case where there are existing files with the same names and extensions)
# Default: Unfinished icon directories go into "Unfinished" parent directory.
$unfinishedDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Unfinished"

# (Used for single file conversion)
# Temporary directory for finished icon
# Default: Icons go directly in "Finished" directory.
$finishedSingleIconDirPath = Join-Path -Path $tempDirPath -ChildPath "Finished"

# (Used for converting all images in a directory)
# Temporary parent directory for a finished icons subdirectory
# Default: icon directories go into "Finished"
$finishedDirParentPath = Join-Path -Path $tempDirPath -ChildPath "Finished"

# Directories and example files used by the program with default configuration
# [ArgPath]
#   |_ [ICO] (final destination)
#      |_ a.ico
#      |_ b.ico
#   |_ a.png
#   |_ b.png
#   |_ single_icon.ico (final destination)
#   |_ single_icon.png
#
# [Temp]
#   |_ [PNG-to-ICO]
#      |_ [Unfinished]
#         |_ [my_images_<directory guid>]
#            |_ b.ico
#         |_ [Subicons]
#            |_ [a_<image guid>]
#               |_ a_<image guid>-0.png (256x256)
#               |_ a_<image guid>-1.ico (128x128)
#               |_ a_<image guid>-2.ico (96x96)
#               |_ a_<image guid>-3.ico (64x64)
#               |_ a_<image guid>-4.ico (48x48)
#               |_ a_<image guid>-5.ico (32x32)
#               |_ a_<image guid>-6.ico (24x24)
#               |_ a_<image guid>-7.ico (16x16)
#         |_ single_icon_<image guid>.ico
#      |_ [Finished]
#         |_ single_icon_<image guid>.ico
#         |_ [my_images_<directory guid>]
#            |_ a.ico
#            |_ b.ico

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 4. Finalization
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
#region 5. Function definitions
#
#########################################################################################

# Functions are defined in the the order they are first called.
function ConvertTo-IcoMultiRes {
	param (
		# Add aliases that mimics offcical cmdlet parameters,
		# for no reason
		[Parameter(Mandatory = $true)]
		[Alias("Path")][string]$InputPath,
		
		[Parameter(Mandatory = $true)]
		[string]$BaseName,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")][string]$OutputPath
	)

	# Create temporary directory for this icon, for creating subicons in
	$subiconsDirPath = Join-Path -Path $subiconsDirParentPath -ChildPath $BaseName
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

	# Make subicons
	# $subiconCount is suffixed to the subicon file names.
	$script:subiconCount = 0
	# The `ConvertFrom-SubiconsArray` function needs to update the value of $subiconCount. The canonical way to do this is probably:
	# `$subiconCount = ConvertFrom-SubiconsArray <...>`
	# But that would make it look like the purpose of the function is only to update the subiconCount value. 
	# So instead, we set the $subiconCount variable in the Script scope here, and then update it in the other function.
	# (We could also combine the functions, but then we would need to call `magick identify` more times than necessary.)
	ConvertFrom-SubiconsArray -SubiconsArray $sizesPngNormal `
		-InputSize $imageSize `
		-Algorithm Normal `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".png"
	ConvertFrom-SubiconsArray -SubiconsArray $sizesPngSharper `
		-InputSize $imageSize `
		-Algorithm Sharper `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".png"
	ConvertFrom-SubiconsArray -SubiconsArray $sizesIcoNormal `
		-InputSize $imageSize `
		-Algorithm Normal `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".ico"
	ConvertFrom-SubiconsArray -SubiconsArray $sizesIcoSharper `
		-InputSize $imageSize `
		-Algorithm Sharper `
		-InputPath $InputPath `
		-OutputPath $subiconsDirPath `
		-BaseName $BaseName `
		-OutputExt ".ico"

	# Take all the subicons created and assemble into a multi-res icon
	New-MultiResIcon -InputDirectoryPath $subiconsDirPath -OutputPath $OutputPath
	
	# Remove subicons
	Remove-Item -Path $subiconsDirParentPath -Recurse
}

function ConvertFrom-SubiconsArray {
	param (
		[Alias("InputObject")][array]$SubiconsArray,
		
		[Parameter(Mandatory = $true)]
		[Alias("Size")][int]$InputSize,
		
		[Parameter(Mandatory = $true)]
		[Algorithm]$Algorithm,

		[Parameter(Mandatory = $true)]
		[Alias("Path")][string]$InputPath,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")][string]$OutputPath,
		
		[Parameter(Mandatory = $true)]
		[string]$BaseName,

		[Parameter(Mandatory = $true)]
		[Alias("Extension")][string]$OutputExt
	)

	# Get input image extension
	$inputExt = (Get-Item -Path $InputPath).Extension

	# Create the subicons specified in array
	foreach ($subiconSize in $SubiconsArray) {
		$subiconName = "$BaseName-$subiconCount" + $OutputExt
		$subiconPath = Join-Path -Path $OutputPath -ChildPath $subiconName
		# Only resize if original image is greater than $subiconSize
		if ($InputSize -gt $subiconSize) {
			if ($Algorithm -eq [Algorithm]::Sharper) {
				# Create subicon using "Sharper" algorithm
				Convert-ImageResizeSharper -InputPath $InputPath `
					-InputSize $InputSize `
					-OutputSize $subiconSize `
					-OutputPath $subiconPath
				# Increase the subicon count (used for naming subicons).
				$script:subiconCount++
			} elseif ($Algorithm -eq [Algorithm]::Normal) {
				# Create subicon using "Normal" algorithm
				Convert-ImageResizeNormal -InputPath $InputPath `
					-OutputSize $subiconSize `
					-OutputPath $subiconPath
				$script:subiconCount++
			}
		}
		# If original image is equal to $subiconSize, and the format is different, then convert without resize
		elseif (($InputSize -eq $subiconSize) -and ($OutputExt -ne $inputExt)) {
			Convert-Image -InputPath $InputPath `
				-OutputPath $subiconPath
			$script:subiconCount++
		}
		# If original image is equal to $subiconSize, and the format is the same, then simply copy the image to the subicons directory
		elseif (($InputSize -eq $subiconSize) -and ($OutputExt -eq $inputExt)) {
			Copy-Item -Path "$InputPath" `
				-Destination "$subiconPath"
			$script:subiconCount++
		}
		# If upscaling is enabled, and the original image is smaller than $subiconSize, then resize
		elseif (($upscale -eq $true) -and ($InputSize -lt $subiconSize)) {
			Convert-ImageResizeNormal -InputPath $InputPath `
				-OutputSize $subiconSize `
				-OutputPath $subiconPath
			$script:subiconCount++
		}
		# If original image is smaller than anything in the lists (and upscaling is disabled), then convert without resize
		# (`[-1]` gets the last item i the list.
		# The `$null -eq` check is necessary in case the list is empty.)
		elseif (
			($InputSize -lt $sizesPngNormal[-1] -or $null -eq $sizesPngNormal[-1]) -and
			($InputSize -lt $sizesPngSharper[-1] -or $null -eq $sizesPngSharper[-1]) -and
			($InputSize -lt $sizesIcoNormal[-1] -or $null -eq $sizesIcoNormal[-1]) -and
			($InputSize -lt $sizesIcoSharper[-1] -or $null -eq $sizesIcoSharper[-1])) {
			# IMPROVE Make this not run more times than necessary. It currently runs onece for each time time this function is called.
			# This piece of code could be moved to ConvertTo-IcoMultiRes, which would solve the problem, but it feels more organized to keep it here. The performance loss is negligible, unless, perhaps, someone tries to batch convert a huge amount of tiny files, but that's not a very plausible scenario, I think.
			Convert-Image -InputPath $InputPath `
				-OutputPath $subiconPath
			$script:subiconCount++
			return
		}
	}
}

function Convert-Image {
	param (
		[Parameter(Mandatory = $true)]
		[Alias("Path")]$InputPath,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")]$OutputPath
	)

	# Convert with Image Magick
	& $magick "$InputPath" "$OutputPath"
}

function Convert-ImageResizeNormal {
	param (
		[Parameter(Mandatory = $true)]
		[Alias("Path")][string]$InputPath,
		
		[Parameter(Mandatory = $true)]
		[Alias("Size")][string]$OutputSize,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")][string]$OutputPath
	)

	# Convert + resize with Image Magick
	$resize = "$OutputSize" + "x" + "$OutputSize"
	& $magick "$InputPath" `
		-resize $resize `
		"$OutputPath"
}

function Convert-ImageResizeSharper {
	param (
		[Parameter(Mandatory = $true)]
		[Alias("Path")][string]$InputPath,
		
		[Parameter(Mandatory = $true)]
		[int]$InputSize,
		
		[Parameter(Mandatory = $true)]
		[int]$OutputSize,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")][string]$OutputPath
	)

	# "Cubic Filters". https://imagemagick.org/Usage/filter/#cubics
	# "Box". https://imagemagick.org/Usage/filter/#box
	# "Scale-Rotate-Translate (SRT) Distortion". https://imagemagick.org/Usage/distorts/#srt
	# "-crop". https://imagemagick.org/script/command-line-options.php#crop
	# "Transpose and Transverse, Diagonally". https://imagemagick.org/Usage/warping/#transpose
	# "Resizing Images". https://legacy.imagemagick.org/Usage/resize/#resize

	# Parameters for the two resizing algorithms
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
	$scaleFactor = $OutputSize / $inputSize

	# Set parameters based on scale factor
	if ($scaleFactor -gt 1) {
		# Image enlargement
		$cubicCValue = 1.0
	} elseif ($scaleFactor -ge 0.25 -and $scaleFactor -lt 1) {
		$cubicCValue = 2.6 - (1.6 * $scaleFactor)
	} elseif (<# $scaleFactor -le ? -and  #>$scaleFactor -lt 0.25) {
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
	
	# Set -resize parameter
	# `-resize` keeps aspect ratio by default.
	$resize = "$OutputSize" + "x" + "$OutputSize"

	# Convert with Image Magick
	if ($useBoxFilter) {
		# Convert + resize with box filter and cubic filter
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
		# Convert + resize with cubic filter
		& $magick "$InputPath" `
			-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
			-resize $resize `
			"$OutputPath"
	}
}

function New-MultiResIcon {
	param (
		[Parameter(Mandatory = $true)]
		[Alias("Path")][string]$InputDirectoryPath,
		
		[Parameter(Mandatory = $true)]
		[Alias("Destination")][string]$OutputPath
	)

	# Assemble subicons into a multi-res ICO with Image Magick
	$filesPath = Join-Path -Path $InputDirectoryPath -ChildPath "\*"
	& $magick "$filesPath" "$OutputPath"
}

function Move-Files {
	[CmdletBinding(SupportsShouldProcess)]
	param(
		$Files, # Takes Object[] that represents files
		
		[Parameter(Mandatory = $true)]
		[string]$Destination
	)

	# Documentation on ShouldProcess:
	# sdwheeler. "Everything You Wanted to Know about ShouldProcess - PowerShell", 17 november 2022. https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-shouldprocess?view=powershell-7.4.

	# Initialize variables
	$yesToAll = $false
	$noToAll = $false
	$continue = $true

	# Get the number of files
	$fileCount = $Files.Count

	# Move each file to destination, and prompt on conflict
	foreach ($f in $Files) {
		$fileFilename = $f.Name
		$filePath = $f.FullName
		$fileDestination = Join-Path -Path $Destination -ChildPath $fileFilename
		# Check if destination exists and path type (directory/file) matches
		# (Path type matches if both $filePath and $fileDestination are directories, or if both are files;
		# If both or either does not exist, it will return $false, so we do not need to check separetly a if a file/directory exists at the destination path)
		$destinationExists = (
			([bool](Test-Path -Path $filePath -PathType Container) -and [bool](
				Test-Path -Path $fileDestination -PathType Container)) -or 
			([bool](Test-Path -Path $filePath -PathType Leaf) -and [bool](
				Test-Path $fileDestination -PathType Leaf)))
		# If conflict and only one file
		if ($destinationExists -and $fileCount -eq 1 -and -not $yesToAll -and -not $noToAll) {
			# Do not give the options "Yes to All" and "No to All"
			# (only give options "Yes", "No", "Suspend", and "Help")
			$continue = $PSCmdlet.ShouldContinue(
				"""$fileDestination""",
				'Replace the file in the destination?'
			)
			# IMPROVE Skip if folder/directory at path and destination are identical (if they have the same hash)
			# If conflict
		} elseif ($destinationExists -and -not $yesToAll -and -not $noToAll) {
			# Prompt to replace file
			# (Options: "Yes", "Yes to All", "No", "No to All", "Suspend", and "Help")
			Write-Host ""
			Write-Warning "The destination already has a file named ""$fileFilename""."
			$continue = $PSCmdlet.ShouldContinue(
				"""$fileDestination""",
				'Replace the file in the destination?',
				[ref]$yesToAll,
				[ref]$noToAll
			)
		}
		# If conflict and the user has selected "Yes to All"
		elseif ($destinationExists -and $yesToAll) {
			Write-Warning "Replacing ""$fileFilename""."
		}
		# If conflict and the user has selected "No to All"
		elseif ($destinationExists -and $noToAll) {
			Write-Warning "Skipping ""$fileFilename""."
		}

		# Continue if file does not exist at destination,
		# or if user selected "Yes" or "Yes to All"
		if ($continue) {
			# Move file/directory (replace if exists)
			Move-Item -Path "$filePath" -Destination "$fileDestination" -Force
		}
		# Reduce file count (number of files in queue)
		$fileCount--
	}
}

#########################################################################################
#
#endregion
#
#########################################################################################

#########################################################################################
#
#region 6. Convert to ICO
#
#########################################################################################

# If first argument is a directory
if ([bool](Test-Path $argPath -PathType container)) {
	
	Write-Verbose "Directory : $argPath"

	# Name of the directory provided by first argument
	$imagesDirName = (Get-Item -Path $argPath).Name
	
	# Name for $unfinishedDirPath and $finishedDir, where files will be temporarily stored
	# (GUID is just a precaution and perhaps for easier debugging)
	$tempDirName = "$imagesDirName" + "_" + (New-Guid).Guid

	# Create temporary directory to store icons in before the directory is completed (before all ICOs have been created)
	$unfinishedDirPath = Join-Path -Path $unfinishedDirParentPath -ChildPath $tempDirName
	New-Item $unfinishedDirPath -ItemType "directory" 1> $null

	# Get images in argPath
	$dirContentsPath = Join-Path $argPath -ChildPath '*'
	# IMPROVE Add more formats https://imagemagick.org/script/formats.php
	# IMPROVE Do not convert manually created multi-res icons into sloppy automated multi-res icons 
	$images = Get-ChildItem -Path $dirContentsPath `
		-Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg, *.bmp, *.pcx, *.nef, *.ico

	# Go image by image
	foreach ($image in $images) {
		# FIXME If there two files with different extensions but the same file name, the alphabetically latter file will get used without prompt or warning

		# Print image file name (with extension)
		$imageName = $image.Name
		Write-Host "- $imageName"

		# Convert image to multi-resolution ICO
		$dirPath = $unfinishedDirPath
		$imagePath = $image.FullName
		$imageBaseName = (Get-Item -Path $image).BaseName # Name without extension
		$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
		$iconBaseName = $imageBaseName
		$iconPath = Join-Path -Path $dirPath -ChildPath "$iconBaseName.ico"
		ConvertTo-IcoMultiRes -InputPath $imagePath `
			-BaseName $iconTempBaseName `
			-OutputPath $iconPath
	}

	# Move directory with icons to a directory for finished directory, where icons will be stored before they are moved to their final destination
	# (This could technically be skipped, but it arguably gives a more intuitive file structure, and makes it arguably easier to debug, as you will be able to see, just based on the file structure, where in the machinery something went wrong.)
	# First set the path
	$finishedDirPath = Join-Path -Path $finishedDirParentPath -ChildPath $tempDirName
	# And create parent directory, or else Move-Item will fail (skip if the directory already exists)
	if (-not [bool](Test-Path -Path $finishedDirParentPath)) {
		New-Item $finishedDirParentPath -ItemType "directory" 1> $null
	}
	# Then move the directory
	Move-Item -Path $unfinishedDirPath -Destination $finishedDirPath
	
	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirParentPath

	# Create final destination directory for the icons (unless the directory already exists)
	if (-not [bool](Test-Path -Path $iconsFinalDirPath)) {
		New-Item -Path $iconsFinalDirPath -ItemType "directory" 1> $null
	}
	# Move each multi-res icon file, from the temporary completed directory, into $iconsFinalDirPath
	$icons = Get-ChildItem -Path $finishedDirPath
	Move-Files -Files $icons -Destination $iconsFinalDirPath

	# Remove the (now normally empty) temporary completed directory
	Remove-Item $finishedDirPath -Recurse
	
	# Remove the (now normally empty) temporary parent directory for the unfinished data directory
	# Remove-Item $finishedDirParentPath -Recurse

	Remove-Item $tempDirPath -Recurse
	# If first argument is a file
} else {
	Write-Verbose "File : $argPath"

	# Print image file name (with extension)
	$imageFilename = (Get-Item -Path $argPath).Name
	Write-Host "- $imageFilename"

	# Create temporary directory for the icon
	$unfinishedDirPath = $unfinishedDirParentPath
	New-Item $unfinishedDirPath -ItemType "directory" 1> $null

	# Convert image to multi-resolution ICO
	$imagePath = $argPath
	$dirPath = $unfinishedDirPath
	$imageBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
	$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
	$iconBaseName = $imageBaseName
	$iconPath = Join-Path -Path $dirPath -ChildPath "$iconBaseName.ico"
	ConvertTo-IcoMultiRes -InputPath $imagePath `
		-BaseName $iconTempBaseName `
		-OutputPath $iconPath

	# Move icon to a temporary directory for finished icon.
	# (This could technically be skipped, but it arguably gives a more intuitive file structure, and makes it arguably easier to debug, as you will be able to see, just based on the file structure, where in the machinery something went wrong.)
	$finishedDirPath = $finishedSingleIconDirPath
	$finishedIconPath = Join-Path -Path $finishedDirPath -ChildPath "$iconBaseName.ico"
	if (-not [bool](Test-Path -Path $finishedDirPath)) {
		New-Item $finishedDirPath -ItemType "directory" 1> $null
	}
	# Move and replace without prompting if existing file found
	# (it would just be confusing for the user if they are got to choose whether to replace a file in a temporary directory that they did not know about.)
	Move-Item -Path $iconPath -Destination $finishedIconPath -Force

	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirPath

	# Remove the parent directory for the unfinished data directory
	# Remove-Item $unfinishedDirParentPath

	# Move icon to its final destination
	$imageDirPath = Convert-Path -Path $((Get-Item $argPath).Directory)
	$singleIconFinalDirPath = Join-Path -Path $imageDirPath -ChildPath "$singleIconFinalDirName" # Set full path
	if (-not [bool](Test-Path -Path $singleIconFinalDirPath)) {
		New-Item $singleIconFinalDirPath -ItemType "directory" 1> $null
	}
	
	$finishedIcon = Get-Item -Path $finishedIconPath
	Move-Files -Files $finishedIcon -Destination $singleIconFinalDirPath

	# Remove the (now normally empty) temporary completed directory
	Remove-Item $finishedDirPath -Recurse
	
	# Remove the (now normally empty) temporary parent directory for the completed directory
	# Remove-Item $finishedSingleIconDirPath -Recurse

	# Remove the now empty temporary directory
	Remove-Item $tempDirPath -Recurse
}

# IMPROVE Error handling
# IMPROVE Would this be an improvement?: Check if file(s) exists in temporary directory when starting, prompt user for action if found
# IMPROVE Get-Help data

#########################################################################################
#
#endregion
#
#########################################################################################