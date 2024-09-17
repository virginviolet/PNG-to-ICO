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
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

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
	# Turn the string into a path
	$argPath = Resolve-Path -Path $argPath
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
$magick = Join-Path $scriptDir -ChildPath "ImageMagick\magick.exe"
# Uncomment to set custom path:
# $magick = "C:\Program Files\ImageMagick-7.1.1-Q16-HDRI\magick.exe"

# Used for converting all images in a directory:
# Subdirectory, relative to input directory, to save all icons in
# Default: Save into a subdirectory named "ICO"
$iconsFinalDir = Join-Path -Path $argPath -ChildPath "ICO"
# echo "argPath $argPath"
# echo "iconsFinalDir $iconsFinalDir"
# Uncomment to set custom path:
# $iconsFinalDir = "C:\Users\John\My icons"

# Used for single file conversion:
# Subdirectory, relative to input image's directory, to save the icon in
# Default: Save icon to the same folder as input image
$singleIconFinalDirName = ""
# Uncomment to set custom path:
# $singleIconFinalDir = "C:\Users\John\My icons" 

# Temporary directory path
$tempDir = Join-Path -Path $env:TEMP -ChildPath "PNG-to-ICO"

$subiconsDirParent = Join-Path -Path $tempDir -ChildPath "Unfinished\Subicons"

# Temporary directory for unfinished icon directories, used before all images has been converted (only for when first argument is a directory)
# (Saving all files to one directory and move all files at the end, allows for the user to overwrite all files with "[A]", in the edge case where there are existing files with the same names and extensions)
# Default: Unfinished icon directories go into "Unfinished" parent directory.
$unfinishedDirParent = Join-Path -Path $tempDir -ChildPath "Unfinished"

# Temporary directory for finished icons (when first argument is a file)
# Default: Icons go directly in "Finished" directory.
$finishedSingleIconDir = Join-Path -Path $tempDir -ChildPath "Finished"

# Temporary parent directory for a finished icons subdirectory (when first argument is a directory)
# Default: icon directories go into "Finished"
$finishedDirParent = Join-Path -Path $tempDir -ChildPath "Finished"

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
		$image,
		$iconBaseName,
		$icon
	)

	$subiconsDir = Join-Path -Path $subiconsDirParent -ChildPath $iconBaseName
	
	# Create temporary directory for this icon, for storing subicons
	if (-Not [bool](Test-Path -Path $subiconsDir)) {
		New-Item -Path $subiconsDir -ItemType "directory" 1> $null
	}

	# Get width and height of input image
	$width = [int](& $magick identify -ping -format '%w' $image)
	$height = [int](& $magick identify -ping -format '%h' $image)

	# Only the largest dimension is needed
	if ($height -ge $width) {
		$imageSize = $height
	} else {
		$imageSize = $width
	}

	# echo $imageSize
	# echo $iconBaseName

	# Make subicons
	# We need the function to update the value of $subiconCount. The canonical way to do this is probably:
	# `$subiconCount = ConvertFrom-SizeList <...>`
	# But that would make it look like the purpose of that function is only to update the subiconCount value. So instead, we set variable in the Script scope here, and then update it in the other function.
	# $subiconCount is suffixed to the subicon file names.
	$script:subiconCount = 0 # Initialize variable
	ConvertFrom-SizeList $sizesPngNormal $imageSize "normal" $image $subiconsDir $iconBaseName "png"
	# Pause
	ConvertFrom-SizeList $sizesPngSharper $imageSize "sharper" $image $subiconsDir $iconBaseName "png"
	# Pause
	ConvertFrom-SizeList $sizesIcoNormal $imageSize "normal" $image $subiconsDir $iconBaseName "ico"
	ConvertFrom-SizeList $sizesIcoSharper $imageSize "sharper" $image $subiconsDir $iconBaseName "ico"
	
	New-MultiResIcon $subiconsDir $iconBaseName $icon
	# Pause
	# Take all the subicons created and assemble into a multi-res icon
	Remove-Item $subiconsDirParent -Recurse
}

function ConvertFrom-SizeList {
	param (
		$sizeList,
		$imageSize,
		$algorithm,
		$inputImage,
		$outputDir,
		$outputBaseName,
		$outputExt
	)

	# Get input image extension
	$inputExt = (Get-ChildItem -Path $inputImage).Extension

	# echo "$imageSize $sizesPngNormal[-1] $sizesPngSharper[-1] $sizesIcoNormal[-1]) $sizesIcoSharper[-1]"
	# echo $imageSize
	# echo $sizesIcoSharper[-1]
	foreach ($size in $sizeList) {
		# echo "SIC $subiconCount"
		# Only resize if original image is greater than $size
		$subiconName = "$outputBaseName-$subiconCount.$outputExt"
		# echo "'$outputBaseName'"
		# echo "'$subiconName'"
		# echo $size
		$subicon = Join-Path -Path $outputDir -ChildPath $subiconName
		if ($imageSize -gt $size) {
			# echo one
			if ($algorithm -eq "sharper") {
				# echo "$size sharp"
				Convert-ImageResizeSharper $inputImage $size $subicon $imageSize
			} elseif ($algorithm -eq "normal") {
				# echo "$size normal"
				Convert-ImageResizeNormal $inputImage $size $subicon
			} else {
				Write-Error "Incorrect algorithm '$algorithm'"
			}
			$script:subiconCount++
			# Get-Variable -Name subiconCount -Scope Script
			# echo "SIC $subiconCount"
		}
		# If original image is equal to $size, and the format is different, then convert without resize
		elseif (($imageSize -eq $size) -and ($outputExt -ne $inputExt)) {
			# echo two
			Convert-Image $inputImage $subicon
			$script:subiconCount++
		}
		# If original image is equal to $size, and the format is the same, then simply copy the image to the subicons directory
		elseif (($imageSize -eq $size) -and ($outputExt -eq $inputExt)) {
			# echo three
			Copy-Item -Path "$inputImage" -Destination "$subicon"
			$script:subiconCount++
		}
		# If upscaling is enabled, and the original image is smaller than $size
		elseif (($upscale -eq $true) -and ($imageSize -lt $size)) {
			Convert-ImageResizeNormal $inputImage $size $subicon
			$script:subiconCount++
		}
		# If original image is smaller than anything in the lists (and upscaling is disabled)
		# (`[-1]` gets the last item i the list. The `$null -eq` check is necessary in case the list is empty.)
		elseif (
			($imageSize -lt $sizesPngNormal[-1] -or $null -eq $sizesPngNormal[-1]) -and
			($imageSize -lt $sizesPngSharper[-1] -or $null -eq $sizesPngSharper[-1]) -and
			($imageSize -lt $sizesIcoNormal[-1] -or $null -eq $sizesIcoNormal[-1]) -and
			($imageSize -lt $sizesIcoSharper[-1] -or $null -eq $sizesIcoSharper[-1])) {
			# IMPROVE Make this not run more times than necessary. It currently runs onece for each time time this function is called.
			# This piece of code could be moved to ConvertTo-IcoMultiRes, which would solve the problem, but it feels more organized to keep it here. The performance loss is negligible, unless, perhaps, someone tries to batch convert a huge amount of tiny files, but that's not a very plausible scenario, I think.
			Convert-Image $inputImage $subicon
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
		$inputSize
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
	# echo "$outputSize / $inputSize"
	$scaleFactor = $outputSize / $inputSize
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
	} elseif ($scaleFactor -eq 1) {
		Write-Warning "Attempted to resize with scale factor 1.0."
		Pause
	}
	
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
	} else {
		# echo "command: '$magick ""$inputFile"" -transpose -filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue -resize $resize -transpose ""$outputFile""'"
		& $magick "$inputFile" `
			-filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue `
			-resize $resize `
			"$outputFile"
	}
}

function New-MultiResIcon {
	param (
		$inputDir,
		$tempName,
		$outputIcon
	)

	$files = Join-Path -Path $inputDir -ChildPath "\*"
	& $magick "$files" "$outputIcon"
}

function Move-Files {
	[CmdletBinding(SupportsShouldProcess)]
	param($files,
		$destination
	)

	# Documentation on ShouldProcess:
	# sdwheeler. "Everything You Wanted to Know about ShouldProcess - PowerShell", 17 november 2022. https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-shouldprocess?view=powershell-7.4.

	# Move each file to destination, and prompt on conflict

	# Initialize variables
	$yesToAll = $false
	$noToAll = $false
	$continue = $true

	# echo "files: $files"

	# Get the number of files
	$fileCount = $files.Count

	foreach ($f in $files) {
		# Check if file/directory already exists at destination
		$fileFilename = $f.Name
		# echo "fp: " $f.FullName
		# echo "fn: $fileFilename"
		# echo "destination: $destination"
		$filePath = $f.FullName
		$fileDestination = Join-Path -Path $destination -ChildPath $fileFilename
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
			# Move item (replace if exists)
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
	
	# Name for $unfinishedDir and $finishedDir
	# (GUID is just a precaution and perhaps for easier debugging)
	$tempDirName = "$imagesDirName" + "_" + (New-Guid).Guid

	# Create temporary directory to store icons in before the directory is completed 
	$unfinishedDir = Join-Path -Path $unfinishedDirParent -ChildPath $tempDirName
	New-Item $unfinishedDir -ItemType "directory" 1> $null

	# Get images in argPath
	$dirContents = Join-Path $argPath -ChildPath '*'
	# TODO Add more formats? .ico?
	# $images = Get-ChildItem -Path $dirContents
	$images = Get-ChildItem -Path $dirContents -Include *.png, *.bmp, *.gif, *.jpg, *.jpeg, *.svg
	# Iterate through the images
	# Write-Host "file count: $($images.Count)"
	foreach ($i in $images) {
		# FIXME If there two files with different extensions but the same file name, the alphabetically latter file will get used without prompt or warning
		# echo "images $images"

		# Print image file name (with extension)
		$imageName = $i.Name
		Write-Output "- $imageName"

		# Convert image to multi-resolution ICO
		$image = $i
		# echo "image: $imageName"
		$dir = $unfinishedDir
		$imageBaseName = (Get-Item -Path $i).BaseName # Name without extension
		$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
		$iconBaseName = $imageBaseName
		# $iconTemp = Join-Path -Path $dir -ChildPath "$iconTempBaseName.ico"
		$icon = Join-Path -Path $dir -ChildPath "$iconBaseName.ico"
		# echo "icon $icon"
		ConvertTo-IcoMultiRes $image $iconTempBaseName $icon
		# pause
	}

	# Pause
	# Move directory with icons to a directory for finished directories, where icons will be before they are moved to their final destination
	# (This could technically be skipped, but it makes it arguably easier to debug; adds more intuitive file structure)
	# First set the path
	# echo $finishedDirParent
	# echo (-not [bool](Test-Path -Path $finishedDirParent))
	$finishedDir = Join-Path -Path $finishedDirParent -ChildPath $tempDirName
	# echo "finishedDir $finishedDir"
	# And create parent directory, or else Move-Item will fail (skip if the directory already exists)
	if (-not [bool](Test-Path -Path $finishedDirParent)) {
		New-Item $finishedDirParent -ItemType "directory" 1> $null
	}
	# Then move the directory
	# echo "unfinishedDir $unfinishedDir"
	Move-Item -Path $unfinishedDir -Destination $finishedDir
	
	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirParent

	# Create final destination directory for the icons (unless the directory already exists)
	# echo "iconsFinalDir $iconsFinalDir"
	if (-not [bool](Test-Path -Path $iconsFinalDir)) {
		New-Item -Path $iconsFinalDir -ItemType "directory" 1> $null
	}
	# Pause
	# Move each multi-res icon file, from the temporary completed directory, into $iconsFinalDir
	$icons = Get-ChildItem -Path $finishedDir
	Move-Files $icons $iconsFinalDir

	# Pause
	# Remove the (now normally empty) temporary completed directory
	# echo "will remove"
	# IMPROVE Prompt if user wants to permanently remove files remaining in $finishedDir (if they selected to not overwrite files for example, there might be files remaining in $finishedDir) EDIT: Or not.
	Remove-Item $finishedDir -Recurse
	# Remove-Item $finishedDir -Recurse -Force
	# TODO Remove each directory individually?

	# Pause
	Remove-Item $tempDir -Recurse

	# If first argument is a file
} ELSE {
	Write-Verbose "File : $argPath"

	# Print image file name (with extension)
	$imageFilename = (Get-Item -Path $argPath).Name
	Write-Output "- $imageFilename"

	$unfinishedDir = $unfinishedDirParent

	# Convert image to multi-resolution ICO
	$image = $argPath
	$dir = $unfinishedDir
	$imageBaseName = (Get-Item -Path $argPath).BaseName # Name without extension
	$iconTempBaseName = $imageBaseName + "_" + (New-Guid).Guid
	$iconTemp = Join-Path -Path $dir -ChildPath "$iconTempBaseName.ico"
	$iconBaseName = $imageBaseName
	$icon = Join-Path -Path $dir -ChildPath "$iconBaseName.ico"
	ConvertTo-IcoMultiRes $image $iconTempBaseName $icon
	# echo $iconTemp

	# Move icon to a temporary directory for finished icon.
	# (This could technically be skipped, but it makes it arguably easier to debug; adds more intuitive file structure)
	$finishedDir = $finishedSingleIconDir
	$finishedIcon = Join-Path -Path $finishedDir -ChildPath "$iconBaseName.ico"
	# echo "fd $finishedDir"
	# echo "fi $finishedIcon"
	if (-not [bool](Test-Path -Path $finishedDir)) {
		New-Item $finishedDir -ItemType "directory" 1> $null
	}
	# Pause
	# Move and replace without prompting if existing file found (it would just be confusing for the user if they are got to choose whether to replace a file in a temporary directory that they did not know about.)
	Move-Item -Path $icon -Destination $finishedIcon -Force

	# Remove the now empty directory where unfinished data was stored
	Remove-Item $unfinishedDirParent

	# Move icon to its final destination
	$imageDir = Resolve-Path -Path $((Get-Item $argPath).Directory)
	# echo "im dir $imageDir"
	$singleIconFinalDir = Join-Path -Path $imageDir -ChildPath "$singleIconFinalDirName" # Set full path
	if (-not [bool](Test-Path -Path $singleIconFinalDir)) {
		New-Item $singleIconFinalDir -ItemType "directory" 1> $null
	}

	# Check if icon already exists
	
	# TODO Better differentiate varibles that are paths vs objects
	$finishedIconItem = Get-Item -Path $finishedIcon
	Move-Files $finishedIconItem $singleIconFinalDir

	# Remove the (now normally empty) temporary completed directory
	# echo "will remove"
	Remove-Item $finishedDir -Recurse

	# Remove the now empty temporary directory
	Remove-Item $tempDir -Recurse

	# Remove-Item $tempDir -Recurse
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