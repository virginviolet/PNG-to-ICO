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
	Set-Variable -Name outputCount -Value 0 -Scope Script
	# $outputFiles = @()

	if ($height -ge $width) {
		$largestDimension = $height
	}
	else {
		$largestDimension = $width
	}
	# echo $largestDimension
	# if ($largestDimension -le $sizes_) {}
	# echo $fileBaseName
	# The canonical way to do this is `$outputCount = ConvertFrom-SizeList <...>`, but that would make it look misleading in this case, imo
	ConvertFrom-SizeList $sizesPngNormal $largestDimension "normal" $image $iconTempDir $outputCount $fileBaseName "png"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesPngSharper $largestDimension "sharper" $image $iconTempDir $outputCount $fileBaseName "png"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesIcoNormal $largestDimension "normal" $image $iconTempDir $outputCount $fileBaseName "ico"
	Get-Variable -Name outputCount -Scope Script 1> $null
	ConvertFrom-SizeList $sizesIcoSharper $largestDimension "sharper" $image $iconTempDir $outputCount $fileBaseName "ico"
	Get-Variable -Name outputCount -Scope Script 1> $null
}

function ConvertFrom-SizeList ($sizeList, $dimension, $algorithm, $inputImage, $outputDir, $outputCount, $outputBaseName, $outputExt) {
	
	$outputCountLocal = $outputCount

	# Initialize variables
	$iconSizedName = ""
	$iconSized = ""

	foreach ($size in $sizeList) {
		$outputExt = "png"
		# is less than or equal
		# echo "l $outputCountLocal"
		if ($size -le $dimension) {
			$iconSizedName = "$outputBaseName-$outputCountLocal.$outputExt"
			# echo "'$outputBaseName'"
			# echo "'$iconSizedName'"
			# echo $size
			$iconSized = Join-Path -Path $iconTempDir -ChildPath $iconSizedName
			if ($algorithm -eq "sharper") {
				Convert-ImageNormal $inputImage $size $iconSized
			}
			elseif ($algorithm -eq "normal") {
				Convert-ImageSharper $inputImage $size $iconSized $dimension
			} 
			else {
				Write-Error "Incorrect algorithm '$algorithm'"
			}
			$outputCountLocal++
			# Get-Variable -Name outputCount -Scope Script
			# echo "l $outputCountLocal"
		}
		# TODO Only resize if less than. If equal, only convert if not the correct format. I need to get input ext for that.
	}
	Set-Variable -Name outputCount -Value $outputCountLocal -Scope Script
	# Get-Variable -Name outputCount -Scope Script
	# echo "gccc $outputCount"
}

function Convert-ImageNormal ($image, $outputSize, $outputFile) {
	& $magick $image -resize $size -background none -gravity center -extent $outputSize $outputFile
}

function Convert-ImageSharper ($image, $outputSize, $outputFile, $inputSize) {
	
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

	# XXX
	# Determine scale factor
	$scaleFactor = $outputSize / $inputSize
	# echo $scaleFactor
	if ($scaleFactor -gt 1) {
		# Image enlargement
		# TODO Test if ok to use floating point 1.0 instead of just 1
		$cubicCValue = 1.0
	}
	elseif ($scaleFactor -ge 0.25 -and $scaleFactor -lt 1) {
		$cubicCValue = 2.6 - (1.6 * $scaleFactor)
	}
	elseif (<# $scaleFactor -le ? -and  #>$scaleFactor -lt 0.25) {
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
		$boxFilterParameters += "-filter box -define filter:blur=$boxBlurValue"
		$boxFilterParameters += " +distort SRT ""$boxFilterSrt1"" -crop $boxFilterCrop1"
		$boxFilterParameters += " +distort SRT ""$boxFilterSrt2"" -crop $boxFilterCrop2"
		# echo "parameters: '$boxFilterParameters'"
	}
	elseif ($scaleFactor -eq 1) {
		# TODO do not use cubic filter?
		Write-Warning "Scale factor 1. Feature TBA."
	}

	# "Cubic Filters". https://imagemagick.org/Usage/filter/#cubics
	# "Box". https://imagemagick.org/Usage/filter/#box
	# "Scale-Rotate-Translate (SRT) Distortion". https://imagemagick.org/Usage/distorts/#srt
	# "-crop". https://imagemagick.org/script/command-line-options.php#crop
	# "Transpose and Transverse, Diagonally". https://imagemagick.org/Usage/warping/#transpose
	# "Resizing Images". https://legacy.imagemagick.org/Usage/resize/#resize
	# `-resize` keeps aspect ratio by default.
	$resize = "$outputSize" + "x" + "$outputSize"
	$parameters += """$image"""
	$parameters += " $boxFilterParameters"
	$parameters += " -transpose"
	$parameters += " -filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue"
	$parameters += " -resize $resize"
	$parameters += " -transpose"
	$parameters += " ""$outputFile"""
	# echo "command: '$magick $parameters'"
	echo "command: '$magick $parameters'"
	# echo $image
	# & $magick $parameters # This doesn't work for some reason
	& $magick "$image" -filter box -define filter:blur=$boxBlurValue +distort SRT "$boxFilterSrt1" -crop $boxFilterCrop1 +distort SRT "$boxFilterSrt2" -crop $boxFilterCrop2 -transpose -filter cubic -define filter:b=$cubicBValue -define filter:c=$cubicCValue -define filter:blur=$cubicBlurValue -resize $resize -transpose "$outputFile"
	pause
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