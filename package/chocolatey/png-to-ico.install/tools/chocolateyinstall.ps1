$ErrorActionPreference = 'Stop'
$toolsDir   = "$(Split-Path -parent $MyInvocation.MyCommand.Definition)"
$fileLocation = Join-Path $toolsDir 'png_to_ico_setup.exe'
$packageArgs = @{
  packageName   = $env:ChocolateyPackageName
  unzipLocation = $toolsDir
  fileType      = 'EXE'
  file64         = $fileLocation

  softwareName  = 'png-to-ico'

  checksum64    = 'F86EECEE236CF5E1DDCDD61688334D911394746B9657D6C93546247807716CC5'
  checksumType64= 'sha256'

  silentArgs    = '/S'
  validExitCodes= @(0)
}

Install-ChocolateyInstallPackage @packageArgs
