$ErrorActionPreference = 'Continue'

# Let's close the .bat file if it is running (edge case, unconfirmed if it would actually
# interfere with installation)

# Name or path of .bat file to terminate
$targetBatFile = "png_to_ico.bat"

# Get all cmd processes
$cmdProcesses = Get-WmiObject -Query "SELECT * FROM Win32_Process WHERE Name = 'cmd.exe'" -ea 0

foreach ($process in $cmdProcesses) {
    # Get the command line arguments of the process
    $commandLine = $process.CommandLine

    # Check if the command line contains the target .bat file
    if ($commandLine -like "*$targetBatFile*") {
        Write-Host "Terminating process $($process.ProcessId) running $targetBatFile"
        # Terminate the process
        Stop-Process -Id $process.ProcessId -Force -ea 0
    }
}
